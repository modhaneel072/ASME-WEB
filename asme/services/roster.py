"""Roster import (PDF -> accounts) and credential exports."""

from __future__ import annotations

import csv
import io
from datetime import datetime
from pathlib import Path

from flask import current_app
from sqlalchemy import func
from werkzeug.security import generate_password_hash

from asme import events
from asme.auth.session import normalize_role
from asme.constants import EMAIL_RE, ROSTER_BOOL_WORDS
from asme.extensions import db
from asme.models import Member, User
from asme.services import identity
from asme.utils import normalize_text_key, split_name_parts

try:  # pragma: no cover - optional dependency fallback for local dev
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None


def roster_credentials_dir() -> Path:
    path = Path(current_app.instance_path) / "roster_credentials"
    path.mkdir(parents=True, exist_ok=True)
    return path


def parse_roster_pdf_entries(pdf_bytes):
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed. Add pypdf to requirements and redeploy.")
    reader = PdfReader(io.BytesIO(pdf_bytes))
    rows = []
    seen = set()
    for page in reader.pages:
        text_blob = page.extract_text() or ""
        for raw_line in text_blob.splitlines():
            line = " ".join(raw_line.split()).strip()
            if not line:
                continue
            lowered = line.lower()
            if lowered.startswith("first name last name"):
                continue
            if "engage general teams project teams" in lowered:
                continue
            tokens = line.split(" ")
            email = ""
            for idx in range(len(tokens) - 1, -1, -1):
                maybe_email = tokens[idx].strip(";,")
                if EMAIL_RE.match(maybe_email):
                    email = maybe_email.lower()
                    tokens.pop(idx)
                    break
            while tokens and tokens[-1].lower() in ROSTER_BOOL_WORDS:
                tokens.pop()
            if len(tokens) < 2:
                continue
            first_name = tokens[0]
            last_name = " ".join(tokens[1:])
            full_name = f"{first_name} {last_name}".strip()
            dedupe_key = (normalize_text_key(full_name), email)
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            rows.append({"name": full_name, "first_name": first_name, "last_name": last_name, "email": email})
    return rows


def import_roster_entries(entries, default_role="member", member_class="Member", reset_existing_passwords=True, actor=None):
    """Create or update a Member + User per roster entry. Idempotent on (email|name).

    Returns counts plus the credentials generated for the CSV. The caller commits.
    """
    role_value = normalize_role(default_role or "member")
    if role_value == "admin":
        role_value = "member"
    member_class = (member_class or "Member").strip()[:80] or "Member"

    reserved_usernames = set()
    reserved_emails = set()
    for row in User.query.order_by(User.id.asc()).all():
        if row.username:
            reserved_usernames.add(row.username.strip().lower())
        elif row.email and "@" in row.email:
            reserved_usernames.add(normalize_text_key(row.email.split("@", 1)[0]))
        if row.email:
            reserved_emails.add(row.email.strip().lower())
    for row in Member.query.order_by(Member.id.asc()).all():
        if row.email:
            reserved_emails.add(row.email.strip().lower())

    result = {"created_members": 0, "created_users": 0, "updated_users": 0, "reset_passwords": 0, "credentials": []}
    created_user_ids = []

    for entry in entries:
        first_name = entry.get("first_name") or split_name_parts(entry.get("name") or "")[0]
        last_name = entry.get("last_name") or split_name_parts(entry.get("name") or "")[1]
        full_name = " ".join((entry.get("name") or f"{first_name} {last_name}").split()).strip()
        if not full_name:
            continue
        base_username = identity.username_base_from_name(first_name, last_name)

        email = (entry.get("email") or "").strip().lower()
        if email:
            reserved_emails.add(email)

        member = None
        if email:
            member = Member.query.filter(func.lower(Member.email) == email).first()
        if not member:
            member = Member.query.filter(func.lower(Member.name) == full_name.lower()).first()
        if not member:
            if not email:
                email = identity.make_unique_import_email(base_username, reserved_emails=reserved_emails)
            member = Member(name=full_name[:120], email=email[:160], member_class=member_class)
            db.session.add(member)
            db.session.flush()
            result["created_members"] += 1
        else:
            member.name = full_name[:120]
            member.member_class = member_class
            if not member.email:
                if not email:
                    email = identity.make_unique_import_email(base_username, reserved_emails=reserved_emails)
                member.email = email[:160]
            email = (member.email or email or "").strip().lower()
            if email:
                reserved_emails.add(email)

        if not email:
            email = identity.make_unique_import_email(base_username, reserved_emails=reserved_emails)

        user = User.query.filter(func.lower(User.email) == email).first()
        if not user:
            user = User.query.filter(User.member_id == member.id).first()

        generated_password = identity.password_from_name(first_name, last_name)
        exported_password = generated_password
        if user:
            username = identity.make_unique_username(base_username, reserved=reserved_usernames, exclude_user_id=user.id)
            user.name = full_name[:160]
            user.email = email[:160]
            user.member_id = member.id
            user.username = username
            user.is_active = True
            if normalize_role(user.role) not in {"admin", "team_leader"}:
                user.role = role_value
            if reset_existing_passwords:
                user.password_hash = generate_password_hash(generated_password)
                result["reset_passwords"] += 1
            else:
                exported_password = ""  # never export a password that was not applied
            result["updated_users"] += 1
        else:
            username = identity.make_unique_username(base_username, reserved=reserved_usernames)
            user = User(
                name=full_name[:160],
                email=email[:160],
                username=username,
                password_hash=generate_password_hash(generated_password),
                role=role_value,
                is_active=True,
                member_id=member.id,
            )
            db.session.add(user)
            db.session.flush()
            created_user_ids.append(user.id)
            result["created_users"] += 1
            result["reset_passwords"] += 1

        result["credentials"].append(
            {
                "name": full_name,
                "member_id": member.id,
                "email": email,
                "username": user.username or "",
                "password": exported_password,
            }
        )

    result["_created_user_ids"] = created_user_ids
    return result


def emit_import_events(result):
    for user_id in result.get("_created_user_ids") or []:
        events.emit(events.USER_CREATED, user_id=user_id)
    events.emit(events.CHAPTER_CHANGED, reason="roster_import")


def save_roster_credentials_csv(rows) -> str:
    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    filename = f"roster_credentials_{timestamp}.csv"
    output_path = roster_credentials_dir() / filename
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["name", "member_id", "email", "username", "password"])
        for row in rows:
            writer.writerow(
                [row.get("name") or "", row.get("member_id") or "", row.get("email") or "", row.get("username") or "", row.get("password") or ""]
            )
    return filename


def credentials_pdf_bytes(rows) -> bytes:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    now = datetime.now()
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    page_width, page_height = letter
    margin_left = 32
    y = page_height - 42

    def draw_header():
        nonlocal y
        pdf.setFont("Helvetica-Bold", 13)
        pdf.drawString(margin_left, y, "ASME Member Credentials")
        pdf.setFont("Helvetica", 8.5)
        pdf.drawString(margin_left, y - 12, f"Generated: {now.strftime('%Y-%m-%d %H:%M')}")
        y -= 29
        pdf.setFont("Helvetica-Bold", 8)
        for offset, label in ((0, "Name"), (130, "Last Name"), (225, "NFC ID"), (320, "Username"), (415, "Password")):
            pdf.drawString(margin_left + offset, y, label)
        y -= 8
        pdf.line(margin_left, y, page_width - margin_left, y)
        y -= 12
        pdf.setFont("Helvetica", 8)

    draw_header()
    for row in rows:
        if y < 36:
            pdf.showPage()
            y = page_height - 42
            draw_header()
        pdf.drawString(margin_left, y, (row.get("name") or "")[:24])
        pdf.drawString(margin_left + 130, y, (row.get("last_name") or "")[:18])
        pdf.drawString(margin_left + 225, y, (row.get("nfc_uid") or "-")[:28])
        pdf.drawString(margin_left + 320, y, (row.get("username") or "")[:17])
        pdf.drawString(margin_left + 415, y, (row.get("password") or "")[:16])
        y -= 12
    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()
