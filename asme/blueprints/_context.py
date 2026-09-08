"""Template context builders shared by the HTML blueprints.

These read; they never write. Anything that mutates lives in ``asme.services``.
"""

from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta
from pathlib import Path

from flask import current_app, render_template, url_for
from sqlalchemy import func

from asme.auth.session import current_auth_user, current_user_member, get_active_member, is_admin_member, normalize_role, role_allows
from asme.config import settings
from asme.constants import ITEM_TYPES, PRINT_REQUEST_PRINTERS, PRINT_REQUEST_STATUSES, PRINTER_TYPES
from asme.content_data import (
    EXEC_PROFILE_OVERRIDES,
    FRONT_CLUB_HIGHLIGHTS,
    FRONT_CLUB_MISSION,
    FRONT_PROJECT_SHOWCASE,
    MANUAL_EXECUTIVE_PROFILES,
)
from asme.extensions import db
from asme.models import (
    AttendanceRecord,
    AttendanceScan,
    ContactMessage,
    Event,
    Item,
    Member,
    NFCTag,
    PrintRequest,
    Project,
    Transaction,
    User,
)
from asme.services import attendance, audit, content, fabrication, inventory, scheduling
from asme.utils import default_due_date, normalize_text_key

# --------------------------------------------------------------------------- executive cards


def parse_executive_message(raw_message):
    raw_text = (raw_message or "").replace("\r", "\n")
    if not raw_text.strip():
        return {"summary": "", "focus_points": [], "contact": "", "note": ""}
    raw_text = re.sub(r"\s+(?=\d+[.)]\s+)", "\n", raw_text)
    lines = []
    for line in raw_text.splitlines():
        cleaned = re.sub(r"^\s*(?:[-*•]+|\d+[.)])\s*", "", (line or "").strip())
        if cleaned:
            lines.append(cleaned)
    if not lines:
        return {"summary": "", "focus_points": [], "contact": "", "note": ""}

    summary, focus_points, contact, note = "", [], "", ""
    phone_re = re.compile(r"\d{3}[-.\s]?\d{3}[-.\s]?\d{4}")
    for line in lines:
        lowered = line.lower()
        has_contact = ("@" in line and "." in line) or bool(phone_re.search(line))
        if has_contact and not contact:
            contact = line
            continue
        if "headshot" in lowered and not note:
            note = line
            continue
        if not summary:
            summary = line
            continue
        if len(focus_points) < 3:
            focus_points.append(line)
    if not summary:
        summary = lines[0]
    if not focus_points:
        for candidate in lines[1:4]:
            if candidate != contact and "headshot" not in candidate.lower():
                focus_points.append(candidate)
    return {"summary": summary, "focus_points": focus_points[:3], "contact": contact, "note": note}


def executive_profile_override_for_user(user):
    if not user:
        return {}
    email_key = f"email:{(user.email or '').strip().lower()}"
    name_key = f"name:{normalize_text_key(user.name or '')}"
    return EXEC_PROFILE_OVERRIDES.get(email_key) or EXEC_PROFILE_OVERRIDES.get(name_key) or {}


def resolve_exec_headshot_url(headshot_url):
    candidate = (headshot_url or "").strip()
    if not candidate:
        return url_for("static", filename="asme_logo.png")
    lowered = candidate.lower()
    if lowered.startswith("http://") or lowered.startswith("https://"):
        return candidate
    if candidate.startswith("/static/"):
        static_path = Path(current_app.static_folder) / candidate.removeprefix("/static/")
        if static_path.exists():
            return candidate
    return url_for("static", filename="asme_logo.png")


def build_executive_cards():
    cfg = settings()
    executives = User.query.filter(User.is_active.is_(True)).order_by(User.role.desc(), User.name.asc()).all()
    shared_admin_emails = {email for email in (cfg.default_admin_email, cfg.shared_admin_email) if email}
    cards = []
    for exec_user in executives:
        override = executive_profile_override_for_user(exec_user)
        raw_exec_title = (exec_user.exec_title or "").strip()
        raw_exec_message = (exec_user.exec_message or "").strip()
        if not (raw_exec_title or raw_exec_message or override):
            continue
        if (exec_user.email or "").strip().lower() in shared_admin_emails and not override and not raw_exec_title and not raw_exec_message:
            continue
        title = "Executive Member"
        if normalize_role(exec_user.role) == "admin":
            title = "Administrator"
        elif normalize_role(exec_user.role) == "team_leader":
            title = "Team Leader"
        exec_title = raw_exec_title or (override.get("title") or "").strip() or title
        exec_message_raw = raw_exec_message or (override.get("message") or "").strip() or "Executive profile details will be added soon."
        parsed = parse_executive_message(exec_message_raw)
        cards.append(
            {
                "name": (override.get("name") or "").strip() or exec_user.name,
                "title": exec_title,
                "message": exec_message_raw,
                "summary": parsed.get("summary") or "",
                "focus_points": parsed.get("focus_points") or [],
                "contact": parsed.get("contact") or "",
                "note": parsed.get("note") or "",
                "headshot": resolve_exec_headshot_url((override.get("headshot_url") or "").strip() or (exec_user.headshot_url or "").strip()),
                "_alt_name_keys": {normalize_text_key(exec_user.name or "")},
                "_email": (exec_user.email or "").strip().lower(),
            }
        )
    existing_keys, existing_emails = set(), set()
    email_pattern = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
    for card in cards:
        existing_keys.add(normalize_text_key(card.get("name") or ""))
        existing_keys.update(card.pop("_alt_name_keys", set()) or set())
        card_email = card.pop("_email", "") or ""
        if card_email:
            existing_emails.add(card_email)
        match = email_pattern.search(card.get("contact") or "")
        if match:
            existing_emails.add(match.group(0).lower())
    for profile in MANUAL_EXECUTIVE_PROFILES:
        name = (profile.get("name") or "").strip()
        if not name or normalize_text_key(name) in existing_keys:
            continue
        profile_message = (profile.get("message") or "").strip()
        match = email_pattern.search(profile_message)
        if match and match.group(0).lower() in existing_emails:
            continue
        parsed = parse_executive_message(profile_message)
        cards.append(
            {
                "name": name,
                "title": (profile.get("title") or "Executive Member").strip(),
                "message": profile_message,
                "summary": parsed.get("summary") or "",
                "focus_points": parsed.get("focus_points") or [],
                "contact": parsed.get("contact") or "",
                "note": parsed.get("note") or "",
                "headshot": resolve_exec_headshot_url((profile.get("headshot_url") or "").strip()),
            }
        )
        existing_keys.add(normalize_text_key(name))
    if not cards:
        fallback = "Add executive member accounts to populate this section."
        parsed = parse_executive_message(fallback)
        cards = [
            {
                "name": "ASME Executive Team",
                "title": "Leadership",
                "message": fallback,
                "summary": parsed.get("summary") or fallback,
                "focus_points": [],
                "contact": "",
                "note": "",
                "headshot": url_for("static", filename="asme_logo.png"),
            }
        ]
    return cards


# --------------------------------------------------------------------------- public / portal contexts


def public_site_context(page_title):
    user = current_auth_user()
    projects = Project.query.order_by(Project.created_at.desc(), Project.id.desc()).all()
    project_filters = sorted({(project.project_type or "General").strip() for project in projects if project}) or ["General"]
    return {
        "page_title": page_title,
        "current_user": user,
        "projects": projects,
        "project_filters": project_filters,
        "announcements": content.public_announcements(limit=5),
        "executive_cards": build_executive_cards(),
    }


def member_dashboard_context():
    from asme.services.onboarding import engine

    user = current_auth_user()
    member = current_user_member(user)
    open_checkouts = inventory.open_loans_for(user=user, member=member)
    items = Item.query.filter(Item.active.is_(True)).order_by(Item.name.asc(), Item.id.asc()).all()
    print_requests = (
        PrintRequest.query.filter_by(user_id=user.id).order_by(PrintRequest.created_at.desc(), PrintRequest.id.desc()).all()
    )
    my_help_messages = (
        ContactMessage.query.filter_by(user_id=user.id, kind="help")
        .order_by(ContactMessage.created_at.desc(), ContactMessage.id.desc())
        .limit(40)
        .all()
    )
    return {
        "current_user": user,
        "member_profile": member,
        "items": items,
        "open_checkouts": open_checkouts,
        "print_requests": print_requests,
        "portal_print_printers": PRINT_REQUEST_PRINTERS,
        "upcoming_events": scheduling.upcoming_events(limit=25),
        "calendar_embed_url": scheduling.calendar_embed_url(),
        "announcements": content.member_announcements(limit=8),
        "my_help_messages": my_help_messages,
        "launchpad": engine.summary_for_user(user),
        "onboarding_enforced": settings().onboarding_enforce,
    }


def admin_dashboard_context():
    user = current_auth_user()
    now = datetime.now()
    active_members_count = User.query.filter(User.is_active.is_(True), User.role.in_(["member", "team_leader", "admin"])).count()
    checked_out_now_count = (
        db.session.query(func.coalesce(func.sum(Transaction.qty), 0)).filter(Transaction.status == "OUT").scalar() or 0
    )
    cfg = settings()
    return {
        "current_user": user,
        "members": Member.query.order_by(Member.name.asc()).all(),
        "users": User.query.order_by(User.created_at.desc(), User.id.desc()).all(),
        "tags": NFCTag.query.order_by(NFCTag.assigned_at.desc(), NFCTag.id.desc()).all(),
        "items": Item.query.order_by(Item.name.asc()).all(),
        "item_primary_tags": inventory.get_item_primary_tag_map(),
        "transactions": Transaction.query.order_by(Transaction.timestamp.desc(), Transaction.id.desc()).limit(120).all(),
        "print_requests": PrintRequest.query.order_by(PrintRequest.created_at.desc(), PrintRequest.id.desc()).limit(120).all(),
        "events": Event.query.order_by(Event.start_time.desc(), Event.id.desc()).limit(120).all(),
        "attendance_records": AttendanceRecord.query.order_by(AttendanceRecord.checkin_time.desc(), AttendanceRecord.id.desc()).limit(200).all(),
        "projects": Project.query.order_by(Project.created_at.desc(), Project.id.desc()).all(),
        "contact_messages": ContactMessage.query.order_by(ContactMessage.created_at.desc(), ContactMessage.id.desc()).limit(60).all(),
        "audit_logs": audit.recent(150),
        "roles": ["member", "team_leader", "admin"],
        "print_request_statuses": PRINT_REQUEST_STATUSES,
        "portal_print_printers": PRINT_REQUEST_PRINTERS,
        "item_types": ITEM_TYPES,
        "active_members_count": active_members_count,
        "checked_out_now_count": int(checked_out_now_count),
        "attendance_today_count": attendance.today_count(),
        "overdue_items_count": len(inventory.overdue_loans()),
        "low_stock_items_count": len(inventory.low_stock_items()),
        "upcoming_meetings_count": Event.query.filter(Event.start_time >= now).count(),
        "pending_print_count": fabrication.pending_count(),
        "calendar_provider": cfg.calendar_provider,
        "google_embed_set": bool(cfg.google_calendar_embed_url),
        "outlook_embed_set": bool(cfg.outlook_calendar_embed_url),
    }


# --------------------------------------------------------------------------- legacy ops contexts


def legacy_dashboard_context(transaction_limit=15):
    members = Member.query.order_by(Member.name.asc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    recent_transactions = Transaction.query.order_by(Transaction.timestamp.desc()).limit(transaction_limit).all()
    today_attendance = attendance.today_unique_scans()
    queues = fabrication.queue_snapshot()
    return {
        "members": members,
        "items": items,
        "recent_transactions": recent_transactions,
        "today_attendance": today_attendance,
        "attendance_count": len(today_attendance),
        "queues": queues,
        "default_due": str(default_due_date()),
        "today": str(date.today()),
        "low_stock_count": len([item for item in items if item.available_qty <= 2]),
        "active_prints_count": len([printer for printer in PRINTER_TYPES if queues[printer]["active"]]),
        "h2s_waiting_count": len(queues["H2S"]["queued"]),
        "p1s_waiting_count": len(queues["P1S"]["queued"]),
    }


def render_ops_page(template_name, active_page, page_title, page_subtitle, transaction_limit=15):
    cfg = settings()
    active_member = get_active_member()
    context = legacy_dashboard_context(transaction_limit=transaction_limit)
    context.update(
        {
            "active_page": active_page,
            "page_title": page_title,
            "page_subtitle": page_subtitle,
            "active_member": active_member,
            "active_member_is_admin": is_admin_member(active_member),
            "h2s_print_cmd_configured": bool(cfg.h2s_print_cmd),
            "p1s_print_cmd_configured": bool(cfg.p1s_print_cmd),
            "h2s_print_cmd_value": cfg.h2s_print_cmd,
            "p1s_print_cmd_value": cfg.p1s_print_cmd,
            "print_commands_env_file": str(Path(current_app.instance_path) / "print_commands.env"),
        }
    )
    return render_template(template_name, **context)


def frontend_portal_context():
    active_member = get_active_member()
    my_open_count = Transaction.query.filter_by(member_id=active_member.id, status="OUT").count() if active_member else 0
    return {
        "active_member": active_member,
        "active_member_is_admin": is_admin_member(active_member),
        "members": Member.query.order_by(Member.name.asc()).all(),
        "item_count": Item.query.count(),
        "available_total": int(db.session.query(func.coalesce(func.sum(Item.available_qty), 0)).scalar() or 0),
        "my_open_count": my_open_count,
        "club_mission": FRONT_CLUB_MISSION,
        "club_highlights": FRONT_CLUB_HIGHLIGHTS,
        "project_showcase": FRONT_PROJECT_SHOWCASE,
    }


def can_switch_to_team(user):
    return bool(user) and role_allows(user.role, "team_leader")
