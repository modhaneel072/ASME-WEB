"""Identity: users, legacy member profiles, credentials, NFC tags.

``User`` is canonical. ``Member`` rows are created/kept in sync only so that
historical tables keyed by ``member_id`` and the legacy kiosk keep working.
"""

from __future__ import annotations

import secrets
from datetime import datetime, timedelta

from sqlalchemy import func
from sqlalchemy.orm import joinedload
from werkzeug.security import check_password_hash, generate_password_hash

from asme import events
from asme.auth.session import normalize_role, role_allows
from asme.config import settings
from asme.constants import ASSIGNABLE_ROLES
from asme.extensions import db
from asme.models import (
    AttendanceRecord,
    AttendanceScan,
    AuditLog,
    ContactMessage,
    Event,
    Item,
    ItemTag,
    Member,
    NFCTag,
    PasswordResetToken,
    PrintJob,
    PrintRequest,
    Transaction,
    User,
)
from asme.services import audit
from asme.services.errors import Conflict, Forbidden, NotFound, Validation
from asme.utils import clean_tag_value, normalize_text_key, split_name_parts

# --------------------------------------------------------------------------- lookups


def find_user_by_login_identifier(identifier):
    cleaned = (identifier or "").strip().lower()
    if not cleaned:
        return None
    by_email = User.query.filter(func.lower(User.email) == cleaned).first()
    if by_email:
        return by_email
    return User.query.filter(func.lower(func.coalesce(User.username, "")) == cleaned).first()


def find_user_by_nfc_uid(tag_uid, exclude_user_id=None):
    cleaned = clean_tag_value(tag_uid)
    if not cleaned:
        return None
    query = User.query.filter(func.lower(User.nfc_uid) == cleaned.lower())
    if exclude_user_id:
        query = query.filter(User.id != exclude_user_id)
    return query.order_by(User.id.asc()).first()


def resolve_user_from_tag_uid(tag_uid):
    cleaned = clean_tag_value(tag_uid)
    if not cleaned:
        return None
    direct_user = User.query.filter(
        func.lower(User.nfc_uid) == cleaned.lower(),
        User.is_active.is_(True),
    ).first()
    if direct_user:
        return direct_user
    mapped = NFCTag.query.filter(
        func.lower(NFCTag.tag_uid) == cleaned.lower(),
        NFCTag.active.is_(True),
    ).first()
    if mapped and mapped.user and mapped.user.is_active:
        return mapped.user
    return None


def member_for_user(user):
    """Legacy profile for ``user``; links by email if the FK is missing."""
    if not user:
        return None
    if user.member:
        return user.member
    if user.email:
        linked = Member.query.filter(func.lower(Member.email) == user.email.strip().lower()).first()
        if linked:
            user.member_id = linked.id
            db.session.flush()
            return linked
    return None


def ensure_member_for_user(user) -> Member:
    """Return the legacy ``Member`` row for ``user``, creating one if none exists."""
    member = member_for_user(user)
    if member:
        return member
    member = Member(
        name=(user.name or user.email or "Member")[:120],
        email=(user.email or f"user{user.id}@asme.local")[:160],
        member_class="Member",
        nfc_tag=clean_tag_value(user.nfc_uid) or None,
    )
    db.session.add(member)
    db.session.flush()
    user.member_id = member.id
    return member


def resolve_member(member_tag, member_id):
    """Legacy kiosk lookup: NFC tag first, then explicit id."""
    member_tag = clean_tag_value(member_tag)
    member_id = str(member_id or "").strip()
    if member_tag:
        member = Member.query.filter_by(nfc_tag=member_tag).first()
        if member:
            return member
        user = resolve_user_from_tag_uid(member_tag)
        if user:
            linked = member_for_user(user)
            if linked:
                return linked
    if member_id:
        try:
            return db.session.get(Member, int(member_id))
        except Exception:
            return None
    return None


def user_for_member(member):
    if not member:
        return None
    user = User.query.filter(User.member_id == member.id).first()
    if user:
        return user
    if member.email:
        return User.query.filter(func.lower(User.email) == member.email.strip().lower()).first()
    return None


# --------------------------------------------------------------------------- generators


def username_base_from_name(first_name, last_name):
    first_key = normalize_text_key(first_name)
    last_key = normalize_text_key(last_name)
    base = f"{first_key[:1]}{last_key}"
    if not base:
        base = first_key or last_key or "member"
    return base[:50]


def password_from_name(first_name, last_name):
    first_key = normalize_text_key(first_name)
    last_key = normalize_text_key(last_name)
    prefix = f"{first_key[:2]}{last_key[:1]}"
    if len(prefix) < 3:
        prefix = (prefix + "asx")[:3]
    return f"{prefix}{secrets.randbelow(100000):05d}"


def make_unique_username(base_username, reserved=None, exclude_user_id=None):
    reserved = reserved if reserved is not None else set()
    candidate_base = normalize_text_key(base_username) or "member"
    candidate = candidate_base[:50]
    suffix = 2
    while True:
        duplicate = (
            User.query.filter(
                func.lower(func.coalesce(User.username, "")) == candidate.lower(),
                User.id != (exclude_user_id or 0),
            )
            .order_by(User.id.asc())
            .first()
        )
        if not duplicate and candidate.lower() not in reserved:
            reserved.add(candidate.lower())
            return candidate
        tail = str(suffix)
        candidate = f"{candidate_base[: max(1, 50 - len(tail))]}{tail}"
        suffix += 1


def make_unique_username_from_reserved(base_username, reserved):
    candidate_base = normalize_text_key(base_username) or "member"
    candidate = candidate_base[:50]
    suffix = 2
    while candidate.lower() in reserved:
        tail = str(suffix)
        candidate = f"{candidate_base[: max(1, 50 - len(tail))]}{tail}"
        suffix += 1
    reserved.add(candidate.lower())
    return candidate


def make_unique_import_email(base_local, reserved_emails=None):
    reserved_emails = reserved_emails if reserved_emails is not None else set()
    local = normalize_text_key(base_local) or "member"
    domain = "asme.local"
    candidate = f"{local}@{domain}"
    suffix = 2
    while True:
        duplicate_user = User.query.filter(func.lower(User.email) == candidate.lower()).first()
        duplicate_member = Member.query.filter(func.lower(Member.email) == candidate.lower()).first()
        if not duplicate_user and not duplicate_member and candidate.lower() not in reserved_emails:
            reserved_emails.add(candidate.lower())
            return candidate
        candidate = f"{local}{suffix}@{domain}"
        suffix += 1


def hash_bulk_password(password_plain):
    method = settings().bulk_password_hash_method or "pbkdf2:sha256:120000"
    try:
        return generate_password_hash(password_plain, method=method)
    except Exception:
        return generate_password_hash(password_plain)


def generate_unique_member_nfc_uid(user_id, reserved):
    try:
        numeric_id = int(user_id or 0)
    except Exception:
        numeric_id = 0
    base = f"ASME-MEMBER-{max(numeric_id, 0):05d}"
    candidate = base
    suffix = 2
    while candidate.lower() in reserved:
        candidate = f"{base}-{suffix}"
        suffix += 1
    reserved.add(candidate.lower())
    return candidate


# --------------------------------------------------------------------------- auth flows


def authenticate(identifier, password):
    user = find_user_by_login_identifier(identifier)
    if not user or not user.is_active:
        return None
    if not check_password_hash(user.password_hash, password or ""):
        return None
    return user


def signup(name, email, password, role="member") -> User:
    name = (name or "").strip()
    email = (email or "").strip().lower()
    role = normalize_role(role)
    if role == "admin":
        role = "member"
    if len(name) < 2:
        raise Validation("Please enter your full name.", field="name")
    if "@" not in email or len(email) < 5:
        raise Validation("Please enter a valid email.", field="email")
    if len(password or "") < 8:
        raise Validation("Password must be at least 8 characters.", field="password")
    if User.query.filter(func.lower(User.email) == email).first():
        raise Conflict("An account already exists for that email.", code="email_taken")

    linked_member = Member.query.filter(func.lower(Member.email) == email).first()
    username = make_unique_username(email.split("@", 1)[0] if "@" in email else name)
    user = User(
        name=name[:160],
        email=email,
        username=username,
        password_hash=generate_password_hash(password),
        role=role,
        is_active=True,
        member_id=linked_member.id if linked_member else None,
    )
    db.session.add(user)
    db.session.commit()
    events.emit(events.USER_CREATED, user_id=user.id)
    return user


def ensure_shared_admin_user(email, password):
    if not email or not password:
        return None
    user = User.query.filter(func.lower(User.email) == email).first()
    if not user:
        user = User(
            name="ASME Admin",
            email=email,
            username=(email.split("@", 1)[0] if "@" in email else "admin"),
            password_hash=generate_password_hash(password),
            role="admin",
            is_active=True,
        )
        db.session.add(user)
    else:
        user.role = "admin"
        user.is_active = True
        if not check_password_hash(user.password_hash, password):
            user.password_hash = generate_password_hash(password)
        if not (user.username or "").strip():
            user.username = make_unique_username(email.split("@", 1)[0] if "@" in email else f"admin{user.id}")
    db.session.commit()
    return user


def create_password_reset(user, hours=None) -> str:
    hours = hours or settings().password_reset_hours
    token = secrets.token_urlsafe(32)
    db.session.add(
        PasswordResetToken(user_id=user.id, token=token, expires_at=datetime.utcnow() + timedelta(hours=hours))
    )
    db.session.commit()
    return token


def find_valid_reset(token):
    reset_row = PasswordResetToken.query.filter_by(token=(token or "").strip()).first()
    if (
        not reset_row
        or reset_row.used_at is not None
        or reset_row.expires_at is None
        or reset_row.expires_at < datetime.utcnow()
    ):
        return None
    return reset_row


def consume_password_reset(reset_row, password, confirm) -> User:
    if len(password or "") < 8:
        raise Validation("Password must be at least 8 characters.", field="password")
    if password != confirm:
        raise Validation("Passwords do not match.", field="confirm_password")
    user = db.session.get(User, reset_row.user_id)
    if not user:
        raise NotFound("User no longer exists.")
    user.password_hash = generate_password_hash(password)
    reset_row.used_at = datetime.utcnow()
    db.session.commit()
    return user


def update_profile(user, name, email, major=None, graduation_year=None, phone=None) -> User:
    name = (name or "").strip()
    email = (email or "").strip().lower()
    if not name or "@" not in email:
        raise Validation("Provide a valid name and email.")
    if email != user.email and User.query.filter(func.lower(User.email) == email, User.id != user.id).first():
        raise Conflict("Email already in use by another account.", code="email_taken")
    user.name = name[:160]
    user.email = email[:160]
    if major is not None:
        user.major = (major or "").strip()[:120] or None
    if graduation_year is not None:
        user.graduation_year = graduation_year or None
    if phone is not None:
        user.phone = (phone or "").strip()[:40] or None
    if user.member:
        user.member.name = user.name
        user.member.email = user.email
    db.session.commit()
    events.emit(events.USER_UPDATED, user_id=user.id)
    return user


def change_password(user, current_password, new_password, confirm_password):
    if not check_password_hash(user.password_hash, current_password or ""):
        raise Validation("Current password is incorrect.", field="current_password")
    if len(new_password or "") < 8:
        raise Validation("New password must be at least 8 characters.", field="new_password")
    if new_password != confirm_password:
        raise Validation("New password and confirm password do not match.", field="confirm_password")
    user.password_hash = generate_password_hash(new_password)
    db.session.commit()


# --------------------------------------------------------------------------- admin: people


def _validate_grad_year(raw):
    from asme.utils import parse_non_negative_int

    raw = (raw or "").strip()
    year = parse_non_negative_int(raw, default=0) if raw else 0
    if year and (year < 1900 or year > 2100):
        return 0
    return year


def _check_nfc_free(nfc_uid, exclude_user_id=None, exclude_member_id=None):
    if not nfc_uid:
        return
    if find_user_by_nfc_uid(nfc_uid, exclude_user_id=exclude_user_id):
        raise Conflict("That NFC UID is already assigned to another user.", code="nfc_taken")
    member_conflict = Member.query.filter(
        func.lower(Member.nfc_tag) == nfc_uid.lower(),
        Member.id != (exclude_member_id or 0),
    ).first()
    if member_conflict:
        raise Conflict("That NFC UID is already assigned to another member record.", code="nfc_taken")


def admin_create_member(form, actor):
    """Create a legacy Member profile and (optionally) its login."""
    name = (form.get("name") or "").strip()
    email = (form.get("email") or "").strip().lower()
    member_class = (form.get("member_class") or "").strip() or "Member"
    create_user = (form.get("create_user") or "1").strip() in {"1", "true", "on", "yes"}
    role = normalize_role(form.get("role") or "member")
    password = form.get("password") or ""
    major = (form.get("major") or "").strip()[:120] or None
    graduation_year = _validate_grad_year(form.get("graduation_year"))
    nfc_uid = clean_tag_value(form.get("nfc_uid"))
    exec_title = (form.get("exec_title") or "").strip()[:160] or None
    exec_message = (form.get("exec_message") or "").strip()[:500] or None
    headshot_url = (form.get("headshot_url") or "").strip()[:500] or None
    first_name, last_name = split_name_parts(name)
    username_base = username_base_from_name(first_name, last_name)

    if not name or "@" not in email:
        raise Validation("Enter a valid member name and email.")
    if Member.query.filter(func.lower(Member.email) == email).first():
        raise Conflict("A member with that email already exists.", code="email_taken")
    _check_nfc_free(nfc_uid)

    member = Member(
        name=name[:120],
        email=email[:160],
        member_class=member_class[:80] or "Member",
        nfc_tag=nfc_uid or None,
    )
    db.session.add(member)
    db.session.flush()

    linked_user = None
    if create_user:
        existing_user = User.query.filter(func.lower(User.email) == email).first()
        if existing_user:
            if existing_user.member_id and existing_user.member_id != member.id:
                db.session.rollback()
                raise Conflict("That email already belongs to a different user/member link.", code="link_conflict")
            if password and len(password) < 8:
                db.session.rollback()
                raise Validation("Password must be at least 8 characters.", field="password")
            existing_user.member_id = member.id
            existing_user.name = name[:160]
            existing_user.is_active = True
            existing_user.role = role
            if not (existing_user.username or "").strip():
                existing_user.username = make_unique_username(
                    username_base or email.split("@", 1)[0], exclude_user_id=existing_user.id
                )
            existing_user.major = major
            existing_user.graduation_year = graduation_year or None
            existing_user.nfc_uid = nfc_uid or None
            existing_user.exec_title = exec_title
            existing_user.exec_message = exec_message
            existing_user.headshot_url = headshot_url
            if password:
                existing_user.password_hash = generate_password_hash(password)
            linked_user = existing_user
        else:
            if len(password) < 8:
                db.session.rollback()
                raise Validation("Password is required (min 8 chars) when creating a new login account.", field="password")
            linked_user = User(
                name=name[:160],
                email=email[:160],
                username=make_unique_username(username_base or email.split("@", 1)[0]),
                password_hash=generate_password_hash(password),
                role=role,
                is_active=True,
                nfc_uid=nfc_uid or None,
                major=major,
                graduation_year=graduation_year or None,
                exec_title=exec_title,
                exec_message=exec_message,
                headshot_url=headshot_url,
                member_id=member.id,
            )
            db.session.add(linked_user)

    audit.record("create_member", f"member_id={member.id} email={member.email} create_user={bool(linked_user)}", actor=actor)
    db.session.commit()
    if linked_user:
        events.emit(events.USER_CREATED, user_id=linked_user.id)
    return member, linked_user


def admin_create_user(form, actor) -> User:
    from asme.utils import parse_positive_int

    name = (form.get("name") or "").strip()
    email = (form.get("email") or "").strip().lower()
    password = form.get("password") or ""
    role = normalize_role(form.get("role") or "member")
    member_id = parse_positive_int(form.get("member_id"), default=0)
    major = (form.get("major") or "").strip()[:120] or None
    graduation_year = _validate_grad_year(form.get("graduation_year"))
    nfc_uid = clean_tag_value(form.get("nfc_uid"))
    exec_title = (form.get("exec_title") or "").strip()[:160] or None
    exec_message = (form.get("exec_message") or "").strip()[:500] or None
    headshot_url = (form.get("headshot_url") or "").strip()[:500] or None
    first_name, last_name = split_name_parts(name)
    username_base = username_base_from_name(first_name, last_name)

    if not name or "@" not in email or len(password) < 8:
        raise Validation("Enter valid name, email, and password (min 8 chars).")
    if User.query.filter(func.lower(User.email) == email).first():
        raise Conflict("User with that email already exists.", code="email_taken")
    _check_nfc_free(nfc_uid)

    member = db.session.get(Member, member_id) if member_id else None
    if not member:
        member = Member.query.filter(func.lower(Member.email) == email).first()
    if member:
        linked_existing = User.query.filter(User.member_id == member.id).first()
        if linked_existing:
            raise Conflict(f"Member ID #{member.id} is already linked to {linked_existing.email}.", code="link_conflict")

    new_user = User(
        name=name[:160],
        email=email[:160],
        username=make_unique_username(username_base or email.split("@", 1)[0]),
        password_hash=generate_password_hash(password),
        role=role,
        is_active=True,
        nfc_uid=nfc_uid or None,
        major=major,
        graduation_year=graduation_year or None,
        exec_title=exec_title,
        exec_message=exec_message,
        headshot_url=headshot_url,
        member_id=member.id if member else None,
    )
    if member and nfc_uid:
        member.nfc_tag = nfc_uid
    db.session.add(new_user)
    audit.record("create_user", f"{new_user.email} role={new_user.role}", actor=actor)
    db.session.commit()
    events.emit(events.USER_CREATED, user_id=new_user.id)
    return new_user


def admin_update_user(user, form, actor) -> User:
    from asme.utils import parse_positive_int

    previous_member = user.member
    previous_nfc_uid = clean_tag_value(user.nfc_uid)
    previous_role = normalize_role(user.role)
    role = normalize_role(form.get("role") or "member")
    active_raw = (form.get("is_active") or "1").strip()
    name = (form.get("name") or user.name).strip()
    email = (form.get("email") or user.email).strip().lower()
    username_raw = (form.get("username") or "").strip().lower()
    member_id_raw = (form.get("member_id") or "").strip()
    member_id = parse_positive_int(member_id_raw, default=0) if member_id_raw else 0
    member_row = db.session.get(Member, member_id) if member_id else None
    if member_id and not member_row:
        raise NotFound("Member ID not found.")
    if member_row:
        member_conflict = User.query.filter(User.member_id == member_row.id, User.id != user.id).first()
        if member_conflict:
            raise Conflict("That member ID is already linked to another user.", code="link_conflict")
    major = (form.get("major") or "").strip()[:120] or None
    graduation_year = _validate_grad_year(form.get("graduation_year"))
    nfc_uid = clean_tag_value(form.get("nfc_uid"))
    exec_title = (form.get("exec_title") or "").strip()[:160] or None
    exec_message = (form.get("exec_message") or "").strip()[:500] or None
    headshot_url = (form.get("headshot_url") or "").strip()[:500] or None
    if "@" not in email:
        raise Validation("Enter a valid email.", field="email")
    duplicate = User.query.filter(func.lower(User.email) == email, User.id != user.id).order_by(User.id.asc()).first()
    if duplicate:
        raise Conflict("Another user already has that email.", code="email_taken")
    requested_username = normalize_text_key(username_raw or "")
    if requested_username:
        username_conflict = (
            User.query.filter(
                func.lower(func.coalesce(User.username, "")) == requested_username.lower(),
                User.id != user.id,
            )
            .order_by(User.id.asc())
            .first()
        )
        if username_conflict:
            raise Conflict("That username is already in use.", code="username_taken")
    _check_nfc_free(nfc_uid, exclude_user_id=user.id, exclude_member_id=member_row.id if member_row else None)

    user.name = name[:160] or user.name
    user.email = email[:160]
    if requested_username:
        user.username = requested_username[:80]
    elif not (user.username or "").strip():
        first_name, last_name = split_name_parts(user.name)
        user.username = make_unique_username(
            username_base_from_name(first_name, last_name) or user.email.split("@", 1)[0],
            exclude_user_id=user.id,
        )
    user.role = role
    user.is_active = active_raw == "1"
    user.member_id = member_row.id if member_row else None
    user.nfc_uid = nfc_uid or None
    user.major = major
    user.graduation_year = graduation_year or None
    user.exec_title = exec_title
    user.exec_message = exec_message
    user.headshot_url = headshot_url
    if previous_member and (not member_row or previous_member.id != member_row.id):
        if clean_tag_value(previous_member.nfc_tag).lower() == previous_nfc_uid.lower():
            previous_member.nfc_tag = None
    if member_row:
        member_row.nfc_tag = nfc_uid or None
    audit.record(
        "update_user_role",
        (
            f"user_id={user.id} role={role} active={user.is_active} "
            f"email={user.email} member_id={user.member_id or ''} nfc_uid={user.nfc_uid or ''}"
        ),
        actor=actor,
    )
    db.session.commit()
    events.emit(events.USER_UPDATED, user_id=user.id, role_changed=(previous_role != role))
    if nfc_uid and nfc_uid.lower() != previous_nfc_uid.lower():
        events.emit(events.NFC_ASSIGNED, user_id=user.id, tag_uid=nfc_uid)
    return user


def admin_reset_password(user, provided, actor) -> str:
    from uuid import uuid4

    provided = (provided or "").strip()
    if provided and len(provided) < 8:
        raise Validation("Password must be at least 8 characters.", field="new_password")
    new_password = provided or f"ASME-{uuid4().hex[:10]}"
    user.password_hash = generate_password_hash(new_password)
    audit.record("reset_user_password", f"user_id={user.id}", actor=actor)
    db.session.commit()
    return new_password


def admin_invite_link_token(user, actor) -> str:
    token = secrets.token_urlsafe(32)
    db.session.add(PasswordResetToken(user_id=user.id, token=token, expires_at=datetime.utcnow() + timedelta(hours=72)))
    audit.record("create_user_invite_link", f"user_id={user.id}", actor=actor)
    db.session.commit()
    return token


def _user_reference_counts(user_id):
    return {
        "transactions": Transaction.query.filter(Transaction.user_id == user_id).count(),
        "print_requests": PrintRequest.query.filter(PrintRequest.user_id == user_id).count(),
        "events_created": Event.query.filter(Event.created_by_user_id == user_id).count(),
        "events_requested": Event.query.filter(Event.requested_by_user_id == user_id).count(),
        "attendance": AttendanceRecord.query.filter(AttendanceRecord.user_id == user_id).count(),
        "contact_messages": ContactMessage.query.filter(ContactMessage.user_id == user_id).count(),
        "audit_logs": AuditLog.query.filter(AuditLog.admin_user_id == user_id).count(),
    }


def admin_delete_user(user, actor) -> str:
    """Hard-delete a user with no history, otherwise deactivate. Returns which happened."""
    if actor and actor.id == user.id:
        raise Forbidden("You cannot delete your own account while logged in.", code="self_delete")

    previous_uid = clean_tag_value(user.nfc_uid)
    related_member = user.member
    for row in NFCTag.query.filter_by(user_id=user.id, active=True).all():
        row.active = False
        row.unassigned_at = datetime.utcnow()
    user.nfc_uid = None
    if related_member and previous_uid and clean_tag_value(related_member.nfc_tag).lower() == previous_uid.lower():
        related_member.nfc_tag = None

    references = _user_reference_counts(user.id)
    references["password_tokens"] = PasswordResetToken.query.filter(PasswordResetToken.user_id == user.id).count()
    references["nfc_history"] = NFCTag.query.filter(NFCTag.user_id == user.id).count()
    total_refs = sum(references.values())

    if total_refs == 0:
        db.session.delete(user)
        audit.record("delete_user", f"user_id={user.id} hard_delete=1", actor=actor)
        db.session.commit()
        return "deleted"

    user.is_active = False
    audit.record("delete_user", f"user_id={user.id} hard_delete=0 refs={total_refs}", actor=actor)
    db.session.commit()
    return "deactivated"


def reset_members(actor) -> dict:
    """End-of-year purge: drop member/team_leader logins without history, deactivate the rest."""
    now = datetime.utcnow()
    counts = {"deleted_users": 0, "deactivated_users": 0, "deleted_members": 0, "kept_members": 0}

    target_users = User.query.filter(User.role.in_(["member", "team_leader"])).order_by(User.id.asc()).all()
    for user in target_users:
        existing_uid = clean_tag_value(user.nfc_uid)
        for row in NFCTag.query.filter_by(user_id=user.id, active=True).all():
            row.active = False
            row.unassigned_at = now
        NFCTag.query.filter(NFCTag.assigned_by_user_id == user.id).update(
            {NFCTag.assigned_by_user_id: None}, synchronize_session=False
        )
        NFCTag.query.filter(NFCTag.user_id == user.id).delete(synchronize_session=False)
        PasswordResetToken.query.filter(PasswordResetToken.user_id == user.id).delete(synchronize_session=False)
        if user.member and existing_uid and clean_tag_value(user.member.nfc_tag).lower() == existing_uid.lower():
            user.member.nfc_tag = None
        user.nfc_uid = None

        if sum(_user_reference_counts(user.id).values()) == 0:
            db.session.delete(user)
            counts["deleted_users"] += 1
            continue
        user.is_active = False
        user.member_id = None
        counts["deactivated_users"] += 1

    for member in Member.query.order_by(Member.id.asc()).all():
        member.nfc_tag = None
        has_user = User.query.filter(User.member_id == member.id).first() is not None
        has_refs = (
            Transaction.query.filter(Transaction.member_id == member.id).count()
            + AttendanceScan.query.filter(AttendanceScan.member_id == member.id).count()
            + AttendanceRecord.query.filter(AttendanceRecord.member_id == member.id).count()
            + PrintJob.query.filter(PrintJob.member_id == member.id).count()
            + PrintRequest.query.filter(PrintRequest.member_id == member.id).count()
        )
        if has_user or has_refs:
            counts["kept_members"] += 1
            continue
        db.session.delete(member)
        counts["deleted_members"] += 1

    audit.record(
        "reset_members",
        " ".join(f"{key}={value}" for key, value in counts.items()),
        actor=actor,
    )
    db.session.commit()
    return counts


# --------------------------------------------------------------------------- NFC tags


def assign_nfc(user, tag_uid, notes, actor) -> NFCTag:
    tag_uid = clean_tag_value(tag_uid)
    if not tag_uid:
        raise Validation("Tag UID is required.", field="tag_uid")
    if find_user_by_nfc_uid(tag_uid, exclude_user_id=user.id):
        raise Conflict("That tag UID is already assigned to another user account.", code="nfc_taken")
    if ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag_uid.lower()).first():
        raise Conflict("That UID is already assigned to an inventory item tag.", code="nfc_taken")
    if Item.query.filter(func.lower(Item.nfc_tag) == tag_uid.lower()).first():
        raise Conflict("That UID is already assigned to an inventory item.", code="nfc_taken")
    existing_active = NFCTag.query.filter(func.lower(NFCTag.tag_uid) == tag_uid.lower(), NFCTag.active.is_(True)).first()
    if existing_active and existing_active.user_id != user.id:
        raise Conflict("That tag UID is already assigned.", code="nfc_taken")

    for row in NFCTag.query.filter_by(user_id=user.id, active=True).all():
        row.active = False
        row.unassigned_at = datetime.utcnow()

    if existing_active and existing_active.user_id == user.id:
        existing_active.notes = notes
        existing_active.assigned_at = datetime.utcnow()
        existing_active.active = True
        existing_active.unassigned_at = None
        tag_row = existing_active
    else:
        tag_row = NFCTag(
            tag_uid=tag_uid,
            user_id=user.id,
            active=True,
            assigned_at=datetime.utcnow(),
            assigned_by_user_id=actor.id if actor else None,
            notes=notes,
        )
        db.session.add(tag_row)
    user.nfc_uid = tag_uid
    if user.member:
        user.member.nfc_tag = tag_uid
    audit.record("assign_nfc", f"user_id={user.id} tag_uid={tag_uid}", actor=actor)
    db.session.commit()
    events.emit(events.NFC_ASSIGNED, user_id=user.id, tag_uid=tag_uid)
    return tag_row


def unassign_nfc(tag_row, actor):
    tag_row.active = False
    tag_row.unassigned_at = datetime.utcnow()
    if tag_row.user and clean_tag_value(tag_row.user.nfc_uid).lower() == clean_tag_value(tag_row.tag_uid).lower():
        tag_row.user.nfc_uid = None
        if tag_row.user.member and clean_tag_value(tag_row.user.member.nfc_tag).lower() == clean_tag_value(tag_row.tag_uid).lower():
            tag_row.user.member.nfc_tag = None
    audit.record("unassign_nfc", f"tag_id={tag_row.id} uid={tag_row.tag_uid}", actor=actor)
    db.session.commit()
    if tag_row.user:
        events.emit(events.USER_UPDATED, user_id=tag_row.user_id)


# --------------------------------------------------------------------------- credentials export


def build_fresh_member_credentials():
    """Regenerate usernames/passwords/NFC ids for every active member login.

    Mutates users in the session; the caller commits.
    """
    users = (
        User.query.options(joinedload(User.member))
        .filter(User.is_active.is_(True), User.role.in_(["member", "team_leader"]))
        .order_by(User.name.asc(), User.id.asc())
        .all()
    )
    if not users:
        return []

    target_user_ids = {user.id for user in users}
    target_member_ids = {user.member_id for user in users if user.member_id}
    reserved_nfc = set()

    def _reserve(values):
        for (tag_value,) in values:
            normalized = clean_tag_value(tag_value)
            if normalized:
                reserved_nfc.add(normalized.lower())

    _reserve(ItemTag.query.with_entities(ItemTag.tag_value).all())
    _reserve(Item.query.filter(Item.nfc_tag.isnot(None)).with_entities(Item.nfc_tag).all())
    other_users_query = User.query.filter(User.nfc_uid.isnot(None))
    if target_user_ids:
        other_users_query = other_users_query.filter(User.id.notin_(target_user_ids))
    _reserve(other_users_query.with_entities(User.nfc_uid).all())
    other_member_query = Member.query.filter(Member.nfc_tag.isnot(None))
    if target_member_ids:
        other_member_query = other_member_query.filter(Member.id.notin_(target_member_ids))
    _reserve(other_member_query.with_entities(Member.nfc_tag).all())
    other_active_tag_query = NFCTag.query.filter(NFCTag.active.is_(True))
    if target_user_ids:
        other_active_tag_query = other_active_tag_query.filter(NFCTag.user_id.notin_(target_user_ids))
    _reserve(other_active_tag_query.with_entities(NFCTag.tag_uid).all())

    active_tag_by_user = {}
    for tag_row in NFCTag.query.filter(NFCTag.active.is_(True)).order_by(NFCTag.assigned_at.desc(), NFCTag.id.desc()).all():
        if tag_row.user_id and tag_row.user_id not in active_tag_by_user:
            active_tag_by_user[tag_row.user_id] = clean_tag_value(tag_row.tag_uid)

    reserved_usernames = {
        (row.username or "").strip().lower()
        for row in User.query.filter(User.username.isnot(None)).all()
        if (row.username or "").strip()
    }

    rows = []
    for user in users:
        first_name, last_name = split_name_parts(user.name or "")
        if not (first_name or last_name):
            first_name, last_name = split_name_parts(user.email.split("@", 1)[0] if user.email else "")
        if not (user.username or "").strip():
            base_username = username_base_from_name(first_name, last_name) or (
                user.email.split("@", 1)[0] if user.email and "@" in user.email else f"user{user.id}"
            )
            user.username = make_unique_username_from_reserved(base_username, reserved_usernames)
        password_plain = password_from_name(first_name, last_name)
        user.password_hash = hash_bulk_password(password_plain)
        member_nfc = clean_tag_value(user.member.nfc_tag) if user.member and user.member.nfc_tag else ""
        tag_table_nfc = clean_tag_value(active_tag_by_user.get(user.id, ""))
        user_nfc = clean_tag_value(user.nfc_uid)
        resolved_nfc = ""
        for candidate in (user_nfc, tag_table_nfc, member_nfc):
            if candidate and candidate.lower() not in reserved_nfc:
                resolved_nfc = candidate
                reserved_nfc.add(candidate.lower())
                break
        if not resolved_nfc:
            resolved_nfc = generate_unique_member_nfc_uid(user.id, reserved_nfc)
        user.nfc_uid = resolved_nfc
        if user.member:
            user.member.nfc_tag = resolved_nfc
        rows.append(
            {
                "name": (user.name or "").strip(),
                "first_name": first_name,
                "last_name": last_name,
                "member_id": user.member_id or "",
                "nfc_uid": resolved_nfc,
                "username": (user.username or "").strip(),
                "password": password_plain,
            }
        )
    return rows


def can_manage_roles(actor, target_role):
    return bool(actor) and role_allows(actor.role, "admin") and normalize_role(target_role) in ASSIGNABLE_ROLES
