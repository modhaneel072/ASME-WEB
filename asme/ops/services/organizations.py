"""Organisation profile, memberships, roles and the setup banner."""

from __future__ import annotations

from datetime import datetime

from sqlalchemy import func
from werkzeug.security import generate_password_hash

from asme import events
from asme.extensions import db
from asme.models import User
from asme.ops import audit, authz
from asme.ops.models import Membership, Organization, Role
from asme.ops.services.common import user_ref
from asme.ops.tenancy import default_organization, ensure_membership, role_by_key
from asme.services.errors import Conflict, NotFound, Validation

ORG_AUDIT_FIELDS = ("name", "timezone", "academic_year_start_month", "logo_url")


def ensure_all_memberships(organization: Organization | None = None) -> int:
    """Backfill: every existing user gets a membership from their legacy role."""
    organization = organization or default_organization()
    authz.ensure_system_roles(organization)
    created = 0
    for user in User.query.order_by(User.id.asc()).all():
        if Membership.query.filter_by(organization_id=organization.id, user_id=user.id).first() is None:
            ensure_membership(user, organization)
            created += 1
    return created


def list_roles(organization: Organization):
    return Role.query.filter_by(organization_id=organization.id).order_by(Role.rank.desc(), Role.name.asc()).all()


def list_members(organization: Organization, include_inactive=False, include_invited=False):
    """Active members by default; ``include_invited`` adds pending invitations (Users page),
    ``include_inactive`` adds suspended memberships and disabled accounts (admins only)."""
    query = (
        Membership.query.join(User, User.id == Membership.user_id)
        .filter(Membership.organization_id == organization.id)
        .order_by(User.name.asc())
    )
    if not include_inactive:
        statuses = ["active", "invited"] if include_invited else ["active"]
        query = query.filter(Membership.member_status.in_(statuses), User.is_active.is_(True))
    return query.all()


def membership_for(organization: Organization, user_id: int) -> Membership:
    row = Membership.query.filter_by(organization_id=organization.id, user_id=user_id).first()
    if row is None:
        raise NotFound("Member not found.")
    return row


def serialize_member(membership: Membership) -> dict:
    user = membership.user
    payload = user_ref(user, membership)
    payload.update(
        {
            "membership_id": membership.id,
            "member_status": membership.member_status,
            "title": membership.title,
            "joined_at": membership.joined_at.isoformat() if membership.joined_at else None,
            "last_login_at": user.last_login_at.isoformat() if user.last_login_at else None,
            "is_active": bool(user.is_active),
        }
    )
    return payload


def update_profile(organization: Organization, data, fields_set: set[str], actor) -> Organization:
    before = audit.snapshot(organization, ORG_AUDIT_FIELDS)
    for key in fields_set:
        setattr(organization, key, getattr(data, key))
    after = audit.snapshot(organization, ORG_AUDIT_FIELDS)
    audit.record_event(
        "organization.updated", "organization", organization.id, organization_id=organization.id, actor=actor,
        before=before, after=after, metadata={"changed": audit.diff(before, after)},
    )
    db.session.commit()
    return organization


def set_setup_banner(membership: Membership, dismissed: bool):
    settings = membership.settings
    settings["setup_banner_dismissed"] = bool(dismissed)
    settings["setup_banner_updated_at"] = datetime.utcnow().isoformat()
    membership.settings = settings
    db.session.commit()
    return membership


def invite_user(organization: Organization, data, actor) -> Membership:
    role = role_by_key(organization, data.role_key)
    if role is None:
        raise Validation("Unknown role.", field="role_key")
    email = data.email.strip().lower()
    user = User.query.filter(func.lower(User.email) == email).first()
    if user and Membership.query.filter_by(organization_id=organization.id, user_id=user.id).first():
        raise Conflict("That person is already a member.", code="already_member")
    if user is None:
        from asme.services.identity import make_unique_username

        password = data.password or f"Invite-{organization.slug}-{datetime.utcnow().strftime('%y%m%d%H%M')}"
        user = User(
            name=data.name,
            email=email,
            username=make_unique_username(email.split("@", 1)[0]),
            password_hash=generate_password_hash(password),
            role=authz.legacy_role_for(role.system_key),
            is_active=True,
        )
        db.session.add(user)
        db.session.flush()
    membership = Membership(
        organization_id=organization.id,
        user_id=user.id,
        role_id=role.id,
        member_status="active" if data.password else "invited",
        title=data.title,
    )
    db.session.add(membership)
    db.session.flush()
    audit.record_event(
        "membership.created", "membership", membership.id, organization_id=organization.id, actor=actor,
        after={"user_id": user.id, "role": role.system_key, "status": membership.member_status},
    )
    db.session.commit()
    events.emit(events.USER_CREATED, user_id=user.id)
    return membership


def update_member(organization: Organization, membership: Membership, data, fields_set: set[str], actor) -> Membership:
    before = {"role": membership.role.system_key, "status": membership.member_status, "title": membership.title}
    if "role_key" in fields_set and data.role_key:
        role = role_by_key(organization, data.role_key)
        if role is None:
            raise Validation("Unknown role.", field="role_key")
        if membership.user_id == actor.id and role.system_key != "chapter_admin" and membership.role.system_key == "chapter_admin":
            others = (
                Membership.query.join(Role, Role.id == Membership.role_id)
                .filter(Membership.organization_id == organization.id, Role.system_key == "chapter_admin", Membership.user_id != actor.id, Membership.member_status == "active")
                .count()
            )
            if others == 0:
                raise Conflict("You are the last Chapter Administrator; promote someone else first.", code="last_admin")
        membership.role_id = role.id
        membership.role = role
        # keep the legacy portal's role in step
        membership.user.role = authz.legacy_role_for(role.system_key)
    if "member_status" in fields_set and data.member_status:
        membership.member_status = data.member_status
        membership.user.is_active = data.member_status != "suspended"
    if "title" in fields_set:
        membership.title = data.title
    after = {"role": membership.role.system_key, "status": membership.member_status, "title": membership.title}
    audit.record_event(
        "membership.updated", "membership", membership.id, organization_id=organization.id, actor=actor,
        before=before, after=after, metadata={"user_id": membership.user_id, "changed": audit.diff(before, after)},
    )
    db.session.commit()
    events.emit(events.USER_UPDATED, user_id=membership.user_id, role_changed=(before["role"] != after["role"]))
    return membership
