"""Small helpers shared by ops services."""

from __future__ import annotations

from datetime import datetime

from asme.models import User
from asme.ops.models import Membership


def user_ref(user, membership: Membership | None = None) -> dict | None:
    if user is None:
        return None
    parts = [p for p in (user.name or "").split() if p]
    initials = "".join(p[0].upper() for p in parts[:2]) or (user.email or "?")[0].upper()
    payload = {"id": user.id, "name": user.name, "email": user.email, "initials": initials}
    role = membership.role if membership else None
    payload["role_key"] = role.system_key if role else None
    payload["role_name"] = role.name if role else None
    return payload


def ref(obj) -> dict | None:
    if obj is None:
        return None
    return {"id": obj.id, "name": getattr(obj, "name", None) or getattr(obj, "title", "")}


def iso(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.isoformat()
    return value.isoformat() if hasattr(value, "isoformat") else str(value)


def resolve_users(org_id: str, user_ids) -> dict[int, User]:
    """Users that are active members of the organisation, keyed by id."""
    ids = {int(v) for v in (user_ids or []) if v is not None}
    if not ids:
        return {}
    rows = (
        User.query.join(Membership, Membership.user_id == User.id)
        .filter(Membership.organization_id == org_id, Membership.member_status == "active", User.id.in_(ids))
        .all()
    )
    return {u.id: u for u in rows}


def touch(record, actor):
    if actor is not None and hasattr(record, "updated_by_user_id"):
        record.updated_by_user_id = actor.id


def apply_patch(record, data, fields_set: set[str], allowed: set[str]):
    """Copy only the keys the client actually sent (``PatchModel`` semantics)."""
    for key in fields_set:
        if key in allowed:
            setattr(record, key, getattr(data, key))
