"""Entitlements: what a user has *proven*, orthogonal to their role."""

from __future__ import annotations

from datetime import datetime

from asme.auth.session import role_allows
from asme.constants import ALL_ENTITLEMENTS
from asme.extensions import db
from asme.models import Entitlement, Phase
from asme.services import audit


def active_rows(user):
    rows = Entitlement.query.filter(Entitlement.user_id == user.id, Entitlement.revoked_at.is_(None)).all()
    return [row for row in rows if row.is_active]


def user_entitlements(user) -> set[str]:
    if not user:
        return set()
    keys = {row.key for row in active_rows(user)}
    if role_allows(user.role, "admin"):
        keys.update(ALL_ENTITLEMENTS)
    return keys


def has_entitlement(user, key: str) -> bool:
    if not user:
        return False
    if role_allows(user.role, "admin"):
        return True
    return any(row.key == key for row in active_rows(user))


def grant(user, key, *, source="phase", source_ref=None, granted_by=None, expires_at=None, reason=None, commit=False):
    """Idempotent: an active grant for ``(user, key, source)`` is returned as-is."""
    existing = (
        Entitlement.query.filter(
            Entitlement.user_id == user.id,
            Entitlement.key == key,
            Entitlement.source == source,
            Entitlement.revoked_at.is_(None),
        )
        .order_by(Entitlement.id.desc())
        .first()
    )
    if existing and existing.is_active:
        if source_ref and existing.source_ref != source_ref:
            existing.source_ref = source_ref
        return existing, False
    row = Entitlement(
        user_id=user.id,
        key=key,
        source=source,
        source_ref=source_ref,
        granted_by_user_id=granted_by.id if granted_by else None,
        expires_at=expires_at,
        reason=(reason or "")[:300] or None,
    )
    db.session.add(row)
    audit.record("grant_entitlement", f"user_id={user.id} key={key} source={source} ref={source_ref or ''}", actor=granted_by)
    if commit:
        db.session.commit()
    return row, True


def revoke(user, key, *, source=None, revoked_by=None, reason=None, commit=False) -> int:
    query = Entitlement.query.filter(Entitlement.user_id == user.id, Entitlement.key == key, Entitlement.revoked_at.is_(None))
    if source:
        query = query.filter(Entitlement.source == source)
    count = 0
    for row in query.all():
        row.revoked_at = datetime.utcnow()
        row.revoked_by_user_id = revoked_by.id if revoked_by else None
        if reason:
            row.reason = reason[:300]
        count += 1
    if count:
        audit.record("revoke_entitlement", f"user_id={user.id} key={key} source={source or '*'} count={count}", actor=revoked_by)
    if commit:
        db.session.commit()
    return count


def revoke_phase_grants(user, source_ref, commit=False) -> int:
    rows = Entitlement.query.filter(
        Entitlement.user_id == user.id,
        Entitlement.source == "phase",
        Entitlement.source_ref == source_ref,
        Entitlement.revoked_at.is_(None),
    ).all()
    for row in rows:
        row.revoked_at = datetime.utcnow()
        row.reason = "phase requirements no longer met"
    if commit and rows:
        db.session.commit()
    return len(rows)


def phase_that_grants(key: str):
    """First phase (lowest order, any track) whose grants include ``key``."""
    for phase in Phase.query.order_by(Phase.order.asc(), Phase.id.asc()).all():
        if key in phase.grants:
            return phase
    return None
