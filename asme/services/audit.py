"""Append-only audit trail. Services call ``record`` inside the transaction that
performs the privileged action, so the log commits with it."""

from __future__ import annotations

from flask import has_request_context, request

from asme.extensions import db
from asme.models import AuditLog


def record(action: str, details: str = "", actor=None) -> AuditLog:
    if actor is None and has_request_context():
        from asme.auth.session import current_auth_user

        actor = current_auth_user()
    ip_address = None
    if has_request_context():
        ip_address = (request.remote_addr or "")[:120] or None
    log = AuditLog(
        admin_user_id=actor.id if actor else None,
        action=(action or "").strip()[:160] or "action",
        details=(details or "").strip()[:4000] or None,
        ip_address=ip_address,
    )
    db.session.add(log)
    return log


def recent(limit: int = 150):
    return AuditLog.query.order_by(AuditLog.created_at.desc(), AuditLog.id.desc()).limit(limit).all()
