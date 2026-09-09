"""Immutable audit events for ops actions. Written inside the caller's transaction."""

from __future__ import annotations

import json
from datetime import date, datetime
from decimal import Decimal

from flask import g, has_request_context, request

from asme.extensions import db
from asme.ops.models import AuditEvent


def _jsonable(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, set):
        return sorted(value)
    return value


def snapshot(record, fields: tuple[str, ...] | list[str]) -> dict:
    return {name: _jsonable(getattr(record, name, None)) for name in fields}


def record_event(
    event_type: str,
    entity_type: str,
    entity_id: str,
    *,
    organization_id: str,
    actor=None,
    before: dict | None = None,
    after: dict | None = None,
    metadata: dict | None = None,
) -> AuditEvent:
    if actor is None and has_request_context():
        ctx = getattr(g, "ops_ctx", None)
        actor = ctx.user if ctx else None
    row = AuditEvent(
        organization_id=organization_id,
        actor_user_id=actor.id if actor else None,
        event_type=event_type[:80],
        entity_type=entity_type[:40],
        entity_id=str(entity_id)[:36],
        before_json=json.dumps(before, default=_jsonable) if before is not None else None,
        after_json=json.dumps(after, default=_jsonable) if after is not None else None,
        metadata_json=json.dumps(metadata or {}, default=_jsonable),
        request_id=(getattr(g, "request_id", None) if has_request_context() else None),
        ip_address=((request.remote_addr or "")[:120] or None) if has_request_context() else None,
    )
    db.session.add(row)
    return row


def diff(before: dict, after: dict) -> dict:
    """Only the keys whose value changed - keeps the audit rows small."""
    changed = {}
    for key in set(before) | set(after):
        if before.get(key) != after.get(key):
            changed[key] = {"from": before.get(key), "to": after.get(key)}
    return changed
