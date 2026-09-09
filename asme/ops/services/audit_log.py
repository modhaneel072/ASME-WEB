"""Read side of the audit history."""

from __future__ import annotations

import json

from asme.ops.models import AuditEvent
from asme.ops.services.common import user_ref


def serialize_event(row: AuditEvent) -> dict:
    def _load(value):
        try:
            return json.loads(value) if value else None
        except Exception:
            return None

    return {
        "id": row.id,
        "event_type": row.event_type,
        "entity_type": row.entity_type,
        "entity_id": row.entity_id,
        "actor": user_ref(row.actor),
        "before": _load(row.before_json),
        "after": _load(row.after_json),
        "metadata": _load(row.metadata_json) or {},
        "occurred_at": row.occurred_at.isoformat(),
    }


def list_events(ctx, *, entity_type=None, entity_id=None, event_type=None, actor_id=None, limit=100, offset=0):
    q = AuditEvent.query.filter(AuditEvent.organization_id == ctx.organization.id)
    if entity_type:
        q = q.filter(AuditEvent.entity_type == entity_type)
    if entity_id:
        q = q.filter(AuditEvent.entity_id == entity_id)
    if event_type:
        q = q.filter(AuditEvent.event_type.like(f"{event_type}%"))
    if actor_id:
        q = q.filter(AuditEvent.actor_user_id == actor_id)
    rows = q.order_by(AuditEvent.occurred_at.desc(), AuditEvent.id.desc()).offset(offset).limit(limit + 1).all()
    has_more = len(rows) > limit
    return rows[:limit], has_more
