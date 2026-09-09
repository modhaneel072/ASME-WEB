"""In-app notifications (email delivery goes through the outbox when SMTP is configured)."""

from __future__ import annotations

from datetime import datetime

from asme.extensions import db
from asme.ops.models import Notification


def notify(org_id: str, user_ids, *, type: str, title: str, body: str | None = None, entity_type=None, entity_id=None, exclude_user_id=None):
    rows = []
    for user_id in {int(u) for u in user_ids if u is not None}:
        if exclude_user_id and user_id == exclude_user_id:
            continue
        row = Notification(organization_id=org_id, user_id=user_id, type=type, title=title[:220], body=body, entity_type=entity_type, entity_id=entity_id)
        db.session.add(row)
        rows.append(row)
    return rows


def serialize(row: Notification) -> dict:
    return {
        "id": row.id,
        "type": row.type,
        "title": row.title,
        "body": row.body,
        "entity_type": row.entity_type,
        "entity_id": row.entity_id,
        "read_at": row.read_at.isoformat() if row.read_at else None,
        "created_at": row.created_at.isoformat(),
    }


def list_for(ctx, unread_only=False, limit=50):
    q = Notification.query.filter(Notification.organization_id == ctx.organization.id, Notification.user_id == ctx.user.id)
    if unread_only:
        q = q.filter(Notification.read_at.is_(None))
    return q.order_by(Notification.created_at.desc()).limit(limit).all()


def unread_count(ctx) -> int:
    return Notification.query.filter(Notification.organization_id == ctx.organization.id, Notification.user_id == ctx.user.id, Notification.read_at.is_(None)).count()


def mark_read(ctx, notification_id: str | None = None) -> int:
    q = Notification.query.filter(Notification.organization_id == ctx.organization.id, Notification.user_id == ctx.user.id, Notification.read_at.is_(None))
    if notification_id:
        q = q.filter(Notification.id == notification_id)
    count = q.update({Notification.read_at: datetime.utcnow()}, synchronize_session=False)
    db.session.commit()
    return count
