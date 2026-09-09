"""Comments on any ops entity."""

from __future__ import annotations

import re
from datetime import datetime

from asme.extensions import db
from asme.ops import audit
from asme.ops.models import Comment
from asme.ops.services.common import user_ref
from asme.services.errors import Forbidden, NotFound

MENTION_RE = re.compile(r"@\[(\d+)\]")


def serialize(comment: Comment) -> dict:
    return {
        "id": comment.id,
        "entity_type": comment.entity_type,
        "entity_id": comment.entity_id,
        "body": comment.body if not comment.deleted_at else "",
        "deleted": bool(comment.deleted_at),
        "author": user_ref(comment.author),
        "parent_comment_id": comment.parent_comment_id,
        "created_at": comment.created_at.isoformat(),
        "edited_at": comment.edited_at.isoformat() if comment.edited_at else None,
    }


def list_for(ctx, entity_type: str, entity_id: str):
    return (
        Comment.query.filter(Comment.organization_id == ctx.organization.id, Comment.entity_type == entity_type, Comment.entity_id == entity_id)
        .order_by(Comment.created_at.asc())
        .all()
    )


def count_for(ctx, entity_type: str, entity_ids: list[str]) -> dict:
    from sqlalchemy import func

    if not entity_ids:
        return {}
    rows = (
        db.session.query(Comment.entity_id, func.count(Comment.id))
        .filter(Comment.organization_id == ctx.organization.id, Comment.entity_type == entity_type, Comment.entity_id.in_(entity_ids), Comment.deleted_at.is_(None))
        .group_by(Comment.entity_id)
        .all()
    )
    return {eid: int(n) for eid, n in rows}


def mentioned_user_ids(body: str) -> set[int]:
    return {int(m) for m in MENTION_RE.findall(body or "")}


def create(ctx, entity_type: str, entity_id: str, body: str, parent_comment_id: str | None = None) -> Comment:
    if parent_comment_id:
        parent = Comment.query.filter_by(id=parent_comment_id, organization_id=ctx.organization.id, entity_type=entity_type, entity_id=entity_id).first()
        if parent is None:
            raise NotFound("Parent comment not found.")
    comment = Comment(
        organization_id=ctx.organization.id,
        entity_type=entity_type,
        entity_id=entity_id,
        author_user_id=ctx.user.id,
        body=body,
        parent_comment_id=parent_comment_id,
        created_by_user_id=ctx.user.id,
        updated_by_user_id=ctx.user.id,
    )
    db.session.add(comment)
    db.session.flush()
    audit.record_event("comment.created", entity_type, entity_id, organization_id=ctx.organization.id, actor=ctx.user, after={"comment_id": comment.id, "excerpt": body[:120]})
    return comment


def edit(ctx, comment: Comment, body: str) -> Comment:
    if comment.author_user_id != ctx.user.id and not ctx.is_admin:
        raise Forbidden("You can only edit your own comments.", code="not_author")
    comment.body = body
    comment.edited_at = datetime.utcnow()
    comment.updated_by_user_id = ctx.user.id
    audit.record_event("comment.edited", comment.entity_type, comment.entity_id, organization_id=ctx.organization.id, actor=ctx.user, metadata={"comment_id": comment.id})
    db.session.commit()
    return comment


def remove(ctx, comment: Comment) -> Comment:
    if comment.author_user_id != ctx.user.id and not ctx.is_admin:
        raise Forbidden("You can only delete your own comments.", code="not_author")
    comment.deleted_at = datetime.utcnow()
    comment.updated_by_user_id = ctx.user.id
    audit.record_event("comment.deleted", comment.entity_type, comment.entity_id, organization_id=ctx.organization.id, actor=ctx.user, metadata={"comment_id": comment.id})
    db.session.commit()
    return comment
