"""Saved filters ("My Filters"): private, team-shared or chapter-shared views."""

from __future__ import annotations

import json

from asme.extensions import db
from asme.ops import audit, authz
from asme.ops.models import SavedFilter, Team
from asme.ops.services.common import user_ref
from asme.ops.tenancy import get_or_404
from asme.services.errors import Forbidden, Validation


def serialize(row: SavedFilter, ctx) -> dict:
    return {
        "id": row.id,
        "entity_type": row.entity_type,
        "name": row.name,
        "visibility": row.visibility,
        "team_id": row.team_id,
        "filters": json.loads(row.filter_json or "{}"),
        "sort": (json.loads(row.sort_json or "{}") or {}).get("sort"),
        "view_type": row.view_type,
        "is_default": bool(row.is_default),
        "owner": user_ref(row.owner),
        "is_mine": row.owner_user_id == ctx.user.id,
        "created_at": row.created_at.isoformat(),
    }


def list_for(ctx, entity_type: str):
    team_ids = ctx.member_team_ids()
    q = SavedFilter.query.filter(SavedFilter.organization_id == ctx.organization.id, SavedFilter.entity_type == entity_type).filter(
        (SavedFilter.owner_user_id == ctx.user.id)
        | (SavedFilter.visibility == "chapter")
        | ((SavedFilter.visibility == "team") & (SavedFilter.team_id.in_(team_ids or ["-"])))
    )
    rows = q.order_by(SavedFilter.name.asc()).all()
    mine = [r for r in rows if r.owner_user_id == ctx.user.id]
    shared = [r for r in rows if r.owner_user_id != ctx.user.id]
    return mine, shared


def _check_share(ctx, visibility: str, team_id: str | None):
    if visibility != "private" and not authz.can(ctx, "saved_filter.share"):
        raise Forbidden("You cannot share filters.", code="forbidden")
    if visibility == "team":
        if not team_id:
            raise Validation("Pick a team to share with.", field="team_id")
        get_or_404(Team, team_id, ctx, "Team")


def create(ctx, data) -> SavedFilter:
    _check_share(ctx, data.visibility, data.team_id)
    if data.is_default:
        SavedFilter.query.filter_by(organization_id=ctx.organization.id, owner_user_id=ctx.user.id, entity_type=data.entity_type, is_default=True).update({SavedFilter.is_default: False}, synchronize_session=False)
    row = SavedFilter(
        organization_id=ctx.organization.id,
        entity_type=data.entity_type,
        owner_user_id=ctx.user.id,
        name=data.name,
        visibility=data.visibility,
        team_id=data.team_id if data.visibility == "team" else None,
        filter_json=json.dumps(data.filters or {}),
        sort_json=json.dumps({"sort": data.sort} if data.sort else {}),
        view_type=data.view_type,
        is_default=data.is_default,
        created_by_user_id=ctx.user.id,
        updated_by_user_id=ctx.user.id,
    )
    db.session.add(row)
    db.session.flush()
    audit.record_event("saved_filter.created", "saved_filter", row.id, organization_id=ctx.organization.id, actor=ctx.user, after={"name": row.name, "visibility": row.visibility})
    db.session.commit()
    return row


def update(ctx, row: SavedFilter, data, fields_set: set[str]) -> SavedFilter:
    if row.owner_user_id != ctx.user.id and not ctx.is_admin:
        raise Forbidden("You can only edit your own filters.", code="not_owner")
    visibility = data.visibility if "visibility" in fields_set and data.visibility else row.visibility
    team_id = data.team_id if "team_id" in fields_set else row.team_id
    _check_share(ctx, visibility, team_id)
    if "name" in fields_set and data.name:
        row.name = data.name
    row.visibility = visibility
    row.team_id = team_id if visibility == "team" else None
    if "filters" in fields_set and data.filters is not None:
        row.filter_json = json.dumps(data.filters)
    if "sort" in fields_set:
        row.sort_json = json.dumps({"sort": data.sort} if data.sort else {})
    if "view_type" in fields_set and data.view_type:
        row.view_type = data.view_type
    if "is_default" in fields_set and data.is_default is not None:
        if data.is_default:
            SavedFilter.query.filter(SavedFilter.organization_id == ctx.organization.id, SavedFilter.owner_user_id == ctx.user.id, SavedFilter.entity_type == row.entity_type, SavedFilter.id != row.id).update({SavedFilter.is_default: False}, synchronize_session=False)
        row.is_default = data.is_default
    row.updated_by_user_id = ctx.user.id
    db.session.commit()
    return row


def remove(ctx, row: SavedFilter):
    if row.owner_user_id != ctx.user.id and not ctx.is_admin:
        raise Forbidden("You can only delete your own filters.", code="not_owner")
    audit.record_event("saved_filter.deleted", "saved_filter", row.id, organization_id=ctx.organization.id, actor=ctx.user, before={"name": row.name})
    db.session.delete(row)
    db.session.commit()
