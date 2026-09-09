"""Saved filters."""

from __future__ import annotations

from flask import request

from asme.ops.api import body, bp, contract, ctx, ok
from asme.ops.models import SavedFilter
from asme.ops.schemas import SavedFilterCreate, SavedFilterUpdate
from asme.ops.services import saved_filters
from asme.services.errors import NotFound, Validation


@bp.get("/saved-filters")
def saved_filters_list():
    entity_type = (request.args.get("entity_type") or "work_order").strip()
    if entity_type not in {"work_order", "project", "asset"}:
        raise Validation("Unknown entity type.", field="entity_type")
    mine, shared = saved_filters.list_for(ctx(), entity_type)
    return ok({"mine": [saved_filters.serialize(r, ctx()) for r in mine], "shared": [saved_filters.serialize(r, ctx()) for r in shared]})


@bp.post("/saved-filters")
@contract(SavedFilterCreate, summary="Save the current filters as a view", tags=("saved-filters",))
def saved_filters_create():
    data = body(SavedFilterCreate)
    return ok(saved_filters.serialize(saved_filters.create(ctx(), data), ctx()), status=201)


def _row(saved_filter_id):
    row = SavedFilter.query.filter_by(id=saved_filter_id, organization_id=ctx().organization.id).first()
    if row is None:
        raise NotFound("Saved filter not found.")
    return row


@bp.patch("/saved-filters/<saved_filter_id>")
@contract(SavedFilterUpdate, summary="Rename, share or change a saved view", tags=("saved-filters",))
def saved_filters_update(saved_filter_id):
    row = _row(saved_filter_id)
    data = body(SavedFilterUpdate)
    return ok(saved_filters.serialize(saved_filters.update(ctx(), row, data, data.model_fields_set), ctx()))


@bp.delete("/saved-filters/<saved_filter_id>")
def saved_filters_delete(saved_filter_id):
    row = _row(saved_filter_id)
    saved_filters.remove(ctx(), row)
    return ok({"deleted": True})
