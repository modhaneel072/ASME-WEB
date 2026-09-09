"""Assets."""

from __future__ import annotations

from asme.ops import authz
from asme.ops.api import body, bp, contract, ctx, ok
from asme.ops.api.query import encode_cursor, parse_list_query
from asme.ops.models import Asset, AssetStatusHistory
from asme.ops.schemas import AssetCreate, AssetStatusChange, AssetUpdate
from asme.ops.services import assets as assets_service
from asme.ops.services import files
from asme.ops.tenancy import get_or_404

ASSET_FILTERS = {"status", "criticality", "project", "location", "team", "parent", "type"}
ASSET_SORTS = {"name_asc", "updated_desc", "status"}


@bp.get("/assets")
@authz.permission_required("asset.read")
def assets_list():
    query = parse_list_query(allowed_filters=ASSET_FILTERS, allowed_sorts=ASSET_SORTS, default_sort="name_asc")
    rows, next_cursor = assets_service.list_assets(ctx(), query)
    child_counts = assets_service.child_counts(ctx())
    wo_counts = assets_service.open_work_order_counts(ctx())
    return ok({"items": [assets_service.serialize_asset(a, child_counts, wo_counts) for a in rows], "next_cursor": encode_cursor(next_cursor) if next_cursor else None})


@bp.get("/asset-types")
@authz.permission_required("asset.read")
def asset_types_list():
    return ok([{"id": t.id, "name": t.name, "color": t.color, "icon": t.icon} for t in assets_service.list_asset_types(ctx())])


@bp.post("/assets")
@contract(AssetCreate, summary="Create an asset", tags=("assets",))
def assets_create():
    data = body(AssetCreate)
    authz.require_create(ctx(), "asset.manage", project_id=data.project_id, team_id=data.responsible_team_id)
    asset = assets_service.create_asset(ctx(), data)
    return ok(assets_service.serialize_asset(asset), status=201)


@bp.get("/assets/<asset_id>")
@authz.permission_required("asset.read")
def assets_get(asset_id):
    asset = get_or_404(Asset, asset_id, ctx(), "Asset")
    child_counts = assets_service.child_counts(ctx())
    wo_counts = assets_service.open_work_order_counts(ctx())
    payload = assets_service.serialize_asset(asset, child_counts, wo_counts)
    payload["children"] = [assets_service.serialize_asset(c, child_counts, wo_counts) for c in Asset.query.filter_by(parent_asset_id=asset.id, organization_id=ctx().organization.id).order_by(Asset.name).all()]
    payload["files"] = [files.serialize(f) for f in files.list_for(ctx(), "asset", asset.id)]
    payload["permissions"] = {"edit": authz.can(ctx(), "asset.manage", record=asset)}
    return ok(payload)


@bp.patch("/assets/<asset_id>")
@contract(AssetUpdate, summary="Update an asset", tags=("assets",))
def assets_update(asset_id):
    asset = get_or_404(Asset, asset_id, ctx(), "Asset")
    authz.require(ctx(), "asset.manage", record=asset)
    data = body(AssetUpdate)
    asset = assets_service.update_asset(ctx(), asset, data, data.model_fields_set)
    return ok(assets_service.serialize_asset(asset))


@bp.post("/assets/<asset_id>/status")
@contract(AssetStatusChange, summary="Change an asset's status", tags=("assets",))
def assets_status(asset_id):
    asset = get_or_404(Asset, asset_id, ctx(), "Asset")
    authz.require(ctx(), "asset.manage", record=asset)
    data = body(AssetStatusChange)
    asset = assets_service.change_status(ctx(), asset, data)
    return ok(assets_service.serialize_asset(asset))


@bp.get("/assets/<asset_id>/history")
@authz.permission_required("asset.read")
def assets_history(asset_id):
    asset = get_or_404(Asset, asset_id, ctx(), "Asset")
    rows = AssetStatusHistory.query.filter_by(asset_id=asset.id).order_by(AssetStatusHistory.started_at.desc()).all()
    return ok([assets_service.serialize_status_history(r) for r in rows])
