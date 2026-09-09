"""Assets: create/edit, hierarchy, status changes with history."""

from __future__ import annotations

import json
import re
from datetime import datetime

from sqlalchemy import func

from asme.extensions import db
from asme.ops import audit
from asme.ops.models import Asset, AssetStatusHistory, AssetType, AssetTypeLink, Location, OpsProject, Team
from asme.ops.services.common import apply_patch, ref, resolve_users, touch, user_ref
from asme.ops.tenancy import get_or_404, scoped
from asme.services.errors import Conflict, Validation

ASSET_FIELDS = (
    "name", "code", "description", "parent_asset_id", "project_id", "location_id", "responsible_team_id", "owner_user_id",
    "manufacturer", "model", "serial_number", "purchase_date", "purchase_cost", "warranty_end", "criticality", "status",
)
DEFAULT_ASSET_TYPES = [
    ("Vehicle / Rover", "#0878d1", "truck"),
    ("Subsystem", "#475569", "component"),
    ("3D Printer", "#7c5ce7", "printer"),
    ("Soldering Station", "#e58a00", "flame"),
    ("Battery / Charger", "#00a878", "battery-charging"),
    ("Test Equipment", "#0891b2", "activity"),
    ("Tool Kit", "#6b7280", "wrench"),
    ("Computer", "#1d4ed8", "laptop"),
]


def serialize_asset(asset: Asset, child_counts: dict | None = None, open_wo_counts: dict | None = None) -> dict:
    return {
        "id": asset.id,
        "name": asset.name,
        "code": asset.code,
        "description": asset.description,
        "parent_asset_id": asset.parent_asset_id,
        "parent": ref(asset.parent),
        "project_id": asset.project_id,
        "project": ref(asset.project),
        "location_id": asset.location_id,
        "location": ref(asset.location),
        "responsible_team_id": asset.responsible_team_id,
        "responsible_team": ref(asset.responsible_team),
        "owner": user_ref(asset.owner),
        "manufacturer": asset.manufacturer,
        "model": asset.model,
        "serial_number": asset.serial_number,
        "purchase_date": asset.purchase_date.isoformat() if asset.purchase_date else None,
        "purchase_cost": float(asset.purchase_cost) if asset.purchase_cost is not None else None,
        "warranty_end": asset.warranty_end.isoformat() if asset.warranty_end else None,
        "criticality": asset.criticality,
        "status": asset.status,
        "qr_code": asset.qr_code,
        "asset_types": [{"id": l.asset_type.id, "name": l.asset_type.name, "color": l.asset_type.color, "icon": l.asset_type.icon} for l in asset.type_links],
        "custom_fields": asset.custom_fields,
        "child_count": (child_counts or {}).get(asset.id, 0),
        "open_work_orders": (open_wo_counts or {}).get(asset.id, 0),
        "archived_at": asset.archived_at.isoformat() if asset.archived_at else None,
        "created_at": asset.created_at.isoformat(),
        "updated_at": asset.updated_at.isoformat(),
    }


def serialize_status_history(row: AssetStatusHistory) -> dict:
    return {
        "id": row.id,
        "from_status": row.from_status,
        "to_status": row.to_status,
        "downtime_type": row.downtime_type,
        "downtime_reason": row.downtime_reason,
        "note": row.note,
        "started_at": row.started_at.isoformat() if row.started_at else None,
        "ended_at": row.ended_at.isoformat() if row.ended_at else None,
        "changed_by": user_ref(row.changed_by),
    }


def list_asset_types(ctx):
    return scoped(AssetType, ctx).order_by(AssetType.name.asc()).all()


def ensure_default_asset_types(org_id: str) -> int:
    existing = {t.name.lower() for t in AssetType.query.filter_by(organization_id=org_id).all()}
    created = 0
    for name, color, icon in DEFAULT_ASSET_TYPES:
        if name.lower() not in existing:
            db.session.add(AssetType(organization_id=org_id, name=name, color=color, icon=icon))
            created += 1
    db.session.flush()
    return created


def next_code(ctx, name: str) -> str:
    base = re.sub(r"[^A-Z0-9]+", "-", (name or "ASSET").upper()).strip("-")[:20] or "ASSET"
    candidate = base
    n = 2
    while scoped(Asset, ctx).filter(func.lower(Asset.code) == candidate.lower()).first():
        candidate = f"{base}-{n}"
        n += 1
    return candidate


def _validate_refs(ctx, *, parent_asset_id=None, project_id=None, location_id=None, responsible_team_id=None, owner_user_id=None, asset_id=None):
    if parent_asset_id:
        parent = get_or_404(Asset, parent_asset_id, ctx, "Parent asset")
        if asset_id and parent.id == asset_id:
            raise Validation("An asset cannot be its own parent.", field="parent_asset_id")
        if asset_id:
            cursor = parent
            depth = 0
            while cursor is not None and depth < 50:
                if cursor.parent_asset_id == asset_id:
                    raise Validation("That parent would create a cycle in the hierarchy.", field="parent_asset_id")
                cursor = cursor.parent
                depth += 1
    if project_id:
        get_or_404(OpsProject, project_id, ctx, "Project")
    if location_id:
        get_or_404(Location, location_id, ctx, "Location")
    if responsible_team_id:
        get_or_404(Team, responsible_team_id, ctx, "Team")
    if owner_user_id and owner_user_id not in resolve_users(ctx.organization.id, [owner_user_id]):
        raise Validation("Owner must be an active member.", field="owner_user_id")


def _set_types(ctx, asset: Asset, type_ids):
    wanted = set(type_ids or [])
    for tid in wanted:
        get_or_404(AssetType, tid, ctx, "Asset type")
    existing = {l.asset_type_id: l for l in asset.type_links}
    for tid, link in existing.items():
        if tid not in wanted:
            db.session.delete(link)
    for tid in wanted:
        if tid not in existing:
            db.session.add(AssetTypeLink(asset_id=asset.id, asset_type_id=tid))
    db.session.flush()


def create_asset(ctx, data) -> Asset:
    from asme.ops.services.directory import default_location

    _validate_refs(
        ctx, parent_asset_id=data.parent_asset_id, project_id=data.project_id, location_id=data.location_id,
        responsible_team_id=data.responsible_team_id, owner_user_id=data.owner_user_id,
    )
    code = (data.code or "").strip().upper() or next_code(ctx, data.name)
    if scoped(Asset, ctx).filter(func.lower(Asset.code) == code.lower()).first():
        raise Conflict("An asset with that code already exists.", code="code_taken")
    asset = Asset(
        organization_id=ctx.organization.id,
        name=data.name,
        code=code,
        description=data.description,
        parent_asset_id=data.parent_asset_id,
        project_id=data.project_id,
        location_id=data.location_id or default_location(ctx).id,
        responsible_team_id=data.responsible_team_id,
        owner_user_id=data.owner_user_id,
        manufacturer=data.manufacturer,
        model=data.model,
        serial_number=data.serial_number,
        purchase_date=data.purchase_date,
        purchase_cost=data.purchase_cost,
        warranty_end=data.warranty_end,
        criticality=data.criticality,
        status=data.status,
        custom_fields_json=json.dumps(data.custom_fields or {}),
        created_by_user_id=ctx.user.id,
        updated_by_user_id=ctx.user.id,
    )
    db.session.add(asset)
    db.session.flush()
    _set_types(ctx, asset, data.asset_type_ids)
    db.session.add(AssetStatusHistory(asset_id=asset.id, from_status=None, to_status=asset.status, changed_by_user_id=ctx.user.id))
    audit.record_event("asset.created", "asset", asset.id, organization_id=ctx.organization.id, actor=ctx.user, after=audit.snapshot(asset, ASSET_FIELDS))
    db.session.commit()
    from asme import events

    events.emit(events.CHAPTER_CHANGED, reason="asset_created")
    return asset


def update_asset(ctx, asset: Asset, data, fields_set: set[str]) -> Asset:
    _validate_refs(
        ctx,
        parent_asset_id=getattr(data, "parent_asset_id", None) if "parent_asset_id" in fields_set else None,
        project_id=getattr(data, "project_id", None) if "project_id" in fields_set else None,
        location_id=getattr(data, "location_id", None) if "location_id" in fields_set else None,
        responsible_team_id=getattr(data, "responsible_team_id", None) if "responsible_team_id" in fields_set else None,
        owner_user_id=getattr(data, "owner_user_id", None) if "owner_user_id" in fields_set else None,
        asset_id=asset.id,
    )
    if "code" in fields_set and data.code:
        code = data.code.strip().upper()
        if scoped(Asset, ctx).filter(func.lower(Asset.code) == code.lower(), Asset.id != asset.id).first():
            raise Conflict("An asset with that code already exists.", code="code_taken")
        data.code = code
    before = audit.snapshot(asset, ASSET_FIELDS)
    apply_patch(asset, data, fields_set, set(ASSET_FIELDS) - {"status"})
    if "custom_fields" in fields_set and data.custom_fields is not None:
        asset.custom_fields_json = json.dumps(data.custom_fields)
    if "asset_type_ids" in fields_set and data.asset_type_ids is not None:
        _set_types(ctx, asset, data.asset_type_ids)
    touch(asset, ctx.user)
    after = audit.snapshot(asset, ASSET_FIELDS)
    audit.record_event("asset.updated", "asset", asset.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"changed": audit.diff(before, after)})
    db.session.commit()
    return asset


def change_status(ctx, asset: Asset, data) -> Asset:
    if data.status == asset.status:
        return asset
    if data.status in {"OFFLINE_PLANNED", "OFFLINE_UNPLANNED"} and not data.downtime_type:
        data.downtime_type = "planned" if data.status == "OFFLINE_PLANNED" else "unplanned"
    now = datetime.utcnow()
    open_row = AssetStatusHistory.query.filter_by(asset_id=asset.id, ended_at=None).order_by(AssetStatusHistory.started_at.desc()).first()
    if open_row:
        open_row.ended_at = now
    from_status = asset.status
    asset.status = data.status
    touch(asset, ctx.user)
    db.session.add(
        AssetStatusHistory(
            asset_id=asset.id, from_status=from_status, to_status=data.status, downtime_type=data.downtime_type,
            downtime_reason=data.downtime_reason, note=data.note, started_at=now, changed_by_user_id=ctx.user.id,
        )
    )
    audit.record_event(
        "asset.status_changed", "asset", asset.id, organization_id=ctx.organization.id, actor=ctx.user,
        before={"status": from_status}, after={"status": data.status}, metadata={"downtime_type": data.downtime_type, "reason": data.downtime_reason},
    )
    db.session.commit()
    return asset


def child_counts(ctx) -> dict:
    rows = db.session.query(Asset.parent_asset_id, func.count(Asset.id)).filter(Asset.organization_id == ctx.organization.id, Asset.parent_asset_id.isnot(None)).group_by(Asset.parent_asset_id).all()
    return {pid: int(n) for pid, n in rows}


def open_work_order_counts(ctx) -> dict:
    from asme.ops.models import WO_OPEN_STATUSES, WorkOrder

    rows = (
        db.session.query(WorkOrder.primary_asset_id, func.count(WorkOrder.id))
        .filter(WorkOrder.organization_id == ctx.organization.id, WorkOrder.status.in_(WO_OPEN_STATUSES), WorkOrder.primary_asset_id.isnot(None))
        .group_by(WorkOrder.primary_asset_id)
        .all()
    )
    return {aid: int(n) for aid, n in rows}


def list_assets(ctx, query):
    q = scoped(Asset, ctx).filter(Asset.archived_at.is_(None))
    filters = query.filters
    if "status" in filters:
        q = q.filter(Asset.status.in_(filters["status"]))
    if "criticality" in filters:
        q = q.filter(Asset.criticality.in_(filters["criticality"]))
    if "project" in filters:
        q = q.filter(Asset.project_id.in_(filters["project"]))
    if "location" in filters:
        q = q.filter(Asset.location_id.in_(filters["location"]))
    if "team" in filters:
        q = q.filter(Asset.responsible_team_id.in_(filters["team"]))
    if "parent" in filters:
        values = filters["parent"]
        q = q.filter(Asset.parent_asset_id.is_(None)) if values == ["none"] else q.filter(Asset.parent_asset_id.in_(values))
    if "type" in filters:
        q = q.join(AssetTypeLink, AssetTypeLink.asset_id == Asset.id).filter(AssetTypeLink.asset_type_id.in_(filters["type"]))
    if query.q:
        like = f"%{query.q.lower()}%"
        q = q.filter(func.lower(Asset.name).like(like) | func.lower(func.coalesce(Asset.code, "")).like(like) | func.lower(func.coalesce(Asset.serial_number, "")).like(like))
    sort = query.sort
    order = {"name_asc": (Asset.name.asc(), Asset.id.asc()), "updated_desc": (Asset.updated_at.desc(), Asset.id.desc()), "status": (Asset.status.asc(), Asset.name.asc())}
    q = q.order_by(*order.get(sort, order["name_asc"]))
    offset = int((query.cursor or {}).get("offset", 0))
    rows = q.offset(offset).limit(query.limit + 1).all()
    has_more = len(rows) > query.limit
    return rows[: query.limit], ({"offset": offset + query.limit} if has_more else None)
