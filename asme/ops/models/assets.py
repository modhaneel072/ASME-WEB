"""Assets: hierarchy, types, status history."""

from __future__ import annotations

import json

from asme.extensions import db
from asme.ops.models.base import OpsBase, new_uuid, utcnow

ASSET_STATUSES = ("ONLINE", "OFFLINE_PLANNED", "OFFLINE_UNPLANNED", "DO_NOT_TRACK", "RETIRED")
ASSET_CRITICALITIES = ("low", "medium", "high", "critical")


class Asset(OpsBase, db.Model):
    __tablename__ = "assets"
    __table_args__ = (db.UniqueConstraint("organization_id", "code", name="uq_assets_org_code"),)
    name = db.Column(db.String(200), nullable=False)
    code = db.Column(db.String(60), nullable=True)
    description = db.Column(db.Text, nullable=True)
    parent_asset_id = db.Column(db.String(36), db.ForeignKey("assets.id"), nullable=True, index=True)
    project_id = db.Column(db.String(36), db.ForeignKey("ops_projects.id"), nullable=True, index=True)
    location_id = db.Column(db.String(36), db.ForeignKey("locations.id"), nullable=True, index=True)
    responsible_team_id = db.Column(db.String(36), db.ForeignKey("teams.id"), nullable=True, index=True)
    owner_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    manufacturer = db.Column(db.String(160), nullable=True)
    model = db.Column(db.String(160), nullable=True)
    serial_number = db.Column(db.String(160), nullable=True)
    purchase_date = db.Column(db.Date, nullable=True)
    purchase_cost = db.Column(db.Numeric(12, 2), nullable=True)
    warranty_end = db.Column(db.Date, nullable=True)
    criticality = db.Column(db.String(20), nullable=False, default="medium")
    status = db.Column(db.String(30), nullable=False, default="ONLINE", index=True)
    qr_code = db.Column(db.String(120), nullable=True)
    photo_attachment_id = db.Column(db.String(36), nullable=True)
    custom_fields_json = db.Column(db.Text, nullable=False, default="{}")
    archived_at = db.Column(db.DateTime, nullable=True)

    parent = db.relationship("Asset", remote_side="Asset.id", foreign_keys=[parent_asset_id])
    project = db.relationship("OpsProject", foreign_keys=[project_id])
    location = db.relationship("Location", foreign_keys=[location_id])
    responsible_team = db.relationship("Team", foreign_keys=[responsible_team_id])
    owner = db.relationship("User", foreign_keys=[owner_user_id])
    type_links = db.relationship("AssetTypeLink", back_populates="asset", cascade="all, delete-orphan")

    @property
    def custom_fields(self) -> dict:
        try:
            value = json.loads(self.custom_fields_json or "{}")
        except Exception:
            return {}
        return value if isinstance(value, dict) else {}


class AssetType(OpsBase, db.Model):
    __tablename__ = "asset_types"
    __table_args__ = (db.UniqueConstraint("organization_id", "name", name="uq_asset_types_org_name"),)
    name = db.Column(db.String(120), nullable=False)
    color = db.Column(db.String(20), nullable=False, default="#475569")
    icon = db.Column(db.String(60), nullable=False, default="box")


class AssetTypeLink(db.Model):
    __tablename__ = "asset_type_links"
    __table_args__ = (db.UniqueConstraint("asset_id", "asset_type_id", name="uq_asset_type_link"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    asset_id = db.Column(db.String(36), db.ForeignKey("assets.id"), nullable=False, index=True)
    asset_type_id = db.Column(db.String(36), db.ForeignKey("asset_types.id"), nullable=False, index=True)

    asset = db.relationship("Asset", back_populates="type_links")
    asset_type = db.relationship("AssetType")


class AssetStatusHistory(db.Model):
    __tablename__ = "asset_status_history"
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    asset_id = db.Column(db.String(36), db.ForeignKey("assets.id"), nullable=False, index=True)
    from_status = db.Column(db.String(30), nullable=True)
    to_status = db.Column(db.String(30), nullable=False)
    downtime_type = db.Column(db.String(30), nullable=True)  # planned / unplanned
    downtime_reason = db.Column(db.String(160), nullable=True)
    note = db.Column(db.String(500), nullable=True)
    started_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    ended_at = db.Column(db.DateTime, nullable=True)
    changed_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)

    asset = db.relationship("Asset")
    changed_by = db.relationship("User", foreign_keys=[changed_by_user_id])
