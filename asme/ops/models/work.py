"""Work orders and everything hanging off them."""

from __future__ import annotations

import json
from datetime import datetime

from asme.extensions import db
from asme.ops.models.base import OpsBase, new_uuid, utcnow

WO_STATUSES = ("DRAFT", "OPEN", "IN_PROGRESS", "ON_HOLD", "DONE", "CANCELED", "SKIPPED")
WO_OPEN_STATUSES = ("DRAFT", "OPEN", "IN_PROGRESS", "ON_HOLD")
WO_CLOSED_STATUSES = ("DONE", "CANCELED", "SKIPPED")
WO_PRIORITIES = ("NONE", "LOW", "MEDIUM", "HIGH", "CRITICAL")
WO_PRIORITY_RANK = {"NONE": 0, "LOW": 1, "MEDIUM": 2, "HIGH": 3, "CRITICAL": 4}
WO_WORK_TYPES = ("REACTIVE", "PREVENTIVE", "PROJECT", "EVENT", "INSPECTION", "SAFETY", "PROCUREMENT", "DOCUMENTATION")

# status -> allowed next statuses (spec §10.6)
WO_TRANSITIONS = {
    "DRAFT": {"OPEN", "CANCELED"},
    "OPEN": {"IN_PROGRESS", "DONE", "CANCELED", "ON_HOLD"},
    "IN_PROGRESS": {"ON_HOLD", "DONE", "CANCELED"},
    "ON_HOLD": {"IN_PROGRESS", "CANCELED"},
    "DONE": set(),
    "CANCELED": set(),
    "SKIPPED": set(),
}


class WorkOrderCounter(db.Model):
    """Per-organisation sequence for the human-facing work-order number."""

    __tablename__ = "work_order_counters"
    organization_id = db.Column(db.String(36), db.ForeignKey("organizations.id"), primary_key=True)
    next_number = db.Column(db.Integer, nullable=False, default=1)


class WorkOrder(OpsBase, db.Model):
    __tablename__ = "work_orders"
    __table_args__ = (
        db.UniqueConstraint("organization_id", "number", name="uq_work_orders_org_number"),
        db.Index("ix_work_orders_org_status", "organization_id", "status"),
        db.Index("ix_work_orders_due", "organization_id", "due_at"),
    )
    number = db.Column(db.Integer, nullable=False)
    title = db.Column(db.String(220), nullable=False)
    description = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), nullable=False, default="OPEN")
    priority = db.Column(db.String(20), nullable=False, default="NONE")
    priority_rank = db.Column(db.Integer, nullable=False, default=0)
    work_type = db.Column(db.String(30), nullable=False, default="PROJECT")
    project_id = db.Column(db.String(36), db.ForeignKey("ops_projects.id"), nullable=True, index=True)
    team_id = db.Column(db.String(36), db.ForeignKey("teams.id"), nullable=True, index=True)
    location_id = db.Column(db.String(36), db.ForeignKey("locations.id"), nullable=True, index=True)
    primary_asset_id = db.Column(db.String(36), db.ForeignKey("assets.id"), nullable=True, index=True)
    parent_work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=True, index=True)
    source_request_id = db.Column(db.String(36), nullable=True)
    maintenance_plan_id = db.Column(db.String(36), nullable=True)
    generation_key = db.Column(db.String(120), nullable=True, unique=True)  # idempotency for plans/recurrence
    vendor_id = db.Column(db.String(36), nullable=True)
    budget_code = db.Column(db.String(60), nullable=True)
    start_at = db.Column(db.DateTime, nullable=True)
    due_at = db.Column(db.DateTime, nullable=True)
    completed_at = db.Column(db.DateTime, nullable=True)
    canceled_at = db.Column(db.DateTime, nullable=True)
    estimated_minutes = db.Column(db.Integer, nullable=True)
    actual_minutes = db.Column(db.Integer, nullable=False, default=0)
    recurring_rule_json = db.Column(db.Text, nullable=True)
    completion_note = db.Column(db.Text, nullable=True)
    cancel_reason = db.Column(db.String(500), nullable=True)
    is_blocked = db.Column(db.Boolean, nullable=False, default=False)
    parent_completion_policy = db.Column(db.String(20), nullable=False, default="manual")  # manual / auto
    last_activity_at = db.Column(db.DateTime, default=utcnow, nullable=False)

    project = db.relationship("OpsProject", foreign_keys=[project_id])
    team = db.relationship("Team", foreign_keys=[team_id])
    location = db.relationship("Location", foreign_keys=[location_id])
    primary_asset = db.relationship("Asset", foreign_keys=[primary_asset_id])
    parent = db.relationship("WorkOrder", remote_side="WorkOrder.id", foreign_keys=[parent_work_order_id], backref="children")
    creator = db.relationship("User", foreign_keys="WorkOrder.created_by_user_id")
    assignees = db.relationship("WorkOrderAssignee", back_populates="work_order", cascade="all, delete-orphan")
    categories = db.relationship("WorkOrderCategory", back_populates="work_order", cascade="all, delete-orphan")
    asset_links = db.relationship("WorkOrderAsset", back_populates="work_order", cascade="all, delete-orphan")
    watchers = db.relationship("WorkOrderWatcher", back_populates="work_order", cascade="all, delete-orphan")
    status_history = db.relationship(
        "WorkOrderStatusHistory", back_populates="work_order", cascade="all, delete-orphan", order_by="WorkOrderStatusHistory.changed_at"
    )
    time_entries = db.relationship("TimeEntry", back_populates="work_order", cascade="all, delete-orphan")
    cost_entries = db.relationship("CostEntry", back_populates="work_order", cascade="all, delete-orphan")

    @property
    def is_open(self) -> bool:
        return self.status in WO_OPEN_STATUSES

    @property
    def is_overdue(self) -> bool:
        return bool(self.due_at) and self.is_open and self.due_at < datetime.utcnow()

    @property
    def assignee_user_ids(self) -> set[int]:
        return {a.user_id for a in self.assignees if a.user_id}

    @property
    def assignee_team_ids(self) -> set[str]:
        return {a.team_id for a in self.assignees if a.team_id}

    @property
    def category_ids(self) -> list[str]:
        return [c.category_id for c in self.categories]

    @property
    def recurring_rule(self) -> dict | None:
        if not self.recurring_rule_json:
            return None
        try:
            return json.loads(self.recurring_rule_json)
        except Exception:
            return None


class WorkOrderAssignee(db.Model):
    __tablename__ = "work_order_assignees"
    __table_args__ = (
        db.UniqueConstraint("work_order_id", "user_id", name="uq_wo_assignee_user"),
        db.UniqueConstraint("work_order_id", "team_id", name="uq_wo_assignee_team"),
    )
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    team_id = db.Column(db.String(36), db.ForeignKey("teams.id"), nullable=True, index=True)
    assigned_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    assigned_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)

    work_order = db.relationship("WorkOrder", back_populates="assignees")
    user = db.relationship("User", foreign_keys=[user_id])
    team = db.relationship("Team", foreign_keys=[team_id])


class WorkOrderCategory(db.Model):
    __tablename__ = "work_order_categories"
    __table_args__ = (db.UniqueConstraint("work_order_id", "category_id", name="uq_wo_category"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    category_id = db.Column(db.String(36), db.ForeignKey("categories.id"), nullable=False, index=True)

    work_order = db.relationship("WorkOrder", back_populates="categories")
    category = db.relationship("Category")


class WorkOrderAsset(db.Model):
    __tablename__ = "work_order_assets"
    __table_args__ = (db.UniqueConstraint("work_order_id", "asset_id", name="uq_wo_asset"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    asset_id = db.Column(db.String(36), db.ForeignKey("assets.id"), nullable=False, index=True)
    relationship_type = db.Column(db.String(20), nullable=False, default="related")  # primary / related

    work_order = db.relationship("WorkOrder", back_populates="asset_links")
    asset = db.relationship("Asset")


class WorkOrderWatcher(db.Model):
    __tablename__ = "work_order_watchers"
    __table_args__ = (db.UniqueConstraint("work_order_id", "user_id", name="uq_wo_watcher"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)

    work_order = db.relationship("WorkOrder", back_populates="watchers")
    user = db.relationship("User")


class WorkOrderStatusHistory(db.Model):
    __tablename__ = "work_order_status_history"
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    from_status = db.Column(db.String(20), nullable=True)
    to_status = db.Column(db.String(20), nullable=False)
    changed_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    note = db.Column(db.String(500), nullable=True)
    changed_at = db.Column(db.DateTime, default=utcnow, nullable=False)

    work_order = db.relationship("WorkOrder", back_populates="status_history")
    changed_by = db.relationship("User", foreign_keys=[changed_by_user_id])


class WorkOrderDependency(db.Model):
    __tablename__ = "work_order_dependencies"
    __table_args__ = (db.UniqueConstraint("blocking_work_order_id", "blocked_work_order_id", name="uq_wo_dependency"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    blocking_work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    blocked_work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    dependency_type = db.Column(db.String(20), nullable=False, default="blocks")

    blocking = db.relationship("WorkOrder", foreign_keys=[blocking_work_order_id])
    blocked = db.relationship("WorkOrder", foreign_keys=[blocked_work_order_id])


class TimeEntry(OpsBase, db.Model):
    __tablename__ = "time_entries"
    work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    started_at = db.Column(db.DateTime, nullable=True)
    ended_at = db.Column(db.DateTime, nullable=True)
    minutes = db.Column(db.Integer, nullable=False)
    note = db.Column(db.String(500), nullable=True)

    work_order = db.relationship("WorkOrder", back_populates="time_entries")
    user = db.relationship("User", foreign_keys=[user_id])


class CostEntry(OpsBase, db.Model):
    __tablename__ = "cost_entries"
    work_order_id = db.Column(db.String(36), db.ForeignKey("work_orders.id"), nullable=False, index=True)
    type = db.Column(db.String(20), nullable=False, default="other")  # part / labor / vendor / other
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    vendor_id = db.Column(db.String(36), nullable=True)
    description = db.Column(db.String(300), nullable=True)

    work_order = db.relationship("WorkOrder", back_populates="cost_entries")
