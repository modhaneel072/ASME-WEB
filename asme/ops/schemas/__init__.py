"""Typed request/response contracts (pydantic v2). The OpenAPI document at
``/api/v1/ops/openapi.json`` is generated from these; the frontend mirrors them in
``apps/ops-web/src/contracts``."""

from __future__ import annotations

from datetime import date, datetime, timezone
from decimal import Decimal
from typing import Annotated, Any, Literal, Optional

from pydantic import AfterValidator, BaseModel, ConfigDict, Field, field_validator

from asme.constants import EMAIL_RE
from asme.ops.models.assets import ASSET_CRITICALITIES, ASSET_STATUSES
from asme.ops.models.projects import MILESTONE_STATUSES, PROJECT_ROLES, PROJECT_STATUSES, RISK_LEVELS
from asme.ops.models.work import WO_PRIORITIES, WO_WORK_TYPES


def _to_naive_utc(value: datetime) -> datetime:
    """Clients send ISO-8601 with an offset (``...Z``); the database stores naive UTC."""
    if value.tzinfo is not None:
        return value.astimezone(timezone.utc).replace(tzinfo=None)
    return value


UtcDateTime = Annotated[datetime, AfterValidator(_to_naive_utc)]


class OpsModel(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class PatchModel(BaseModel):
    """PATCH bodies: unknown keys rejected, every field optional; ``model_fields_set``
    tells the service which keys were actually sent."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


def _non_empty(value: str, label: str = "value") -> str:
    if not value or not value.strip():
        raise ValueError(f"{label} is required")
    return value.strip()


# --------------------------------------------------------------------------- session / org


class LoginRequest(OpsModel):
    identifier: str = Field(min_length=1, max_length=160, description="Email or username")
    password: str = Field(min_length=1, max_length=200)


class OrgProfileUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=2, max_length=200)
    timezone: Optional[str] = Field(default=None, min_length=2, max_length=60)
    academic_year_start_month: Optional[int] = Field(default=None, ge=1, le=12)
    logo_url: Optional[str] = Field(default=None, max_length=500)


class SetupBannerUpdate(OpsModel):
    dismissed: bool


class UserInvite(OpsModel):
    name: str = Field(min_length=2, max_length=160)
    email: str = Field(min_length=5, max_length=160)
    role_key: str = Field(min_length=2, max_length=60)
    password: Optional[str] = Field(default=None, min_length=8, max_length=200)
    title: Optional[str] = Field(default=None, max_length=120)

    @field_validator("email")
    @classmethod
    def _email(cls, value: str) -> str:
        value = (value or "").strip().lower()
        if not EMAIL_RE.match(value):
            raise ValueError("Enter a valid email address")
        return value


class MemberUpdate(PatchModel):
    role_key: Optional[str] = Field(default=None, min_length=2, max_length=60)
    member_status: Optional[Literal["active", "invited", "suspended"]] = None
    title: Optional[str] = Field(default=None, max_length=120)


# --------------------------------------------------------------------------- teams / locations / categories


class TeamMemberInput(OpsModel):
    user_id: int
    is_lead: bool = False


class TeamCreate(OpsModel):
    name: str = Field(min_length=2, max_length=160)
    description: Optional[str] = Field(default=None, max_length=2000)
    parent_team_id: Optional[str] = None
    project_id: Optional[str] = None
    color: Optional[str] = Field(default=None, max_length=20)
    escalation_note: Optional[str] = Field(default=None, max_length=500)
    members: list[TeamMemberInput] = Field(default_factory=list)


class TeamUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=2, max_length=160)
    description: Optional[str] = Field(default=None, max_length=2000)
    parent_team_id: Optional[str] = None
    project_id: Optional[str] = None
    color: Optional[str] = Field(default=None, max_length=20)
    escalation_note: Optional[str] = Field(default=None, max_length=500)


class TeamMembersSet(OpsModel):
    members: list[TeamMemberInput]


class LocationCreate(OpsModel):
    name: str = Field(min_length=1, max_length=160)
    description: Optional[str] = Field(default=None, max_length=2000)
    parent_location_id: Optional[str] = None
    building: Optional[str] = Field(default=None, max_length=160)
    room: Optional[str] = Field(default=None, max_length=80)


class LocationUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=1, max_length=160)
    description: Optional[str] = Field(default=None, max_length=2000)
    parent_location_id: Optional[str] = None
    building: Optional[str] = Field(default=None, max_length=160)
    room: Optional[str] = Field(default=None, max_length=80)
    is_default: Optional[bool] = None


class CategoryCreate(OpsModel):
    name: str = Field(min_length=1, max_length=120)
    color: str = Field(default="#0878d1", pattern=r"^#[0-9a-fA-F]{6}$")
    icon: str = Field(default="tag", max_length=60)
    description: Optional[str] = Field(default=None, max_length=2000)


class CategoryUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=1, max_length=120)
    color: Optional[str] = Field(default=None, pattern=r"^#[0-9a-fA-F]{6}$")
    icon: Optional[str] = Field(default=None, max_length=60)
    description: Optional[str] = Field(default=None, max_length=2000)


# --------------------------------------------------------------------------- assets


class AssetCreate(OpsModel):
    name: str = Field(min_length=1, max_length=200)
    code: Optional[str] = Field(default=None, max_length=60)
    description: Optional[str] = Field(default=None, max_length=5000)
    parent_asset_id: Optional[str] = None
    project_id: Optional[str] = None
    location_id: Optional[str] = None
    responsible_team_id: Optional[str] = None
    owner_user_id: Optional[int] = None
    manufacturer: Optional[str] = Field(default=None, max_length=160)
    model: Optional[str] = Field(default=None, max_length=160)
    serial_number: Optional[str] = Field(default=None, max_length=160)
    purchase_date: Optional[date] = None
    purchase_cost: Optional[Decimal] = Field(default=None, ge=0)
    warranty_end: Optional[date] = None
    criticality: Literal[ASSET_CRITICALITIES] = "medium"  # type: ignore[valid-type]
    status: Literal[ASSET_STATUSES] = "ONLINE"  # type: ignore[valid-type]
    asset_type_ids: list[str] = Field(default_factory=list)
    custom_fields: dict[str, Any] = Field(default_factory=dict)


class AssetUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=1, max_length=200)
    code: Optional[str] = Field(default=None, max_length=60)
    description: Optional[str] = Field(default=None, max_length=5000)
    parent_asset_id: Optional[str] = None
    project_id: Optional[str] = None
    location_id: Optional[str] = None
    responsible_team_id: Optional[str] = None
    owner_user_id: Optional[int] = None
    manufacturer: Optional[str] = Field(default=None, max_length=160)
    model: Optional[str] = Field(default=None, max_length=160)
    serial_number: Optional[str] = Field(default=None, max_length=160)
    purchase_date: Optional[date] = None
    purchase_cost: Optional[Decimal] = Field(default=None, ge=0)
    warranty_end: Optional[date] = None
    criticality: Optional[Literal[ASSET_CRITICALITIES]] = None  # type: ignore[valid-type]
    asset_type_ids: Optional[list[str]] = None
    custom_fields: Optional[dict[str, Any]] = None


class AssetStatusChange(OpsModel):
    status: Literal[ASSET_STATUSES]  # type: ignore[valid-type]
    downtime_type: Optional[Literal["planned", "unplanned"]] = None
    downtime_reason: Optional[str] = Field(default=None, max_length=160)
    note: Optional[str] = Field(default=None, max_length=500)


# --------------------------------------------------------------------------- projects


class ProjectCreate(OpsModel):
    name: str = Field(min_length=2, max_length=200)
    code: Optional[str] = Field(default=None, min_length=2, max_length=30)
    description: Optional[str] = Field(default=None, max_length=10000)
    status: Literal[PROJECT_STATUSES] = "active"  # type: ignore[valid-type]
    risk_level: Literal[RISK_LEVELS] = "medium"  # type: ignore[valid-type]
    lead_user_id: Optional[int] = None
    faculty_advisor_user_id: Optional[int] = None
    competition: Optional[str] = Field(default=None, max_length=200)
    academic_year: Optional[str] = Field(default=None, max_length=9)
    start_date: Optional[date] = None
    target_date: Optional[date] = None
    budget_amount: Optional[Decimal] = Field(default=None, ge=0)
    budget_code: Optional[str] = Field(default=None, max_length=60)
    repository_url: Optional[str] = Field(default=None, max_length=500)
    cad_url: Optional[str] = Field(default=None, max_length=500)
    requirements_url: Optional[str] = Field(default=None, max_length=500)
    visibility: Literal["private", "members", "public"] = "members"
    public_project_id: Optional[int] = None


class ProjectUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=2, max_length=200)
    code: Optional[str] = Field(default=None, min_length=2, max_length=30)
    description: Optional[str] = Field(default=None, max_length=10000)
    status: Optional[Literal[PROJECT_STATUSES]] = None  # type: ignore[valid-type]
    risk_level: Optional[Literal[RISK_LEVELS]] = None  # type: ignore[valid-type]
    lead_user_id: Optional[int] = None
    faculty_advisor_user_id: Optional[int] = None
    competition: Optional[str] = Field(default=None, max_length=200)
    academic_year: Optional[str] = Field(default=None, max_length=9)
    start_date: Optional[date] = None
    target_date: Optional[date] = None
    budget_amount: Optional[Decimal] = Field(default=None, ge=0)
    budget_code: Optional[str] = Field(default=None, max_length=60)
    repository_url: Optional[str] = Field(default=None, max_length=500)
    cad_url: Optional[str] = Field(default=None, max_length=500)
    requirements_url: Optional[str] = Field(default=None, max_length=500)
    visibility: Optional[Literal["private", "members", "public"]] = None
    public_project_id: Optional[int] = None


class ProjectMemberInput(OpsModel):
    user_id: int
    project_role: Literal[PROJECT_ROLES] = "member"  # type: ignore[valid-type]
    team_id: Optional[str] = None


class ProjectMembersSet(OpsModel):
    members: list[ProjectMemberInput]


class MilestoneCreate(OpsModel):
    name: str = Field(min_length=1, max_length=200)
    description: Optional[str] = Field(default=None, max_length=5000)
    due_date: Optional[date] = None
    status: Literal[MILESTONE_STATUSES] = "planned"  # type: ignore[valid-type]
    owner_user_id: Optional[int] = None
    weight: int = Field(default=1, ge=1, le=100)


class MilestoneUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=1, max_length=200)
    description: Optional[str] = Field(default=None, max_length=5000)
    due_date: Optional[date] = None
    status: Optional[Literal[MILESTONE_STATUSES]] = None  # type: ignore[valid-type]
    owner_user_id: Optional[int] = None
    weight: Optional[int] = Field(default=None, ge=1, le=100)
    order: Optional[int] = Field(default=None, ge=0)


# --------------------------------------------------------------------------- work orders


class RecurrenceRule(OpsModel):
    frequency: Literal["daily", "weekly", "monthly"]
    interval: int = Field(default=1, ge=1, le=52)
    mode: Literal["fixed", "floating"] = "fixed"


class WorkOrderCreate(OpsModel):
    title: str = Field(min_length=1, max_length=220)
    description: Optional[str] = Field(default=None, max_length=20000)
    project_id: Optional[str] = None
    location_id: Optional[str] = None
    primary_asset_id: Optional[str] = None
    asset_ids: list[str] = Field(default_factory=list)
    assignee_user_ids: list[int] = Field(default_factory=list)
    assignee_team_ids: list[str] = Field(default_factory=list)
    team_id: Optional[str] = None
    estimated_minutes: Optional[int] = Field(default=None, ge=0, le=100000)
    due_at: Optional[UtcDateTime] = None
    start_at: Optional[UtcDateTime] = None
    recurrence: Optional[RecurrenceRule] = None
    work_type: Literal[WO_WORK_TYPES] = "PROJECT"  # type: ignore[valid-type]
    priority: Literal[WO_PRIORITIES] = "NONE"  # type: ignore[valid-type]
    category_ids: list[str] = Field(default_factory=list)
    budget_code: Optional[str] = Field(default=None, max_length=60)
    watcher_user_ids: list[int] = Field(default_factory=list)
    parent_work_order_id: Optional[str] = None
    status: Literal["DRAFT", "OPEN"] = "OPEN"
    parent_completion_policy: Literal["manual", "auto"] = "manual"

    @field_validator("title")
    @classmethod
    def _title(cls, value):
        return _non_empty(value, "Title")


class WorkOrderUpdate(PatchModel):
    title: Optional[str] = Field(default=None, min_length=1, max_length=220)
    description: Optional[str] = Field(default=None, max_length=20000)
    project_id: Optional[str] = None
    location_id: Optional[str] = None
    primary_asset_id: Optional[str] = None
    asset_ids: Optional[list[str]] = None
    team_id: Optional[str] = None
    estimated_minutes: Optional[int] = Field(default=None, ge=0, le=100000)
    due_at: Optional[UtcDateTime] = None
    start_at: Optional[UtcDateTime] = None
    recurrence: Optional[RecurrenceRule] = None
    work_type: Optional[Literal[WO_WORK_TYPES]] = None  # type: ignore[valid-type]
    priority: Optional[Literal[WO_PRIORITIES]] = None  # type: ignore[valid-type]
    category_ids: Optional[list[str]] = None
    budget_code: Optional[str] = Field(default=None, max_length=60)
    is_blocked: Optional[bool] = None
    parent_completion_policy: Optional[Literal["manual", "auto"]] = None


class AssigneesSet(OpsModel):
    user_ids: list[int] = Field(default_factory=list)
    team_ids: list[str] = Field(default_factory=list)


class WatchersSet(OpsModel):
    user_ids: list[int] = Field(default_factory=list)


class StatusNote(OpsModel):
    note: Optional[str] = Field(default=None, max_length=500)


class CostInput(OpsModel):
    type: Literal["part", "labor", "vendor", "other"] = "other"
    amount: Decimal = Field(ge=0)
    description: Optional[str] = Field(default=None, max_length=300)


class WorkOrderComplete(OpsModel):
    completion_note: Optional[str] = Field(default=None, max_length=5000)
    time_minutes: Optional[int] = Field(default=None, ge=0, le=100000)
    costs: list[CostInput] = Field(default_factory=list)
    asset_status: Optional[Literal[ASSET_STATUSES]] = None  # type: ignore[valid-type]
    follow_up_title: Optional[str] = Field(default=None, max_length=220)


class CommentCreate(OpsModel):
    body: str = Field(min_length=1, max_length=10000)
    parent_comment_id: Optional[str] = None


class CommentUpdate(OpsModel):
    body: str = Field(min_length=1, max_length=10000)


class TimeEntryCreate(OpsModel):
    minutes: int = Field(ge=1, le=100000)
    started_at: Optional[UtcDateTime] = None
    ended_at: Optional[UtcDateTime] = None
    note: Optional[str] = Field(default=None, max_length=500)
    user_id: Optional[int] = None  # leads may log on behalf of a member


class SubWorkOrderCreate(OpsModel):
    title: str = Field(min_length=1, max_length=220)
    description: Optional[str] = Field(default=None, max_length=20000)
    assignee_user_ids: list[int] = Field(default_factory=list)
    due_at: Optional[UtcDateTime] = None
    priority: Literal[WO_PRIORITIES] = "NONE"  # type: ignore[valid-type]
    estimated_minutes: Optional[int] = Field(default=None, ge=0, le=100000)


# --------------------------------------------------------------------------- saved filters


class SavedFilterCreate(OpsModel):
    entity_type: Literal["work_order", "project", "asset"]
    name: str = Field(min_length=1, max_length=120)
    visibility: Literal["private", "team", "chapter"] = "private"
    team_id: Optional[str] = None
    filters: dict[str, Any] = Field(default_factory=dict)
    sort: Optional[str] = Field(default=None, max_length=60)
    view_type: Literal["panel", "table"] = "panel"
    is_default: bool = False


class SavedFilterUpdate(PatchModel):
    name: Optional[str] = Field(default=None, min_length=1, max_length=120)
    visibility: Optional[Literal["private", "team", "chapter"]] = None
    team_id: Optional[str] = None
    filters: Optional[dict[str, Any]] = None
    sort: Optional[str] = Field(default=None, max_length=60)
    view_type: Optional[Literal["panel", "table"]] = None
    is_default: Optional[bool] = None


# --------------------------------------------------------------------------- response shapes (for OpenAPI + contract tests)


class UserRef(BaseModel):
    id: int
    name: str
    email: str
    initials: str
    role_key: Optional[str] = None
    role_name: Optional[str] = None


class Ref(BaseModel):
    id: str
    name: str


class StatusHistoryOut(BaseModel):
    id: str
    from_status: Optional[str]
    to_status: str
    changed_by: Optional[UserRef]
    note: Optional[str]
    changed_at: datetime


class CommentOut(BaseModel):
    id: str
    body: str
    author: UserRef
    parent_comment_id: Optional[str]
    created_at: datetime
    edited_at: Optional[datetime]


class AttachmentOut(BaseModel):
    id: str
    original_name: str
    content_type: str
    size_bytes: int
    kind: str
    uploaded_by: Optional[UserRef]
    created_at: datetime
    download_url: str
    scan_status: str


class TimeEntryOut(BaseModel):
    id: str
    user: UserRef
    minutes: int
    started_at: Optional[datetime]
    ended_at: Optional[datetime]
    note: Optional[str]
    created_at: datetime


class CostEntryOut(BaseModel):
    id: str
    type: str
    amount: float
    description: Optional[str]
    created_at: datetime


class WorkOrderListItem(BaseModel):
    id: str
    number: int
    title: str
    status: str
    priority: str
    work_type: str
    is_blocked: bool
    is_overdue: bool
    due_at: Optional[datetime]
    start_at: Optional[datetime]
    project: Optional[Ref]
    team: Optional[Ref]
    location: Optional[Ref]
    primary_asset: Optional[Ref]
    assignees: list[UserRef]
    assignee_teams: list[Ref]
    categories: list[dict]
    parent_work_order_id: Optional[str]
    child_count: int
    comment_count: int
    updated_at: datetime
    last_activity_at: datetime


class WorkOrderOut(WorkOrderListItem):
    description: Optional[str]
    estimated_minutes: Optional[int]
    actual_minutes: int
    completed_at: Optional[datetime]
    canceled_at: Optional[datetime]
    completion_note: Optional[str]
    cancel_reason: Optional[str]
    budget_code: Optional[str]
    recurrence: Optional[dict]
    parent_completion_policy: str
    creator: Optional[UserRef]
    watchers: list[UserRef]
    related_assets: list[Ref]
    children: list[WorkOrderListItem]
    status_history: list[StatusHistoryOut]
    time_entries: list[TimeEntryOut]
    cost_entries: list[CostEntryOut]
    total_cost: float
    allowed_transitions: list[str]
    permissions: dict[str, bool]
    created_at: datetime


class ProjectListItem(BaseModel):
    id: str
    name: str
    code: str
    status: str
    risk_level: str
    lead: Optional[UserRef]
    completion_percent: int
    open_work_orders: int
    overdue_work_orders: int
    next_milestone: Optional[dict]
    target_date: Optional[date]
    academic_year: Optional[str]
    competition: Optional[str]
    team_count: int
    updated_at: datetime


class SessionOut(BaseModel):
    user: UserRef
    organization: dict
    role: dict
    permissions: list[str]
    setup_banner_dismissed: bool
    unread_notifications: int
