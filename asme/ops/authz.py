"""Permission-based authorisation with scopes.

``can(ctx, "work_order.edit", record=wo)`` answers whether the current member may do
something *to that record*. The role supplies ``(permission, scope)`` pairs; the scope
is resolved against the record's organisation, project, team and assignees on the
server. See docs/permissions-matrix.md - this module is its source of truth.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from functools import wraps

from flask import g, jsonify

from asme.extensions import db
from asme.ops.models import Membership, OpsProject, OpsProjectMember, Permission, Role, RolePermission, Team, TeamMember

SCOPES = ("chapter", "project", "team", "assigned", "own")
SCOPE_RANK = {"chapter": 4, "project": 3, "team": 2, "assigned": 1, "own": 0}

PERMISSIONS: dict[str, str] = {
    "org.read": "See the organisation and its members",
    "org.manage": "Edit organisation profile and security settings",
    "setup.manage": "Run Setup Center and dismiss the setup banner for others",
    "user.manage": "Invite, deactivate and re-role members",
    "role.manage": "Create and edit custom roles",
    "team.read": "See teams and their members",
    "team.manage": "Create/edit teams and membership",
    "location.manage": "Create/edit locations",
    "category.manage": "Create/edit categories",
    "asset.read": "See assets",
    "asset.manage": "Create/edit assets and change their status",
    "project.read": "See projects",
    "project.create": "Create projects",
    "project.manage": "Edit projects, milestones and project membership",
    "work_order.read_all": "See every work order in scope",
    "work_order.read_assigned": "See work orders assigned to you",
    "work_order.create": "Create work orders",
    "work_order.edit": "Edit work-order fields",
    "work_order.assign": "Assign users and teams",
    "work_order.start": "Start, hold and resume work",
    "work_order.complete": "Complete work",
    "work_order.cancel": "Cancel work",
    "work_order.set_critical": "Use the CRITICAL priority",
    "comment.create": "Comment on records",
    "file.upload": "Upload files to records",
    "time_entry.create": "Log time",
    "cost_entry.create": "Log costs",
    "request.submit": "Submit requests",
    "request.review": "Review, approve, decline and convert requests",
    "inventory.read": "See parts and quantities",
    "inventory.manage": "Post inventory transactions, counts and transfers",
    "purchase.review": "Approve purchase requests",
    "procedure.publish": "Publish procedure versions",
    "safety.manage": "Manage safety inspections and corrective actions",
    "report.view": "View reports and dashboards",
    "report.export": "Export data",
    "dashboard.manage": "Create and share dashboards",
    "automation.manage": "Create and edit automations",
    "saved_filter.share": "Share saved filters with a team or the chapter",
    "audit.read": "Read the audit history",
}

_ALL = [(key, "chapter") for key in PERMISSIONS]

# role key -> (display name, rank, [(permission, scope)])
SYSTEM_ROLES: dict[str, tuple[str, int, list[tuple[str, str]]]] = {
    "chapter_admin": ("Chapter Administrator", 100, _ALL),
    "executive_officer": (
        "Executive Officer",
        90,
        [(k, "chapter") for k in PERMISSIONS if k not in {"org.manage", "role.manage", "inventory.manage", "safety.manage"}],
    ),
    "project_lead": (
        "Project Lead",
        80,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("team.manage", "project"), ("asset.read", "chapter"),
            ("asset.manage", "project"), ("project.read", "chapter"), ("project.manage", "project"),
            ("work_order.read_all", "chapter"), ("work_order.read_assigned", "chapter"), ("work_order.create", "project"),
            ("work_order.edit", "project"), ("work_order.assign", "project"), ("work_order.start", "project"),
            ("work_order.complete", "project"), ("work_order.cancel", "project"), ("work_order.set_critical", "project"),
            ("comment.create", "chapter"), ("file.upload", "chapter"), ("time_entry.create", "chapter"),
            ("cost_entry.create", "project"), ("request.submit", "chapter"), ("request.review", "project"),
            ("inventory.read", "chapter"), ("report.view", "chapter"), ("report.export", "project"),
            ("dashboard.manage", "chapter"), ("saved_filter.share", "chapter"),
        ],
    ),
    "team_lead": (
        "Team Lead",
        70,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("team.manage", "team"), ("asset.read", "chapter"),
            ("project.read", "chapter"), ("work_order.read_all", "chapter"), ("work_order.read_assigned", "chapter"),
            ("work_order.create", "team"), ("work_order.edit", "team"), ("work_order.assign", "team"),
            ("work_order.start", "team"), ("work_order.complete", "team"), ("work_order.cancel", "team"),
            ("comment.create", "chapter"), ("file.upload", "chapter"), ("time_entry.create", "chapter"),
            ("cost_entry.create", "team"), ("request.submit", "chapter"), ("request.review", "team"),
            ("inventory.read", "chapter"), ("report.view", "chapter"), ("dashboard.manage", "chapter"),
            ("saved_filter.share", "chapter"),
        ],
    ),
    "full_member": (
        "Full Member",
        60,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("asset.read", "chapter"), ("project.read", "chapter"),
            ("work_order.read_all", "chapter"), ("work_order.read_assigned", "chapter"), ("work_order.create", "own"),
            ("work_order.edit", "own"), ("work_order.start", "assigned"), ("work_order.complete", "assigned"),
            ("work_order.cancel", "own"), ("comment.create", "chapter"), ("file.upload", "chapter"),
            ("time_entry.create", "assigned"), ("request.submit", "chapter"), ("inventory.read", "chapter"),
            ("report.view", "chapter"),
        ],
    ),
    "shop_operator": (
        "Shop Operator",
        50,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("asset.read", "chapter"), ("project.read", "chapter"),
            ("work_order.read_assigned", "chapter"), ("work_order.start", "assigned"), ("work_order.complete", "assigned"),
            ("comment.create", "assigned"), ("file.upload", "assigned"), ("time_entry.create", "assigned"),
            ("request.submit", "chapter"), ("inventory.read", "chapter"),
        ],
    ),
    "requester": (
        "Requester",
        40,
        [("org.read", "chapter"), ("team.read", "chapter"), ("project.read", "chapter"), ("request.submit", "chapter")],
    ),
    "inventory_manager": (
        "Inventory Manager",
        65,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("location.manage", "chapter"), ("asset.read", "chapter"),
            ("asset.manage", "chapter"), ("project.read", "chapter"), ("work_order.read_all", "chapter"),
            ("work_order.read_assigned", "chapter"), ("work_order.create", "chapter"), ("work_order.edit", "chapter"),
            ("work_order.assign", "chapter"), ("work_order.start", "chapter"), ("work_order.complete", "chapter"),
            ("work_order.cancel", "chapter"), ("comment.create", "chapter"), ("file.upload", "chapter"),
            ("time_entry.create", "chapter"), ("cost_entry.create", "chapter"), ("request.submit", "chapter"),
            ("request.review", "chapter"), ("inventory.read", "chapter"), ("inventory.manage", "chapter"),
            ("report.view", "chapter"), ("report.export", "chapter"), ("saved_filter.share", "chapter"),
        ],
    ),
    "safety_officer": (
        "Safety Officer",
        65,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("asset.read", "chapter"), ("project.read", "chapter"),
            ("work_order.read_all", "chapter"), ("work_order.read_assigned", "chapter"), ("work_order.create", "chapter"),
            ("work_order.edit", "chapter"), ("work_order.assign", "chapter"), ("work_order.start", "chapter"),
            ("work_order.complete", "chapter"), ("work_order.cancel", "chapter"), ("work_order.set_critical", "chapter"),
            ("comment.create", "chapter"), ("file.upload", "chapter"), ("time_entry.create", "chapter"),
            ("request.submit", "chapter"), ("request.review", "chapter"), ("inventory.read", "chapter"),
            ("procedure.publish", "chapter"), ("safety.manage", "chapter"), ("report.view", "chapter"),
            ("report.export", "chapter"), ("saved_filter.share", "chapter"), ("audit.read", "chapter"),
        ],
    ),
    "treasurer": (
        "Treasurer",
        65,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("asset.read", "chapter"), ("project.read", "chapter"),
            ("work_order.read_all", "chapter"), ("work_order.read_assigned", "chapter"), ("comment.create", "chapter"),
            ("file.upload", "chapter"), ("cost_entry.create", "chapter"), ("request.submit", "chapter"),
            ("inventory.read", "chapter"), ("purchase.review", "chapter"), ("report.view", "chapter"),
            ("report.export", "chapter"), ("audit.read", "chapter"),
        ],
    ),
    "faculty_advisor": (
        "Faculty Advisor",
        65,
        [
            ("org.read", "chapter"), ("team.read", "chapter"), ("asset.read", "chapter"), ("project.read", "chapter"),
            ("work_order.read_all", "chapter"), ("work_order.read_assigned", "chapter"), ("comment.create", "chapter"),
            ("request.submit", "chapter"), ("inventory.read", "chapter"), ("purchase.review", "chapter"),
            ("report.view", "chapter"), ("report.export", "chapter"), ("audit.read", "chapter"),
        ],
    ),
    "sponsor_guest": ("Sponsor / Guest", 10, [("org.read", "chapter")]),
}

LEGACY_ROLE_MAP = {"admin": "chapter_admin", "team_leader": "team_lead", "member": "full_member"}
REVERSE_LEGACY_ROLE_MAP = {
    "chapter_admin": "admin",
    "executive_officer": "admin",
    "project_lead": "team_leader",
    "team_lead": "team_leader",
}


def legacy_role_for(system_key: str | None) -> str:
    return REVERSE_LEGACY_ROLE_MAP.get(system_key or "", "member")


# --------------------------------------------------------------------------- context


@dataclass
class AuthzContext:
    user: object
    organization: object
    membership: Membership
    role: Role
    grants: dict[str, str] = field(default_factory=dict)  # permission -> widest scope
    _led_team_ids: set[str] | None = None
    _member_team_ids: set[str] | None = None
    _project_ids: set[str] | None = None
    _led_project_ids: set[str] | None = None

    @property
    def is_admin(self) -> bool:
        return self.role.system_key == "chapter_admin"

    @property
    def permission_keys(self) -> list[str]:
        return sorted(self.grants)

    # -- lazy relationship lookups ---------------------------------------------

    def led_team_ids(self) -> set[str]:
        if self._led_team_ids is None:
            rows = (
                db.session.query(TeamMember.team_id)
                .join(Team, Team.id == TeamMember.team_id)
                .filter(TeamMember.user_id == self.user.id, TeamMember.is_lead.is_(True), Team.organization_id == self.organization.id)
                .all()
            )
            self._led_team_ids = {r[0] for r in rows}
        return self._led_team_ids

    def member_team_ids(self) -> set[str]:
        if self._member_team_ids is None:
            rows = (
                db.session.query(TeamMember.team_id)
                .join(Team, Team.id == TeamMember.team_id)
                .filter(TeamMember.user_id == self.user.id, Team.organization_id == self.organization.id)
                .all()
            )
            self._member_team_ids = {r[0] for r in rows}
        return self._member_team_ids

    def project_ids(self) -> set[str]:
        """Projects the user leads, advises, or belongs to (directly or via a team)."""
        if self._project_ids is None:
            ids = set()
            for (pid,) in db.session.query(OpsProject.id).filter(
                OpsProject.organization_id == self.organization.id,
                (OpsProject.lead_user_id == self.user.id) | (OpsProject.faculty_advisor_user_id == self.user.id),
            ):
                ids.add(pid)
            for (pid,) in (
                db.session.query(OpsProjectMember.project_id)
                .join(OpsProject, OpsProject.id == OpsProjectMember.project_id)
                .filter(OpsProjectMember.user_id == self.user.id, OpsProject.organization_id == self.organization.id)
            ):
                ids.add(pid)
            team_ids = self.member_team_ids()
            if team_ids:
                for (pid,) in db.session.query(Team.project_id).filter(Team.id.in_(team_ids), Team.project_id.isnot(None)):
                    ids.add(pid)
            self._project_ids = ids
        return self._project_ids

    def led_project_ids(self) -> set[str]:
        if self._led_project_ids is None:
            ids = {
                pid
                for (pid,) in db.session.query(OpsProject.id).filter(
                    OpsProject.organization_id == self.organization.id, OpsProject.lead_user_id == self.user.id
                )
            }
            for (pid,) in db.session.query(OpsProjectMember.project_id).filter(
                OpsProjectMember.user_id == self.user.id, OpsProjectMember.project_role == "lead"
            ):
                ids.add(pid)
            self._led_project_ids = ids
        return self._led_project_ids


def build_context(user, organization, membership) -> AuthzContext:
    role = membership.role
    grants: dict[str, str] = {}
    for rp in role.permissions:
        key = rp.permission.key
        scope = rp.scope_type if rp.scope_type in SCOPES else "own"
        if key not in grants or SCOPE_RANK[scope] > SCOPE_RANK[grants[key]]:
            grants[key] = scope
    return AuthzContext(user=user, organization=organization, membership=membership, role=role, grants=grants)


# --------------------------------------------------------------------------- record shape


@dataclass
class RecordScope:
    """The associations a scope check needs. Services build it from a record."""

    organization_id: str | None = None
    project_id: str | None = None
    team_ids: set[str] = field(default_factory=set)
    assignee_user_ids: set[int] = field(default_factory=set)
    assignee_team_ids: set[str] = field(default_factory=set)
    owner_user_id: int | None = None


def scope_of(record) -> RecordScope:
    """Best-effort extraction from any ops model."""
    rs = RecordScope(organization_id=getattr(record, "organization_id", None))
    rs.project_id = getattr(record, "project_id", None)
    if hasattr(record, "team_id") and getattr(record, "team_id", None):
        rs.team_ids.add(record.team_id)
    if hasattr(record, "assignee_user_ids"):
        rs.assignee_user_ids = set(record.assignee_user_ids)
    if hasattr(record, "assignee_team_ids"):
        rs.assignee_team_ids = set(record.assignee_team_ids)
        rs.team_ids |= rs.assignee_team_ids
    if hasattr(record, "lead_user_id") and getattr(record, "lead_user_id", None):
        rs.assignee_user_ids.add(record.lead_user_id)
    if record.__class__.__name__ == "OpsProject":
        rs.project_id = record.id
    if record.__class__.__name__ == "Team":
        rs.team_ids.add(record.id)
    rs.owner_user_id = getattr(record, "created_by_user_id", None)
    return rs


# --------------------------------------------------------------------------- checks


def can(ctx: AuthzContext, permission: str, record=None, scope: RecordScope | None = None) -> bool:
    if permission not in PERMISSIONS:
        raise KeyError(f"unknown permission {permission}")
    granted_scope = ctx.grants.get(permission)
    if granted_scope is None:
        return False
    if scope is None and record is not None:
        scope = scope_of(record)
    if scope is not None and scope.organization_id and scope.organization_id != ctx.organization.id:
        return False  # never cross tenants, whatever the role
    if granted_scope == "chapter":
        return True
    if scope is None:
        # No record yet (e.g. create). Project/team scopes need a target; own/assigned are fine.
        return granted_scope in {"own", "assigned", "project", "team"}
    if granted_scope == "project":
        return bool(scope.project_id) and scope.project_id in ctx.project_ids()
    if granted_scope == "team":
        led = ctx.led_team_ids()
        if scope.team_ids & led:
            return True
        # a team lead may act on project work belonging to a project one of their teams serves
        if scope.project_id and scope.project_id in ctx.project_ids() and led:
            return True
        return False
    if granted_scope == "assigned":
        if ctx.user.id in scope.assignee_user_ids:
            return True
        return bool(scope.assignee_team_ids & ctx.member_team_ids())
    if granted_scope == "own":
        return scope.owner_user_id == ctx.user.id or ctx.user.id in scope.assignee_user_ids
    return False


def can_create_for(ctx: AuthzContext, permission: str, *, project_id=None, team_id=None) -> bool:
    """Create-time check: does the role's scope cover the intended project/team?"""
    granted_scope = ctx.grants.get(permission)
    if granted_scope is None:
        return False
    if granted_scope == "chapter":
        return True
    if granted_scope == "project":
        return bool(project_id) and project_id in ctx.project_ids()
    if granted_scope == "team":
        if team_id:
            return team_id in ctx.led_team_ids()
        return bool(project_id) and project_id in ctx.project_ids() and bool(ctx.led_team_ids())
    return True  # own / assigned: the creator owns what they create


class Forbidden(Exception):
    def __init__(self, permission: str, message: str | None = None):
        super().__init__(message or f"Missing permission {permission}")
        self.permission = permission
        self.message = message or f"You do not have permission to do that ({permission})."


def require(ctx: AuthzContext, permission: str, record=None, scope: RecordScope | None = None) -> None:
    if not can(ctx, permission, record=record, scope=scope):
        raise Forbidden(permission)


def require_create(ctx: AuthzContext, permission: str, *, project_id=None, team_id=None) -> None:
    if not can_create_for(ctx, permission, project_id=project_id, team_id=team_id):
        raise Forbidden(permission)


def visible_work_scope(ctx: AuthzContext) -> str:
    """'all' if the user may read every work order in the chapter, else 'assigned'."""
    if ctx.grants.get("work_order.read_all") == "chapter":
        return "all"
    return "assigned"


# --------------------------------------------------------------------------- seeding


def ensure_permissions() -> dict[str, Permission]:
    rows = {p.key: p for p in Permission.query.all()}
    for key, description in PERMISSIONS.items():
        if key not in rows:
            row = Permission(key=key, description=description)
            db.session.add(row)
            rows[key] = row
        elif rows[key].description != description:
            rows[key].description = description
    db.session.flush()
    return rows


def ensure_system_roles(organization) -> dict[str, Role]:
    permissions = ensure_permissions()
    roles: dict[str, Role] = {}
    for key, (name, rank, grants) in SYSTEM_ROLES.items():
        role = Role.query.filter_by(organization_id=organization.id, system_key=key).first()
        if role is None:
            role = Role(organization_id=organization.id, name=name, system_key=key, is_custom=False, rank=rank)
            db.session.add(role)
            db.session.flush()
        role.name = name
        role.rank = rank
        existing = {(rp.permission.key): rp for rp in role.permissions}
        wanted = dict(grants)
        for perm_key, scope in wanted.items():
            if perm_key in existing:
                existing[perm_key].scope_type = scope
            else:
                db.session.add(RolePermission(role_id=role.id, permission_id=permissions[perm_key].id, scope_type=scope))
        for perm_key, rp in existing.items():
            if perm_key not in wanted:
                db.session.delete(rp)
        roles[key] = role
    db.session.flush()
    return roles


def forbidden_response(exc: Forbidden):
    return jsonify({"ok": False, "code": "forbidden", "error": exc.message, "permission": exc.permission}), 403


def current_ctx() -> AuthzContext | None:
    return getattr(g, "ops_ctx", None)


def permission_required(permission: str):
    """Route decorator for chapter-scoped permissions (no record involved)."""

    def decorator(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            ctx = current_ctx()
            if ctx is None:
                return jsonify({"ok": False, "code": "login_required", "error": "Login required."}), 401
            if not can(ctx, permission):
                return forbidden_response(Forbidden(permission))
            return fn(*args, **kwargs)

        return wrapped

    return decorator
