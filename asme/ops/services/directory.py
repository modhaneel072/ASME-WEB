"""Teams, locations, categories - the selectors everything else depends on."""

from __future__ import annotations

from datetime import datetime

from sqlalchemy import func

from asme.extensions import db
from asme.ops import audit
from asme.ops.models import Category, Location, OpsProject, Team, TeamMember
from asme.ops.services.common import apply_patch, ref, resolve_users, touch, user_ref
from asme.ops.tenancy import get_or_404, scoped
from asme.services.errors import Conflict, Validation

DEFAULT_CATEGORIES = [
    ("Mechanical", "#0878d1", "wrench"),
    ("Electrical", "#e58a00", "zap"),
    ("Software", "#7c5ce7", "code"),
    ("Embedded Systems", "#00a878", "cpu"),
    ("Fabrication", "#475569", "hammer"),
    ("Inspection", "#0891b2", "search-check"),
    ("Safety", "#d84a4a", "shield-alert"),
    ("Preventive", "#2563eb", "calendar-check"),
    ("Damage", "#b91c1c", "alert-triangle"),
    ("Project", "#ffcd00", "flag"),
    ("Event", "#db2777", "calendar"),
    ("Procurement", "#0f766e", "shopping-cart"),
    ("Documentation", "#6b7280", "file-text"),
    ("Standard Operating Procedure", "#1d4ed8", "clipboard-list"),
]

TEAM_FIELDS = ("name", "description", "parent_team_id", "project_id", "color", "escalation_note")
LOCATION_FIELDS = ("name", "description", "parent_location_id", "building", "room", "is_default")
CATEGORY_FIELDS = ("name", "color", "icon", "description")


# --------------------------------------------------------------------------- teams


def serialize_team(team: Team, memberships_by_user: dict | None = None) -> dict:
    members = []
    for tm in sorted(team.members, key=lambda m: (not m.is_lead, (m.user.name or "").lower())):
        payload = user_ref(tm.user, (memberships_by_user or {}).get(tm.user_id))
        payload["is_lead"] = bool(tm.is_lead)
        members.append(payload)
    return {
        "id": team.id,
        "name": team.name,
        "description": team.description,
        "color": team.color,
        "escalation_note": team.escalation_note,
        "parent_team_id": team.parent_team_id,
        "parent": ref(team.parent),
        "project_id": team.project_id,
        "project": ref(team.project),
        "members": members,
        "lead_ids": sorted(team.lead_user_ids),
        "member_count": len(team.members),
        "archived_at": team.archived_at.isoformat() if team.archived_at else None,
        "created_at": team.created_at.isoformat(),
        "updated_at": team.updated_at.isoformat(),
    }


def list_teams(ctx, include_archived=False):
    query = scoped(Team, ctx)
    if not include_archived:
        query = query.filter(Team.archived_at.is_(None))
    return query.order_by(Team.name.asc()).all()


def _validate_team_refs(ctx, parent_team_id, project_id, team_id=None):
    if parent_team_id:
        parent = get_or_404(Team, parent_team_id, ctx, "Parent team")
        if team_id and parent.id == team_id:
            raise Validation("A team cannot be its own parent.", field="parent_team_id")
    if project_id:
        get_or_404(OpsProject, project_id, ctx, "Project")


def _set_team_members(ctx, team: Team, members):
    users = resolve_users(ctx.organization.id, [m.user_id for m in members])
    wanted = {}
    for m in members:
        if m.user_id not in users:
            raise Validation(f"User {m.user_id} is not an active member of the chapter.", field="members")
        wanted[m.user_id] = bool(m.is_lead) or wanted.get(m.user_id, False)
    existing = {tm.user_id: tm for tm in team.members}
    for user_id, tm in list(existing.items()):
        if user_id not in wanted:
            db.session.delete(tm)
        else:
            tm.is_lead = wanted[user_id]
    for user_id, is_lead in wanted.items():
        if user_id not in existing:
            db.session.add(TeamMember(team_id=team.id, user_id=user_id, is_lead=is_lead))
    db.session.flush()


def create_team(ctx, data) -> Team:
    if scoped(Team, ctx).filter(func.lower(Team.name) == data.name.lower()).first():
        raise Conflict("A team with that name already exists.", code="name_taken")
    _validate_team_refs(ctx, data.parent_team_id, data.project_id)
    team = Team(
        organization_id=ctx.organization.id,
        name=data.name,
        description=data.description,
        parent_team_id=data.parent_team_id,
        project_id=data.project_id,
        color=data.color,
        escalation_note=data.escalation_note,
        created_by_user_id=ctx.user.id,
        updated_by_user_id=ctx.user.id,
    )
    db.session.add(team)
    db.session.flush()
    _set_team_members(ctx, team, data.members)
    audit.record_event("team.created", "team", team.id, organization_id=ctx.organization.id, actor=ctx.user, after=audit.snapshot(team, TEAM_FIELDS))
    db.session.commit()
    return team


def update_team(ctx, team: Team, data, fields_set: set[str]) -> Team:
    if "name" in fields_set and data.name and scoped(Team, ctx).filter(func.lower(Team.name) == data.name.lower(), Team.id != team.id).first():
        raise Conflict("A team with that name already exists.", code="name_taken")
    _validate_team_refs(ctx, getattr(data, "parent_team_id", None), getattr(data, "project_id", None), team.id)
    before = audit.snapshot(team, TEAM_FIELDS)
    apply_patch(team, data, fields_set, set(TEAM_FIELDS))
    touch(team, ctx.user)
    after = audit.snapshot(team, TEAM_FIELDS)
    audit.record_event("team.updated", "team", team.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"changed": audit.diff(before, after)})
    db.session.commit()
    return team


def set_team_members(ctx, team: Team, members) -> Team:
    before = {"members": sorted(team.member_user_ids), "leads": sorted(team.lead_user_ids)}
    _set_team_members(ctx, team, members)
    db.session.refresh(team)
    after = {"members": sorted(team.member_user_ids), "leads": sorted(team.lead_user_ids)}
    audit.record_event("team.members_set", "team", team.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after)
    db.session.commit()
    return team


def archive_team(ctx, team: Team) -> Team:
    team.archived_at = datetime.utcnow()
    touch(team, ctx.user)
    audit.record_event("team.archived", "team", team.id, organization_id=ctx.organization.id, actor=ctx.user)
    db.session.commit()
    return team


# --------------------------------------------------------------------------- locations


def serialize_location(location: Location, counts: dict | None = None) -> dict:
    return {
        "id": location.id,
        "name": location.name,
        "description": location.description,
        "parent_location_id": location.parent_location_id,
        "parent": ref(location.parent),
        "building": location.building,
        "room": location.room,
        "is_default": bool(location.is_default),
        "qr_code": location.qr_code,
        "asset_count": (counts or {}).get(location.id, 0),
        "archived_at": location.archived_at.isoformat() if location.archived_at else None,
        "created_at": location.created_at.isoformat(),
        "updated_at": location.updated_at.isoformat(),
    }


def list_locations(ctx, include_archived=False):
    query = scoped(Location, ctx)
    if not include_archived:
        query = query.filter(Location.archived_at.is_(None))
    return query.order_by(Location.is_default.desc(), Location.name.asc()).all()


def default_location(ctx_or_org_id) -> Location:
    org_id = ctx_or_org_id if isinstance(ctx_or_org_id, str) else ctx_or_org_id.organization.id
    row = Location.query.filter_by(organization_id=org_id, is_default=True).first()
    if row is None:
        row = Location.query.filter_by(organization_id=org_id).filter(func.lower(Location.name) == "general").first()
        if row is None:
            row = Location(organization_id=org_id, name="General", description="Default location for records without an explicit one.")
            db.session.add(row)
        row.is_default = True
        db.session.flush()
    return row


def _validate_location_parent(ctx, parent_id, location_id=None):
    if not parent_id:
        return
    parent = get_or_404(Location, parent_id, ctx, "Parent location")
    if location_id and parent.id == location_id:
        raise Validation("A location cannot be its own parent.", field="parent_location_id")


def create_location(ctx, data) -> Location:
    _validate_location_parent(ctx, data.parent_location_id)
    location = Location(
        organization_id=ctx.organization.id,
        name=data.name,
        description=data.description,
        parent_location_id=data.parent_location_id,
        building=data.building,
        room=data.room,
        created_by_user_id=ctx.user.id,
        updated_by_user_id=ctx.user.id,
    )
    db.session.add(location)
    db.session.flush()
    audit.record_event("location.created", "location", location.id, organization_id=ctx.organization.id, actor=ctx.user, after=audit.snapshot(location, LOCATION_FIELDS))
    db.session.commit()
    return location


def update_location(ctx, location: Location, data, fields_set: set[str]) -> Location:
    _validate_location_parent(ctx, getattr(data, "parent_location_id", None), location.id)
    before = audit.snapshot(location, LOCATION_FIELDS)
    apply_patch(location, data, fields_set, set(LOCATION_FIELDS) - {"is_default"})
    if "is_default" in fields_set and data.is_default:
        for other in scoped(Location, ctx).filter(Location.is_default.is_(True), Location.id != location.id).all():
            other.is_default = False
        location.is_default = True
    touch(location, ctx.user)
    after = audit.snapshot(location, LOCATION_FIELDS)
    audit.record_event("location.updated", "location", location.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"changed": audit.diff(before, after)})
    db.session.commit()
    return location


def archive_location(ctx, location: Location) -> Location:
    if location.is_default:
        raise Conflict("The default location cannot be archived.", code="default_location")
    location.archived_at = datetime.utcnow()
    touch(location, ctx.user)
    audit.record_event("location.archived", "location", location.id, organization_id=ctx.organization.id, actor=ctx.user)
    db.session.commit()
    return location


# --------------------------------------------------------------------------- categories


def serialize_category(category: Category, usage: dict | None = None) -> dict:
    return {
        "id": category.id,
        "name": category.name,
        "color": category.color,
        "icon": category.icon,
        "description": category.description,
        "work_order_count": (usage or {}).get(category.id, 0),
        "archived_at": category.archived_at.isoformat() if category.archived_at else None,
        "created_at": category.created_at.isoformat(),
        "updated_at": category.updated_at.isoformat(),
        "created_by": user_ref(category.creator) if hasattr(category, "creator") else None,
    }


def list_categories(ctx, include_archived=False):
    query = scoped(Category, ctx)
    if not include_archived:
        query = query.filter(Category.archived_at.is_(None))
    return query.order_by(Category.name.asc()).all()


def ensure_default_categories(org_id: str) -> int:
    existing = {c.name.lower() for c in Category.query.filter_by(organization_id=org_id).all()}
    created = 0
    for name, color, icon in DEFAULT_CATEGORIES:
        if name.lower() not in existing:
            db.session.add(Category(organization_id=org_id, name=name, color=color, icon=icon))
            created += 1
    db.session.flush()
    return created


def create_category(ctx, data) -> Category:
    if scoped(Category, ctx).filter(func.lower(Category.name) == data.name.lower()).first():
        raise Conflict("A category with that name already exists.", code="name_taken")
    category = Category(
        organization_id=ctx.organization.id,
        name=data.name,
        color=data.color,
        icon=data.icon,
        description=data.description,
        created_by_user_id=ctx.user.id,
        updated_by_user_id=ctx.user.id,
    )
    db.session.add(category)
    db.session.flush()
    audit.record_event("category.created", "category", category.id, organization_id=ctx.organization.id, actor=ctx.user, after=audit.snapshot(category, CATEGORY_FIELDS))
    db.session.commit()
    return category


def update_category(ctx, category: Category, data, fields_set: set[str]) -> Category:
    if "name" in fields_set and data.name and scoped(Category, ctx).filter(func.lower(Category.name) == data.name.lower(), Category.id != category.id).first():
        raise Conflict("A category with that name already exists.", code="name_taken")
    before = audit.snapshot(category, CATEGORY_FIELDS)
    apply_patch(category, data, fields_set, set(CATEGORY_FIELDS))
    touch(category, ctx.user)
    after = audit.snapshot(category, CATEGORY_FIELDS)
    audit.record_event("category.updated", "category", category.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"changed": audit.diff(before, after)})
    db.session.commit()
    return category


def archive_category(ctx, category: Category) -> Category:
    category.archived_at = datetime.utcnow()
    touch(category, ctx.user)
    audit.record_event("category.archived", "category", category.id, organization_id=ctx.organization.id, actor=ctx.user)
    db.session.commit()
    return category
