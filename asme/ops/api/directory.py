"""Teams, locations, categories, lookups for selectors."""

from __future__ import annotations

from sqlalchemy import func

from asme.extensions import db
from asme.ops import authz
from asme.ops.api import body, bp, contract, ctx, ok
from asme.ops.models import Asset, Category, Location, Membership, OpsProject, Team, WorkOrderCategory
from asme.ops.schemas import CategoryCreate, CategoryUpdate, LocationCreate, LocationUpdate, TeamCreate, TeamMembersSet, TeamUpdate
from asme.ops.services import directory, organizations
from asme.ops.services.common import user_ref
from asme.ops.tenancy import get_or_404


def _memberships_by_user(c):
    return {m.user_id: m for m in Membership.query.filter_by(organization_id=c.organization.id).all()}


# --------------------------------------------------------------------------- teams


@bp.get("/teams")
@authz.permission_required("team.read")
def teams_list():
    from flask import request

    include_archived = request.args.get("include_archived") in {"1", "true"}
    memberships = _memberships_by_user(ctx())
    return ok([directory.serialize_team(t, memberships) for t in directory.list_teams(ctx(), include_archived)])


@bp.post("/teams")
@contract(TeamCreate, summary="Create a team", tags=("teams",))
def teams_create():
    data = body(TeamCreate)
    authz.require_create(ctx(), "team.manage", project_id=data.project_id)
    team = directory.create_team(ctx(), data)
    return ok(directory.serialize_team(team, _memberships_by_user(ctx())), status=201)


@bp.get("/teams/<team_id>")
@authz.permission_required("team.read")
def teams_get(team_id):
    team = get_or_404(Team, team_id, ctx(), "Team")
    return ok(directory.serialize_team(team, _memberships_by_user(ctx())))


@bp.patch("/teams/<team_id>")
@contract(TeamUpdate, summary="Update a team", tags=("teams",))
def teams_update(team_id):
    team = get_or_404(Team, team_id, ctx(), "Team")
    authz.require(ctx(), "team.manage", record=team)
    data = body(TeamUpdate)
    team = directory.update_team(ctx(), team, data, data.model_fields_set)
    return ok(directory.serialize_team(team, _memberships_by_user(ctx())))


@bp.put("/teams/<team_id>/members")
@contract(TeamMembersSet, summary="Replace a team's members and leads", tags=("teams",))
def teams_members(team_id):
    team = get_or_404(Team, team_id, ctx(), "Team")
    authz.require(ctx(), "team.manage", record=team)
    data = body(TeamMembersSet)
    team = directory.set_team_members(ctx(), team, data.members)
    return ok(directory.serialize_team(team, _memberships_by_user(ctx())))


@bp.delete("/teams/<team_id>")
def teams_archive(team_id):
    team = get_or_404(Team, team_id, ctx(), "Team")
    authz.require(ctx(), "team.manage", record=team)
    directory.archive_team(ctx(), team)
    return ok({"archived": True})


# --------------------------------------------------------------------------- locations


@bp.get("/locations")
def locations_list():
    from flask import request

    include_archived = request.args.get("include_archived") in {"1", "true"}
    counts = {lid: int(n) for lid, n in db.session.query(Asset.location_id, func.count(Asset.id)).filter(Asset.organization_id == ctx().organization.id, Asset.archived_at.is_(None)).group_by(Asset.location_id).all()}
    return ok([directory.serialize_location(l, counts) for l in directory.list_locations(ctx(), include_archived)])


@bp.post("/locations")
@contract(LocationCreate, summary="Create a location", tags=("locations",))
@authz.permission_required("location.manage")
def locations_create():
    data = body(LocationCreate)
    return ok(directory.serialize_location(directory.create_location(ctx(), data)), status=201)


@bp.get("/locations/<location_id>")
def locations_get(location_id):
    return ok(directory.serialize_location(get_or_404(Location, location_id, ctx(), "Location")))


@bp.patch("/locations/<location_id>")
@contract(LocationUpdate, summary="Update a location", tags=("locations",))
@authz.permission_required("location.manage")
def locations_update(location_id):
    location = get_or_404(Location, location_id, ctx(), "Location")
    data = body(LocationUpdate)
    return ok(directory.serialize_location(directory.update_location(ctx(), location, data, data.model_fields_set)))


@bp.delete("/locations/<location_id>")
@authz.permission_required("location.manage")
def locations_archive(location_id):
    location = get_or_404(Location, location_id, ctx(), "Location")
    directory.archive_location(ctx(), location)
    return ok({"archived": True})


# --------------------------------------------------------------------------- categories


@bp.get("/categories")
def categories_list():
    from flask import request

    include_archived = request.args.get("include_archived") in {"1", "true"}
    usage = {cid: int(n) for cid, n in db.session.query(WorkOrderCategory.category_id, func.count(WorkOrderCategory.id)).join(Category, Category.id == WorkOrderCategory.category_id).filter(Category.organization_id == ctx().organization.id).group_by(WorkOrderCategory.category_id).all()}
    return ok([directory.serialize_category(c, usage) for c in directory.list_categories(ctx(), include_archived)])


@bp.post("/categories")
@contract(CategoryCreate, summary="Create a category", tags=("categories",))
@authz.permission_required("category.manage")
def categories_create():
    data = body(CategoryCreate)
    return ok(directory.serialize_category(directory.create_category(ctx(), data)), status=201)


@bp.get("/categories/<category_id>")
def categories_get(category_id):
    category = get_or_404(Category, category_id, ctx(), "Category")
    usage = {category.id: WorkOrderCategory.query.filter_by(category_id=category.id).count()}
    return ok(directory.serialize_category(category, usage))


@bp.patch("/categories/<category_id>")
@contract(CategoryUpdate, summary="Update a category", tags=("categories",))
@authz.permission_required("category.manage")
def categories_update(category_id):
    category = get_or_404(Category, category_id, ctx(), "Category")
    data = body(CategoryUpdate)
    return ok(directory.serialize_category(directory.update_category(ctx(), category, data, data.model_fields_set)))


@bp.delete("/categories/<category_id>")
@authz.permission_required("category.manage")
def categories_archive(category_id):
    category = get_or_404(Category, category_id, ctx(), "Category")
    directory.archive_category(ctx(), category)
    return ok({"archived": True})


# --------------------------------------------------------------------------- lookups


@bp.get("/ops/lookups")
def lookups():
    """Everything the selectors need in one round trip."""
    c = ctx()
    memberships = _memberships_by_user(c)
    users = [user_ref(m.user, m) for m in organizations.list_members(c.organization)]
    teams = [{"id": t.id, "name": t.name, "color": t.color, "project_id": t.project_id, "lead_ids": sorted(t.lead_user_ids), "member_ids": sorted(t.member_user_ids)} for t in directory.list_teams(c)]
    locations = [{"id": l.id, "name": l.name, "parent_location_id": l.parent_location_id, "is_default": bool(l.is_default)} for l in directory.list_locations(c)]
    categories = [{"id": x.id, "name": x.name, "color": x.color, "icon": x.icon} for x in directory.list_categories(c)]
    projects = [{"id": p.id, "name": p.name, "code": p.code, "status": p.status, "lead_user_id": p.lead_user_id} for p in OpsProject.query.filter(OpsProject.organization_id == c.organization.id, OpsProject.archived_at.is_(None)).order_by(OpsProject.name).all()]
    assets = [{"id": a.id, "name": a.name, "code": a.code, "parent_asset_id": a.parent_asset_id, "project_id": a.project_id, "location_id": a.location_id, "status": a.status} for a in Asset.query.filter(Asset.organization_id == c.organization.id, Asset.archived_at.is_(None)).order_by(Asset.name).all()]
    return ok({"users": users, "teams": teams, "locations": locations, "categories": categories, "projects": projects, "assets": assets, "roles": [{"key": r.system_key, "name": r.name} for r in organizations.list_roles(c.organization)]})
