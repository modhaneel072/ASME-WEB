"""Projects, members, milestones, health, activity."""

from __future__ import annotations

from asme.ops import authz
from asme.ops.api import body, bp, contract, ctx, ok
from asme.ops.api.query import encode_cursor, parse_list_query
from asme.ops.models import Membership, Milestone, OpsProject
from asme.ops.schemas import MilestoneCreate, MilestoneUpdate, ProjectCreate, ProjectListItem, ProjectMembersSet, ProjectUpdate
from asme.ops.services import files
from asme.ops.services import projects as projects_service
from asme.ops.tenancy import get_or_404

PROJECT_FILTERS = {"lead", "status", "risk", "year", "team"}
PROJECT_SORTS = {"name_asc", "updated_desc", "target_asc", "risk_desc"}


def _memberships(c):
    return {m.user_id: m for m in Membership.query.filter_by(organization_id=c.organization.id).all()}


@bp.get("/projects")
@contract(None, ProjectListItem, summary="List projects", tags=("projects",))
@authz.permission_required("project.read")
def projects_list():
    query = parse_list_query(allowed_filters=PROJECT_FILTERS, allowed_sorts=PROJECT_SORTS, default_sort="name_asc", allowed_extra={"view"})
    rows, counts, team_counts, next_cursor = projects_service.list_projects(ctx(), query)
    memberships = _memberships(ctx())
    return ok({"items": [projects_service.serialize_project_list_item(p, counts, team_counts, memberships) for p in rows], "next_cursor": encode_cursor(next_cursor) if next_cursor else None})


@bp.post("/projects")
@contract(ProjectCreate, summary="Create a project", tags=("projects",))
@authz.permission_required("project.create")
def projects_create():
    data = body(ProjectCreate)
    project = projects_service.create_project(ctx(), data)
    return ok(projects_service.serialize_project(project, ctx(), _memberships(ctx())), status=201)


@bp.get("/projects/<project_id>")
@authz.permission_required("project.read")
def projects_get(project_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    payload = projects_service.serialize_project(project, ctx(), _memberships(ctx()))
    payload["permissions"] = {"edit": authz.can(ctx(), "project.manage", record=project), "create_work_order": authz.can_create_for(ctx(), "work_order.create", project_id=project.id, team_id=None)}
    payload["files"] = [files.serialize(f) for f in files.list_for(ctx(), "project", project.id)]
    return ok(payload)


@bp.patch("/projects/<project_id>")
@contract(ProjectUpdate, summary="Update a project", tags=("projects",))
def projects_update(project_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    authz.require(ctx(), "project.manage", record=project)
    data = body(ProjectUpdate)
    project = projects_service.update_project(ctx(), project, data, data.model_fields_set)
    return ok(projects_service.serialize_project(project, ctx(), _memberships(ctx())))


@bp.post("/projects/<project_id>/archive")
def projects_archive(project_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    authz.require(ctx(), "project.manage", record=project)
    projects_service.archive_project(ctx(), project)
    return ok({"archived": True})


@bp.put("/projects/<project_id>/members")
@contract(ProjectMembersSet, summary="Replace project membership", tags=("projects",))
def projects_members(project_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    authz.require(ctx(), "project.manage", record=project)
    data = body(ProjectMembersSet)
    project = projects_service.set_members(ctx(), project, data.members)
    return ok(projects_service.serialize_project(project, ctx(), _memberships(ctx())))


@bp.get("/projects/<project_id>/health")
@authz.permission_required("project.read")
def projects_health(project_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    return ok(projects_service.health(ctx(), project))


@bp.get("/projects/<project_id>/activity")
@authz.permission_required("project.read")
def projects_activity(project_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    return ok(projects_service.activity(ctx(), project))


@bp.post("/projects/<project_id>/milestones")
@contract(MilestoneCreate, summary="Add a milestone", tags=("projects",))
def milestones_create(project_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    authz.require(ctx(), "project.manage", record=project)
    data = body(MilestoneCreate)
    return ok(projects_service.serialize_milestone(projects_service.create_milestone(ctx(), project, data)), status=201)


@bp.patch("/projects/<project_id>/milestones/<milestone_id>")
@contract(MilestoneUpdate, summary="Update a milestone", tags=("projects",))
def milestones_update(project_id, milestone_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    authz.require(ctx(), "project.manage", record=project)
    milestone = get_or_404(Milestone, milestone_id, ctx(), "Milestone")
    data = body(MilestoneUpdate)
    return ok(projects_service.serialize_milestone(projects_service.update_milestone(ctx(), project, milestone, data, data.model_fields_set)))


@bp.delete("/projects/<project_id>/milestones/<milestone_id>")
def milestones_delete(project_id, milestone_id):
    project = get_or_404(OpsProject, project_id, ctx(), "Project")
    authz.require(ctx(), "project.manage", record=project)
    milestone = get_or_404(Milestone, milestone_id, ctx(), "Milestone")
    projects_service.delete_milestone(ctx(), project, milestone)
    return ok({"deleted": True})
