from datetime import date, timedelta

import pytest

from asme.ops.api.query import ListQuery
from asme.ops.models import AuditEvent, OpsProjectMember
from asme.ops.schemas import (
    AssetCreate,
    AssetStatusChange,
    CategoryCreate,
    LocationCreate,
    LocationUpdate,
    MilestoneCreate,
    MilestoneUpdate,
    ProjectCreate,
    ProjectMemberInput,
    ProjectUpdate,
    TeamCreate,
    TeamMemberInput,
    WorkOrderComplete,
    WorkOrderCreate,
)
from asme.ops.services import assets as assets_service
from asme.ops.services import directory, projects, setup_center
from asme.ops.services import work_orders as wo_service
from asme.services.errors import Conflict, Validation


def test_create_project_assigns_lead_membership_and_audits(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    project = projects.create_project(admin, ProjectCreate(name="Crater Cruncher Rover", lead_user_id=people["project_lead"].id, target_date=date.today() + timedelta(days=200), budget_amount="12000"))
    assert project.code == "CCR"
    assert OpsProjectMember.query.filter_by(project_id=project.id, user_id=people["project_lead"].id, project_role="lead").count() == 1
    assert AuditEvent.query.filter_by(entity_type="project", entity_id=project.id).count() == 1
    with pytest.raises(Conflict):
        projects.create_project(admin, ProjectCreate(name="Other", code="ccr"))


def test_project_health_and_completion(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    project = projects.create_project(admin, ProjectCreate(name="Rover", code="RV", lead_user_id=people["project_lead"].id))
    m1 = projects.create_milestone(admin, project, MilestoneCreate(name="CDR", weight=1, status="complete"))
    projects.create_milestone(admin, project, MilestoneCreate(name="Drive test", weight=3, due_date=date.today() + timedelta(days=5)))
    wo = wo_service.create(admin, WorkOrderCreate(title="Fix", project_id=project.id, due_at=__import__("datetime").datetime.utcnow() - timedelta(days=1)))
    done = wo_service.create(admin, WorkOrderCreate(title="Done", project_id=project.id))
    wo_service.start(admin, done)
    wo_service.complete(admin, done, WorkOrderComplete(time_minutes=90, costs=[{"type": "part", "amount": "40"}]))
    health = projects.health(admin, project)
    assert health["completion_percent"] == 25  # 1 of 4 weight
    assert health["work_orders"]["open"] == 1 and health["work_orders"]["overdue"] == 1 and health["work_orders"]["done"] == 1
    assert health["milestones"]["at_risk"][0]["name"] == "Drive test"
    assert health["budget"]["used"] == 40.0 and health["hours_logged"] == 1.5
    projects.update_milestone(admin, project, m1, MilestoneUpdate(status="planned"), {"status"})
    assert projects.health(admin, project)["completion_percent"] == 0


def test_project_membership_keeps_lead(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    project = projects.create_project(admin, ProjectCreate(name="Rover", code="RV", lead_user_id=people["project_lead"].id))
    with pytest.raises(Validation):
        projects.set_members(admin, project, [ProjectMemberInput(user_id=people["full_member"].id)])
    projects.set_members(admin, project, [ProjectMemberInput(user_id=people["project_lead"].id, project_role="lead"), ProjectMemberInput(user_id=people["full_member"].id)])
    assert project.member_user_ids == {people["project_lead"].id, people["full_member"].id}


def test_project_views_and_activity(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    active = projects.create_project(admin, ProjectCreate(name="Active", code="ACT"))
    archived = projects.create_project(admin, ProjectCreate(name="Old", code="OLD"))
    projects.archive_project(admin, archived)
    rows, *_ = projects.list_projects(admin, ListQuery(extra={"view": "active"}, limit=50))
    assert [p.code for p in rows] == ["ACT"]
    rows, *_ = projects.list_projects(admin, ListQuery(extra={"view": "archived"}, limit=50))
    assert [p.code for p in rows] == ["OLD"]
    projects.update_project(admin, active, ProjectUpdate(risk_level="high"), {"risk_level"})
    events = projects.activity(admin, active)
    assert events[0]["event_type"] == "project.updated" and events[0]["metadata"]["changed"]["risk_level"]["to"] == "high"


def test_teams_locations_categories(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    team = directory.create_team(admin, TeamCreate(name="Wheels", members=[TeamMemberInput(user_id=people["team_lead"].id, is_lead=True)]))
    assert team.lead_user_ids == {people["team_lead"].id}
    with pytest.raises(Conflict):
        directory.create_team(admin, TeamCreate(name="wheels"))
    directory.set_team_members(admin, team, [TeamMemberInput(user_id=people["full_member"].id)])
    assert team.member_user_ids == {people["full_member"].id} and team.lead_user_ids == set()

    default = directory.default_location(admin)
    assert default.is_default and default.name == "General"
    lab = directory.create_location(admin, LocationCreate(name="Robotics Lab", building="Seamans"))
    bench = directory.create_location(admin, LocationCreate(name="Bench", parent_location_id=lab.id))
    assert bench.parent.id == lab.id
    with pytest.raises(Validation):
        directory.update_location(admin, lab, LocationUpdate(parent_location_id=lab.id), {"parent_location_id"})
    directory.update_location(admin, lab, LocationUpdate(is_default=True), {"is_default"})
    assert lab.is_default and not directory.default_location.__wrapped__(admin).is_default if hasattr(directory.default_location, "__wrapped__") else lab.is_default
    with pytest.raises(Conflict):
        directory.archive_location(admin, lab)

    assert len(directory.list_categories(admin)) == 14
    cat = directory.create_category(admin, CategoryCreate(name="Testing", color="#123456", icon="beaker"))
    assert cat.color == "#123456"
    with pytest.raises(Conflict):
        directory.create_category(admin, CategoryCreate(name="mechanical"))


def test_asset_hierarchy_and_status_history(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    rover = assets_service.create_asset(admin, AssetCreate(name="Rover", code="ROVER", criticality="critical"))
    arm = assets_service.create_asset(admin, AssetCreate(name="Arm", parent_asset_id=rover.id))
    assert arm.code == "ARM" and arm.location_id is not None  # default location applied
    from asme.ops.schemas import AssetUpdate

    with pytest.raises(Validation):
        assets_service.update_asset(admin, rover, AssetUpdate(parent_asset_id=arm.id), {"parent_asset_id"})
    assets_service.change_status(admin, rover, AssetStatusChange(status="OFFLINE_UNPLANNED", downtime_reason="Motor failure"))
    assets_service.change_status(admin, rover, AssetStatusChange(status="ONLINE"))
    from asme.ops.models import AssetStatusHistory

    rows = AssetStatusHistory.query.filter_by(asset_id=rover.id).order_by(AssetStatusHistory.started_at).all()
    assert [r.to_status for r in rows] == ["ONLINE", "OFFLINE_UNPLANNED", "ONLINE"]
    assert rows[1].ended_at is not None and rows[1].downtime_type == "unplanned"
    assert assets_service.child_counts(admin)[rover.id] == 1


def test_setup_center_reflects_persisted_rows(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    before = setup_center.evaluate(admin)
    tasks = {t["key"]: t for phase in before["phases"] for t in phase["tasks"]}
    assert tasks["locations"]["complete"] is False
    assert tasks["assets"]["target"] == 5 and tasks["assets"]["complete"] is False
    assert tasks["teams_users"]["complete"] is False  # no team yet
    assert tasks["categories"]["complete"] is True  # seeded defaults
    assert tasks["parts"]["available"] is False and before["next_step"]["task"] == "chapter_profile"
    directory.create_location(admin, LocationCreate(name="Lab"))
    directory.create_team(admin, TeamCreate(name="Board"))
    for i in range(5):
        assets_service.create_asset(admin, AssetCreate(name=f"Asset {i}"))
    setup_center.confirm_profile(admin)
    after = setup_center.evaluate(admin)
    tasks = {t["key"]: t for phase in after["phases"] for t in phase["tasks"]}
    assert all(tasks[k]["complete"] for k in ("chapter_profile", "locations", "assets", "teams_users", "categories"))
    assert after["phases"][0]["complete"] is True
    assert after["percent"] > before["percent"]
