"""Pins the formulas documented in docs/reporting-metrics.md."""

from __future__ import annotations

from datetime import datetime, timedelta

from asme.extensions import db
from asme.ops.schemas import CostInput, ProjectCreate, TimeEntryCreate, WorkOrderComplete, WorkOrderCreate, WorkOrderUpdate
from asme.ops.services import dashboard, projects as projects_service, work_orders as wo_service


def _make(ctx, title, **kw):
    return wo_service.create(ctx, WorkOrderCreate(title=title, **kw))


def test_operations_dashboard_counts_and_rates(people, ctx_for, lookups):
    admin = ctx_for(people["chapter_admin"])
    now = datetime.utcnow()
    project = projects_service.create_project(admin, ProjectCreate(name="Dash Project", code="DASH"))
    team = lookups["team"]

    overdue = _make(admin, "Overdue", due_at=now - timedelta(days=1), team_id=team.id, project_id=project.id, assignee_user_ids=[people["full_member"].id])
    soon = _make(admin, "Due soon", due_at=now + timedelta(days=3), project_id=project.id)
    blocked = _make(admin, "Blocked", project_id=project.id)
    wo_service.update(admin, blocked, WorkOrderUpdate(is_blocked=True), {"is_blocked"})
    done_on_time = _make(admin, "Done on time", due_at=now + timedelta(days=2), project_id=project.id, work_type="PREVENTIVE", recurrence={"frequency": "weekly", "interval": 1, "mode": "fixed"})
    wo_service.start(admin, done_on_time)
    wo_service.complete(admin, done_on_time, WorkOrderComplete(time_minutes=90, costs=[CostInput(type="part", amount=12.5), CostInput(type="labor", amount=7.5)]))
    late = _make(admin, "Done late", due_at=now - timedelta(days=3), project_id=project.id)
    wo_service.start(admin, late)
    wo_service.complete(admin, late, WorkOrderComplete())
    wo_service.add_time_entry(admin, overdue, TimeEntryCreate(minutes=30))
    db.session.commit()

    data = dashboard.operations(admin, "30d", project.id)
    totals = data["totals"]
    # the completed recurring work order generated a next occurrence, which is open with a due date next week
    assert totals["overdue"] == 1
    assert totals["blocked"] == 1
    assert totals["due_soon"] >= 1
    assert totals["completed"] == 2
    assert totals["created"] == 6  # 5 explicit + 1 generated occurrence
    assert totals["open"] == totals["created"] - totals["completed"]
    assert data["on_time_completion_rate"] == 50
    assert data["average_completion_hours"] is not None and data["average_completion_hours"] >= 0
    assert data["by_status"]["DONE"] == 2
    assert data["by_priority"]["NONE"] == totals["open"]
    assert data["by_work_type"]["PREVENTIVE"] == 2
    assert data["repeating"] == {"repeating": 2, "non_repeating": 4}
    assert [w for w in data["workload_by_team"] if w["team_id"] == team.id][0]["open"] == 1
    assert [w for w in data["workload_by_user"] if w["user_id"] == people["full_member"].id][0]["open"] == 1
    assert data["hours_logged"] == 2.0  # 90 + 30 minutes
    assert data["costs"] == {"parts": 12.5, "labor": 7.5, "vendor": 0.0, "other": 0.0, "total": 20.0}
    assert len(data["created_vs_completed"]) == 4
    assert sum(b["completed"] for b in data["created_vs_completed"]) == 2
    assert sum(b["created"] for b in data["created_vs_completed"]) == totals["created"]
    assert data["range"]["days"] == 30

    other = dashboard.operations(admin, "7d", None)
    assert other["range"]["days"] == 7 and len(other["created_vs_completed"]) == 1
    assert other["totals"]["open"] >= totals["open"]


def test_project_health_budget_and_milestones(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    from datetime import date
    from decimal import Decimal

    from asme.ops.schemas import MilestoneCreate

    project = projects_service.create_project(admin, ProjectCreate(name="Health Project", code="HLTH", budget_amount=Decimal("100")))
    projects_service.create_milestone(admin, project, MilestoneCreate(name="Done one", status="complete", weight=3, due_date=date.today() - timedelta(days=10)))
    projects_service.create_milestone(admin, project, MilestoneCreate(name="Soon", status="planned", weight=1, due_date=date.today() + timedelta(days=5)))
    a = _make(admin, "A", project_id=project.id)
    wo_service.start(admin, a)
    wo_service.complete(admin, a, WorkOrderComplete(time_minutes=60, costs=[CostInput(type="vendor", amount=40)]))
    _make(admin, "B", project_id=project.id, due_at=datetime.utcnow() - timedelta(days=1))
    db.session.commit()

    health = projects_service.health(admin, project)
    # milestones exist, so completion is the weighted milestone share: 3 / (3 + 1) = 75%
    assert health["completion_percent"] == 75
    assert health["work_orders"]["open"] == 1 and health["work_orders"]["done"] == 1 and health["work_orders"]["overdue"] == 1
    assert health["milestones"]["total"] == 2 and health["milestones"]["complete"] == 1
    assert [m["name"] for m in health["milestones"]["at_risk"]] == ["Soon"]
    assert health["milestones"]["next"]["name"] == "Soon"
    assert health["budget"] == {"amount": 100.0, "used": 40.0, "remaining": 60.0}
    assert health["hours_logged"] == 1.0
