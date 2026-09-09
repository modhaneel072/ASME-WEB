"""The demo seed is coherent and idempotent."""

from asme.ops.models import Membership, OpsProject, Team, WorkOrder
from asme.ops.seeds import seed_ops_demo


def test_demo_seed_is_idempotent_and_coherent(app, db):
    first = seed_ops_demo()
    second = seed_ops_demo()
    assert first["work_orders"] == second["work_orders"] >= 18
    ccr = OpsProject.query.filter_by(code="CCR").first()
    assert ccr is not None and ccr.lead is not None and len(ccr.milestones) == 6
    statuses = {wo.status for wo in WorkOrder.query.all()}
    assert {"OPEN", "IN_PROGRESS", "ON_HOLD", "DONE", "CANCELED", "DRAFT"} <= statuses
    assert WorkOrder.query.filter(WorkOrder.priority == "CRITICAL").count() >= 1
    assert WorkOrder.query.filter(WorkOrder.parent_work_order_id.isnot(None)).count() == 3
    assert WorkOrder.query.filter(WorkOrder.recurring_rule_json.isnot(None)).count() >= 1
    roles = {m.role.system_key for m in Membership.query.all()}
    assert {"chapter_admin", "project_lead", "team_lead", "full_member", "safety_officer", "inventory_manager", "treasurer", "faculty_advisor", "shop_operator", "executive_officer"} <= roles
    assert Team.query.filter_by(name="Wheels and Mobility").first().parent.name == "Crater Cruncher Rover"
