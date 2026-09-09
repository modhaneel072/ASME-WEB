from datetime import datetime, timedelta

import pytest

from asme.ops.api.query import ListQuery
from asme.ops.models import Notification, WorkOrder, WorkOrderStatusHistory
from asme.ops.schemas import CostInput, SubWorkOrderCreate, TimeEntryCreate, WorkOrderComplete, WorkOrderCreate, WorkOrderUpdate
from asme.ops.services import work_orders as svc
from asme.services.errors import Conflict, Validation


def test_numbers_are_sequential_per_org(people, ctx_for, other_org):
    admin = ctx_for(people["chapter_admin"])
    a = svc.create(admin, WorkOrderCreate(title="First"))
    b = svc.create(admin, WorkOrderCreate(title="Second"))
    foreign = svc.create(other_org["ctx"], WorkOrderCreate(title="Foreign first"))
    assert (a.number, b.number, foreign.number) == (1, 2, 1)


def test_create_records_history_watchers_and_audit(people, ctx_for, lookups):
    admin = ctx_for(people["chapter_admin"])
    wo = svc.create(admin, WorkOrderCreate(title="Inspect hubs", project_id=lookups["project"].id, team_id=lookups["team"].id, primary_asset_id=lookups["asset"].id, category_ids=[lookups["categories"][0].id], assignee_user_ids=[people["full_member"].id], priority="HIGH"))
    assert wo.status == "OPEN" and wo.priority_rank == 3
    assert [h.to_status for h in wo.status_history] == ["OPEN"]
    assert people["chapter_admin"].id in {w.user_id for w in wo.watchers}
    assert any(l.relationship_type == "primary" and l.asset_id == lookups["asset"].id for l in wo.asset_links)
    from asme.ops.models import AuditEvent

    assert AuditEvent.query.filter_by(entity_type="work_order", entity_id=wo.id, event_type="work_order.created").count() == 1
    # assignee got a notification, the actor did not
    notes = Notification.query.filter_by(entity_id=wo.id).all()
    assert {n.user_id for n in notes} == {people["full_member"].id}


def test_lifecycle_transitions(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    wo = svc.create(admin, WorkOrderCreate(title="Lifecycle"))
    svc.start(admin, wo)
    assert wo.status == "IN_PROGRESS" and wo.start_at is not None
    svc.hold(admin, wo, "waiting on parts")
    assert wo.status == "ON_HOLD"
    svc.resume(admin, wo)
    assert wo.status == "IN_PROGRESS"
    with pytest.raises(Conflict) as excinfo:
        svc.hold(admin, svc.create(admin, WorkOrderCreate(title="Done already", status="OPEN")).__class__.query.filter_by(title="Lifecycle").first(), None) if False else svc.resume(admin, wo)
    assert excinfo.value.code == "invalid_transition"
    svc.complete(admin, wo, WorkOrderComplete(completion_note="done"))
    assert wo.status == "DONE" and wo.completed_at is not None
    assert [h.to_status for h in wo.status_history] == ["OPEN", "IN_PROGRESS", "ON_HOLD", "IN_PROGRESS", "DONE"]
    with pytest.raises(Conflict):
        svc.update(admin, wo, WorkOrderUpdate(title="Edit closed"), {"title"})


def test_draft_opens_then_quick_completes(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    draft = svc.create(admin, WorkOrderCreate(title="Draft", status="DRAFT"))
    with pytest.raises(Conflict):
        svc.complete(admin, draft, WorkOrderComplete())
    svc.open_draft(admin, draft)
    svc.complete(admin, draft, WorkOrderComplete(completion_note="quick"))
    assert draft.status == "DONE"


def test_cancel_records_reason(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    wo = svc.create(admin, WorkOrderCreate(title="Cancel me"))
    svc.cancel(admin, wo, "superseded")
    assert wo.status == "CANCELED" and wo.cancel_reason == "superseded" and wo.canceled_at


def test_complete_aggregates_time_costs_and_asset_status(people, ctx_for, lookups):
    admin = ctx_for(people["chapter_admin"])
    wo = svc.create(admin, WorkOrderCreate(title="Fix rover", primary_asset_id=lookups["asset"].id))
    svc.start(admin, wo)
    svc.add_time_entry(admin, wo, TimeEntryCreate(minutes=30))
    wo, follow_up, nxt = svc.complete(admin, wo, WorkOrderComplete(completion_note="fixed", time_minutes=45, costs=[CostInput(type="part", amount="12.50"), CostInput(type="other", amount="3")], asset_status="OFFLINE_PLANNED", follow_up_title="Re-check next week"))
    assert wo.actual_minutes == 75
    assert round(float(sum(c.amount for c in wo.cost_entries)), 2) == 15.5
    assert lookups["asset"].status == "OFFLINE_PLANNED"
    assert follow_up is not None and follow_up.title == "Re-check next week" and follow_up.primary_asset_id == lookups["asset"].id
    assert nxt is None


def test_recurring_completion_generates_next_exactly_once(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    due = datetime.utcnow() + timedelta(days=1)
    wo = svc.create(admin, WorkOrderCreate(title="Monthly printer inspection", due_at=due, recurrence={"frequency": "monthly", "interval": 1, "mode": "fixed"}))
    svc.start(admin, wo)
    _, _, nxt = svc.complete(admin, wo, WorkOrderComplete())
    assert nxt is not None and nxt.status == "OPEN" and nxt.generation_key
    assert (nxt.due_at - due).days == 30
    # retrying the generation step is a no-op
    assert svc._next_occurrence(wo) is None
    assert WorkOrder.query.filter_by(title="Monthly printer inspection").count() == 2


def test_sub_work_orders_and_auto_complete_parent(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    parent = svc.create(admin, WorkOrderCreate(title="Hub revision", parent_completion_policy="auto"))
    a = svc.create_sub(admin, parent, SubWorkOrderCreate(title="CAD"))
    b = svc.create_sub(admin, parent, SubWorkOrderCreate(title="Print"))
    assert a.parent_work_order_id == parent.id and a.number > parent.number
    with pytest.raises(Conflict) as excinfo:
        svc.complete(admin, parent, WorkOrderComplete())
    assert excinfo.value.code == "children_open"
    svc.start(admin, a)
    svc.complete(admin, a, WorkOrderComplete())
    assert parent.status == "OPEN"
    svc.start(admin, b)
    svc.complete(admin, b, WorkOrderComplete())
    parent = WorkOrder.query.get(parent.id)
    assert parent.status == "DONE"
    assert any(h.note and "Auto-completed" in h.note for h in parent.status_history)


def test_duplicate_copies_links_but_not_status(people, ctx_for, lookups):
    admin = ctx_for(people["chapter_admin"])
    wo = svc.create(admin, WorkOrderCreate(title="Original", project_id=lookups["project"].id, team_id=lookups["team"].id, category_ids=[lookups["categories"][0].id], assignee_user_ids=[people["full_member"].id], priority="HIGH"))
    svc.start(admin, wo)
    copy = svc.duplicate(admin, wo)
    assert copy.id != wo.id and copy.number == wo.number + 1 and copy.status == "OPEN"
    assert copy.project_id == wo.project_id and copy.category_ids == wo.category_ids and copy.assignee_user_ids == wo.assignee_user_ids


def test_critical_priority_is_permission_controlled(people, ctx_for):
    member = ctx_for(people["full_member"])
    with pytest.raises(Validation) as excinfo:
        svc.create(member, WorkOrderCreate(title="Emergency", priority="CRITICAL"))
    assert excinfo.value.code == "critical_forbidden"
    safety = ctx_for(people["safety_officer"])
    assert svc.create(safety, WorkOrderCreate(title="Emergency", priority="CRITICAL")).priority == "CRITICAL"


def test_due_before_start_rejected(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    now = datetime.utcnow()
    with pytest.raises(Validation):
        svc.create(admin, WorkOrderCreate(title="Bad dates", start_at=now, due_at=now - timedelta(days=1)))


def _seed_for_listing(people, ctx_for, lookups):
    admin = ctx_for(people["chapter_admin"])
    now = datetime.utcnow()
    rows = [
        svc.create(admin, WorkOrderCreate(title="Overdue high", priority="HIGH", due_at=now - timedelta(days=2), project_id=lookups["project"].id, team_id=lookups["team"].id, assignee_user_ids=[people["full_member"].id])),
        svc.create(admin, WorkOrderCreate(title="Soon medium", priority="MEDIUM", due_at=now + timedelta(days=3), category_ids=[lookups["categories"][1].id])),
        svc.create(admin, WorkOrderCreate(title="Later low", priority="LOW", due_at=now + timedelta(days=20), location_id=lookups["location"].id)),
        svc.create(admin, WorkOrderCreate(title="No due none", priority="NONE")),
    ]
    done = svc.create(admin, WorkOrderCreate(title="Finished", priority="LOW"))
    svc.start(admin, done)
    svc.complete(admin, done, WorkOrderComplete())
    return admin, rows, done


def test_list_tabs_filters_and_sorts(people, ctx_for, lookups):
    admin, rows, done = _seed_for_listing(people, ctx_for, lookups)
    todo, *_ = svc.list_work_orders(admin, ListQuery(sort="due_asc", limit=50, extra={"tab": "todo"}))
    assert [w.title for w in todo] == ["Overdue high", "Soon medium", "Later low", "No due none"]
    done_rows, *_ = svc.list_work_orders(admin, ListQuery(sort="due_asc", limit=50, extra={"tab": "done"}))
    assert [w.title for w in done_rows] == ["Finished"]
    by_priority, *_ = svc.list_work_orders(admin, ListQuery(sort="priority_desc", limit=50, extra={"tab": "todo"}))
    assert [w.priority for w in by_priority] == ["HIGH", "MEDIUM", "LOW", "NONE"]
    overdue, *_ = svc.list_work_orders(admin, ListQuery(filters={"due": ["overdue"]}, limit=50, extra={"tab": "todo"}))
    assert [w.title for w in overdue] == ["Overdue high"]
    mine, *_ = svc.list_work_orders(ctx_for(people["full_member"]), ListQuery(filters={"assignee": ["me"]}, limit=50, extra={"tab": "todo"}))
    assert [w.title for w in mine] == ["Overdue high"]
    by_cat, *_ = svc.list_work_orders(admin, ListQuery(filters={"category": [lookups["categories"][1].id]}, limit=50, extra={"tab": "todo"}))
    assert [w.title for w in by_cat] == ["Soon medium"]
    by_loc, *_ = svc.list_work_orders(admin, ListQuery(filters={"location": [lookups["location"].id]}, limit=50, extra={"tab": "todo"}))
    assert [w.title for w in by_loc] == ["Later low"]
    searched, *_ = svc.list_work_orders(admin, ListQuery(q="soon", limit=50, extra={"tab": "todo"}))
    assert [w.title for w in searched] == ["Soon medium"]
    by_number, *_ = svc.list_work_orders(admin, ListQuery(q=f"#{rows[2].number}", limit=50, extra={"tab": "all"}))
    assert [w.title for w in by_number] == ["Later low"]


def test_list_cursor_pages(people, ctx_for, lookups):
    admin, rows, done = _seed_for_listing(people, ctx_for, lookups)
    page1, _, _, cursor = svc.list_work_orders(admin, ListQuery(sort="number_asc", limit=2, extra={"tab": "todo"}))
    assert len(page1) == 2 and cursor == {"offset": 2}
    page2, _, _, cursor2 = svc.list_work_orders(admin, ListQuery(sort="number_asc", limit=2, cursor=cursor, extra={"tab": "todo"}))
    assert len(page2) == 2 and cursor2 is None
    assert {w.id for w in page1}.isdisjoint({w.id for w in page2})


def test_tab_counts(people, ctx_for, lookups):
    admin, rows, done = _seed_for_listing(people, ctx_for, lookups)
    assert svc.counts_for_tabs(admin) == {"todo": 4, "done": 1}


def test_time_entry_on_behalf_requires_assign_permission(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    member = ctx_for(people["full_member"])
    wo = svc.create(admin, WorkOrderCreate(title="Timed", assignee_user_ids=[people["full_member"].id]))
    svc.add_time_entry(member, wo, TimeEntryCreate(minutes=15))
    with pytest.raises(Validation):
        svc.add_time_entry(member, wo, TimeEntryCreate(minutes=15, user_id=people["team_lead"].id))
    svc.add_time_entry(admin, wo, TimeEntryCreate(minutes=10, user_id=people["full_member"].id))
    assert wo.actual_minutes == 25


def test_comment_notifies_watchers_and_mentions(people, ctx_for):
    admin = ctx_for(people["chapter_admin"])
    member = ctx_for(people["full_member"])
    wo = svc.create(admin, WorkOrderCreate(title="Discuss", assignee_user_ids=[people["full_member"].id]))
    svc.add_comment(member, wo, f"Looks good @[{people['team_lead'].id}]")
    recipients = {n.user_id for n in Notification.query.filter_by(entity_id=wo.id, type="work_order.comment").all()}
    assert recipients == {people["chapter_admin"].id, people["team_lead"].id}
