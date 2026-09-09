"""Work orders."""

from __future__ import annotations

from flask import request

from asme.ops import authz
from asme.ops.api import body, bp, contract, ctx, ok
from asme.ops.api.query import encode_cursor, parse_list_query
from asme.ops.models import Comment, Membership, WorkOrder
from asme.ops.schemas import (
    AssigneesSet,
    CommentCreate,
    CommentUpdate,
    CostInput,
    StatusNote,
    SubWorkOrderCreate,
    TimeEntryCreate,
    WatchersSet,
    WorkOrderComplete,
    WorkOrderCreate,
    WorkOrderListItem,
    WorkOrderOut,
    WorkOrderUpdate,
)
from asme.ops.services import comments as comments_service
from asme.ops.services import files
from asme.ops.services import work_orders as wo_service
from asme.ops.tenancy import get_or_404
from asme.services.errors import NotFound

WO_FILTERS = {"status", "priority", "work_type", "project", "team", "location", "asset", "category", "assignee", "due", "blocked", "parent", "created_by"}
WO_SORTS = set(wo_service.SORTS)


def _memberships(c):
    return {m.user_id: m for m in Membership.query.filter_by(organization_id=c.organization.id).all()}


def _wo(work_order_id) -> WorkOrder:
    return wo_service.get_visible_or_404(ctx(), work_order_id)


@bp.get("/work-orders")
@contract(None, WorkOrderListItem, summary="List work orders (To Do / Done, filters, sort, cursor)", tags=("work-orders",))
@authz.permission_required("work_order.read_assigned")
def work_orders_list():
    query = parse_list_query(allowed_filters=WO_FILTERS, allowed_sorts=WO_SORTS, default_sort="due_asc", allowed_extra={"tab"})
    rows, comment_counts, child_counts, next_cursor = wo_service.list_work_orders(ctx(), query)
    memberships = _memberships(ctx())
    project_id = (query.filters.get("project") or [None])[0]
    return ok(
        {
            "items": [wo_service.serialize_list_item(w, comment_counts=comment_counts, child_counts=child_counts, memberships=memberships) for w in rows],
            "next_cursor": encode_cursor(next_cursor) if next_cursor else None,
            "counts": wo_service.counts_for_tabs(ctx(), project_id if project_id and project_id != "none" else None),
        }
    )


@bp.post("/work-orders")
@contract(WorkOrderCreate, WorkOrderOut, summary="Create a work order", tags=("work-orders",))
def work_orders_create():
    data = body(WorkOrderCreate)
    wo = wo_service.create(ctx(), data)
    return ok(wo_service.serialize_detail(ctx(), wo), status=201)


@bp.get("/work-orders/<work_order_id>")
@contract(None, WorkOrderOut, summary="Work-order detail", tags=("work-orders",))
def work_orders_get(work_order_id):
    wo = _wo(work_order_id)
    payload = wo_service.serialize_detail(ctx(), wo)
    payload["comments"] = [comments_service.serialize(c) for c in comments_service.list_for(ctx(), "work_order", wo.id)]
    payload["files"] = [files.serialize(f) for f in files.list_for(ctx(), "work_order", wo.id)]
    return ok(payload)


@bp.patch("/work-orders/<work_order_id>")
@contract(WorkOrderUpdate, WorkOrderOut, summary="Edit work-order fields", tags=("work-orders",))
def work_orders_update(work_order_id):
    wo = _wo(work_order_id)
    data = body(WorkOrderUpdate)
    wo = wo_service.update(ctx(), wo, data, data.model_fields_set)
    return ok(wo_service.serialize_detail(ctx(), wo))


def _status_action(work_order_id, fn):
    wo = _wo(work_order_id)
    data = body(StatusNote)
    wo = fn(ctx(), wo, data.note)
    return ok(wo_service.serialize_detail(ctx(), wo))


@bp.post("/work-orders/<work_order_id>/start")
@contract(StatusNote, WorkOrderOut, summary="Start work", tags=("work-orders",))
def work_orders_start(work_order_id):
    return _status_action(work_order_id, wo_service.start)


@bp.post("/work-orders/<work_order_id>/hold")
@contract(StatusNote, WorkOrderOut, summary="Put on hold", tags=("work-orders",))
def work_orders_hold(work_order_id):
    return _status_action(work_order_id, wo_service.hold)


@bp.post("/work-orders/<work_order_id>/resume")
@contract(StatusNote, WorkOrderOut, summary="Resume", tags=("work-orders",))
def work_orders_resume(work_order_id):
    return _status_action(work_order_id, wo_service.resume)


@bp.post("/work-orders/<work_order_id>/open")
@contract(StatusNote, WorkOrderOut, summary="Open a draft", tags=("work-orders",))
def work_orders_open(work_order_id):
    wo = _wo(work_order_id)
    wo = wo_service.open_draft(ctx(), wo)
    return ok(wo_service.serialize_detail(ctx(), wo))


@bp.post("/work-orders/<work_order_id>/cancel")
@contract(StatusNote, WorkOrderOut, summary="Cancel", tags=("work-orders",))
def work_orders_cancel(work_order_id):
    return _status_action(work_order_id, wo_service.cancel)


@bp.post("/work-orders/<work_order_id>/complete")
@contract(WorkOrderComplete, WorkOrderOut, summary="Complete with note, time, costs, asset status and follow-up", tags=("work-orders",))
def work_orders_complete(work_order_id):
    wo = _wo(work_order_id)
    data = body(WorkOrderComplete)
    wo, follow_up, next_occurrence = wo_service.complete(ctx(), wo, data)
    payload = wo_service.serialize_detail(ctx(), wo)
    payload["follow_up"] = wo_service.serialize_list_item(follow_up) if follow_up else None
    payload["next_occurrence"] = wo_service.serialize_list_item(next_occurrence) if next_occurrence else None
    return ok(payload)


@bp.post("/work-orders/<work_order_id>/duplicate")
@contract(None, WorkOrderOut, summary="Duplicate", tags=("work-orders",))
def work_orders_duplicate(work_order_id):
    wo = _wo(work_order_id)
    copy = wo_service.duplicate(ctx(), wo)
    return ok(wo_service.serialize_detail(ctx(), copy), status=201)


@bp.post("/work-orders/<work_order_id>/sub-work-orders")
@contract(SubWorkOrderCreate, WorkOrderOut, summary="Create a sub-work order", tags=("work-orders",))
def work_orders_sub(work_order_id):
    parent = _wo(work_order_id)
    data = body(SubWorkOrderCreate)
    child = wo_service.create_sub(ctx(), parent, data)
    return ok(wo_service.serialize_detail(ctx(), child), status=201)


@bp.put("/work-orders/<work_order_id>/assignees")
@contract(AssigneesSet, WorkOrderOut, summary="Replace assignees", tags=("work-orders",))
def work_orders_assignees(work_order_id):
    wo = _wo(work_order_id)
    data = body(AssigneesSet)
    wo = wo_service.set_assignees(ctx(), wo, data.user_ids, data.team_ids)
    return ok(wo_service.serialize_detail(ctx(), wo))


@bp.put("/work-orders/<work_order_id>/watchers")
@contract(WatchersSet, WorkOrderOut, summary="Replace watchers", tags=("work-orders",))
def work_orders_watchers(work_order_id):
    wo = _wo(work_order_id)
    data = body(WatchersSet)
    wo = wo_service.set_watchers(ctx(), wo, data.user_ids)
    return ok(wo_service.serialize_detail(ctx(), wo))


@bp.get("/work-orders/<work_order_id>/history")
def work_orders_history(work_order_id):
    wo = _wo(work_order_id)
    from asme.ops.services import audit_log

    rows, _more = audit_log.list_events(ctx(), entity_type="work_order", entity_id=wo.id, limit=200)
    return ok({"status_history": [wo_service.serialize_history(h) for h in wo.status_history], "events": [audit_log.serialize_event(r) for r in rows]})


# --------------------------------------------------------------------------- comments


@bp.get("/work-orders/<work_order_id>/comments")
def work_orders_comments(work_order_id):
    wo = _wo(work_order_id)
    return ok([comments_service.serialize(c) for c in comments_service.list_for(ctx(), "work_order", wo.id)])


@bp.post("/work-orders/<work_order_id>/comments")
@contract(CommentCreate, summary="Comment on a work order", tags=("work-orders",))
def work_orders_comment_create(work_order_id):
    wo = _wo(work_order_id)
    data = body(CommentCreate)
    comment = wo_service.add_comment(ctx(), wo, data.body, data.parent_comment_id)
    return ok(comments_service.serialize(comment), status=201)


@bp.patch("/comments/<comment_id>")
@contract(CommentUpdate, summary="Edit a comment", tags=("work-orders",))
def comments_update(comment_id):
    comment = Comment.query.filter_by(id=comment_id, organization_id=ctx().organization.id).first()
    if comment is None:
        raise NotFound("Comment not found.")
    data = body(CommentUpdate)
    return ok(comments_service.serialize(comments_service.edit(ctx(), comment, data.body)))


@bp.delete("/comments/<comment_id>")
def comments_delete(comment_id):
    comment = Comment.query.filter_by(id=comment_id, organization_id=ctx().organization.id).first()
    if comment is None:
        raise NotFound("Comment not found.")
    comments_service.remove(ctx(), comment)
    return ok({"deleted": True})


# --------------------------------------------------------------------------- time & cost


@bp.get("/work-orders/<work_order_id>/time-entries")
def work_orders_time_list(work_order_id):
    wo = _wo(work_order_id)
    return ok([wo_service.serialize_time_entry(t) for t in sorted(wo.time_entries, key=lambda t: t.created_at)])


@bp.post("/work-orders/<work_order_id>/time-entries")
@contract(TimeEntryCreate, summary="Log time", tags=("work-orders",))
def work_orders_time_create(work_order_id):
    wo = _wo(work_order_id)
    data = body(TimeEntryCreate)
    entry = wo_service.add_time_entry(ctx(), wo, data)
    return ok({"entry": wo_service.serialize_time_entry(entry), "actual_minutes": wo.actual_minutes}, status=201)


@bp.get("/work-orders/<work_order_id>/cost-entries")
def work_orders_cost_list(work_order_id):
    wo = _wo(work_order_id)
    return ok([wo_service.serialize_cost_entry(c) for c in sorted(wo.cost_entries, key=lambda c: c.created_at)])


@bp.post("/work-orders/<work_order_id>/cost-entries")
@contract(CostInput, summary="Log a cost", tags=("work-orders",))
def work_orders_cost_create(work_order_id):
    wo = _wo(work_order_id)
    data = body(CostInput)
    entry = wo_service.add_cost_entry(ctx(), wo, data)
    return ok({"entry": wo_service.serialize_cost_entry(entry), "total_cost": round(float(sum((c.amount for c in wo.cost_entries), 0)), 2)}, status=201)
