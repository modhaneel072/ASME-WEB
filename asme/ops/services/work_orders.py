"""Work orders: numbering, creation, lifecycle, assignment, sub-work orders, time & cost.

Every state change writes ``work_order_status_history`` and an audit event inside the
same transaction, then emits a domain event after commit.
"""

from __future__ import annotations

import json
from datetime import datetime, timedelta

from sqlalchemy import func, or_
from sqlalchemy.exc import IntegrityError

from asme import events
from asme.extensions import db
from asme.models import User
from asme.ops import audit, authz
from asme.ops.models import (
    WO_CLOSED_STATUSES,
    WO_OPEN_STATUSES,
    WO_PRIORITY_RANK,
    WO_TRANSITIONS,
    Asset,
    Category,
    CostEntry,
    Location,
    OpsProject,
    Team,
    TeamMember,
    TimeEntry,
    WorkOrder,
    WorkOrderAsset,
    WorkOrderAssignee,
    WorkOrderCategory,
    WorkOrderCounter,
    WorkOrderStatusHistory,
    WorkOrderWatcher,
)
from asme.ops.services import comments as comments_service
from asme.ops.services import notifications
from asme.ops.services.common import iso, ref, resolve_users, touch, user_ref
from asme.ops.tenancy import get_or_404, scoped
from asme.services.errors import Conflict, NotFound, Validation

WO_FIELDS = (
    "title", "description", "status", "priority", "work_type", "project_id", "team_id", "location_id", "primary_asset_id",
    "parent_work_order_id", "start_at", "due_at", "estimated_minutes", "budget_code", "is_blocked",
    "parent_completion_policy", "recurring_rule_json",
)
EVENT_CREATED = "ops.work_order.created"
EVENT_UPDATED = "ops.work_order.updated"
EVENT_STATUS = "ops.work_order.status_changed"
EVENT_ASSIGNED = "ops.work_order.assigned"


# --------------------------------------------------------------------------- numbering


def allocate_number(org_id: str) -> int:
    counter = db.session.get(WorkOrderCounter, org_id, with_for_update=True)
    if counter is None:
        counter = WorkOrderCounter(organization_id=org_id, next_number=1)
        db.session.add(counter)
        db.session.flush()
        counter = db.session.get(WorkOrderCounter, org_id, with_for_update=True)
    number = counter.next_number
    counter.next_number = number + 1
    db.session.flush()
    return number


# --------------------------------------------------------------------------- serialization


def allowed_transitions(wo: WorkOrder) -> list[str]:
    return sorted(WO_TRANSITIONS.get(wo.status, set()))


def permissions_for(ctx, wo: WorkOrder) -> dict:
    scope = authz.scope_of(wo)
    return {
        "edit": authz.can(ctx, "work_order.edit", scope=scope),
        "assign": authz.can(ctx, "work_order.assign", scope=scope),
        "start": authz.can(ctx, "work_order.start", scope=scope),
        "complete": authz.can(ctx, "work_order.complete", scope=scope),
        "cancel": authz.can(ctx, "work_order.cancel", scope=scope),
        "comment": authz.can(ctx, "comment.create", scope=scope),
        "upload": authz.can(ctx, "file.upload", scope=scope),
        "log_time": authz.can(ctx, "time_entry.create", scope=scope),
        "log_cost": authz.can(ctx, "cost_entry.create", scope=scope),
        "set_critical": authz.can(ctx, "work_order.set_critical", scope=scope),
    }


def serialize_list_item(wo: WorkOrder, *, comment_counts=None, child_counts=None, memberships=None) -> dict:
    return {
        "id": wo.id,
        "number": wo.number,
        "title": wo.title,
        "status": wo.status,
        "priority": wo.priority,
        "work_type": wo.work_type,
        "is_blocked": bool(wo.is_blocked),
        "is_overdue": wo.is_overdue,
        "due_at": iso(wo.due_at),
        "start_at": iso(wo.start_at),
        "project": ref(wo.project),
        "team": ref(wo.team),
        "location": ref(wo.location),
        "primary_asset": ref(wo.primary_asset),
        "assignees": [user_ref(a.user, (memberships or {}).get(a.user_id)) for a in wo.assignees if a.user],
        "assignee_teams": [ref(a.team) for a in wo.assignees if a.team],
        "categories": [{"id": c.category.id, "name": c.category.name, "color": c.category.color, "icon": c.category.icon} for c in wo.categories],
        "parent_work_order_id": wo.parent_work_order_id,
        "child_count": (child_counts or {}).get(wo.id, 0),
        "comment_count": (comment_counts or {}).get(wo.id, 0),
        "estimated_minutes": wo.estimated_minutes,
        "updated_at": iso(wo.updated_at),
        "last_activity_at": iso(wo.last_activity_at),
        "created_at": iso(wo.created_at),
    }


def serialize_time_entry(t: TimeEntry) -> dict:
    return {"id": t.id, "user": user_ref(t.user), "minutes": t.minutes, "started_at": iso(t.started_at), "ended_at": iso(t.ended_at), "note": t.note, "created_at": iso(t.created_at)}


def serialize_cost_entry(c: CostEntry) -> dict:
    return {"id": c.id, "type": c.type, "amount": float(c.amount), "description": c.description, "created_at": iso(c.created_at)}


def serialize_history(h: WorkOrderStatusHistory) -> dict:
    return {"id": h.id, "from_status": h.from_status, "to_status": h.to_status, "changed_by": user_ref(h.changed_by), "note": h.note, "changed_at": iso(h.changed_at)}


def serialize_detail(ctx, wo: WorkOrder) -> dict:
    payload = serialize_list_item(wo, comment_counts=comments_service.count_for(ctx, "work_order", [wo.id]), child_counts={wo.id: len(wo.children)})
    children = sorted(wo.children, key=lambda c: c.number)
    payload.update(
        {
            "description": wo.description,
            "actual_minutes": wo.actual_minutes,
            "completed_at": iso(wo.completed_at),
            "canceled_at": iso(wo.canceled_at),
            "completion_note": wo.completion_note,
            "cancel_reason": wo.cancel_reason,
            "budget_code": wo.budget_code,
            "recurrence": wo.recurring_rule,
            "parent_completion_policy": wo.parent_completion_policy,
            "parent": ({"id": wo.parent.id, "number": wo.parent.number, "title": wo.parent.title} if wo.parent else None),
            "creator": user_ref(wo.creator),
            "watchers": [user_ref(w.user) for w in wo.watchers],
            "related_assets": [ref(l.asset) for l in wo.asset_links if l.relationship_type != "primary"],
            "children": [serialize_list_item(c) for c in children],
            "child_progress": _child_progress(children),
            "status_history": [serialize_history(h) for h in wo.status_history],
            "time_entries": [serialize_time_entry(t) for t in sorted(wo.time_entries, key=lambda t: t.created_at)],
            "cost_entries": [serialize_cost_entry(c) for c in sorted(wo.cost_entries, key=lambda c: c.created_at)],
            "total_cost": round(float(sum((c.amount for c in wo.cost_entries), 0)), 2),
            "allowed_transitions": allowed_transitions(wo),
            "permissions": permissions_for(ctx, wo),
        }
    )
    return payload


def _child_progress(children) -> dict:
    total = len(children)
    done = sum(1 for c in children if c.status == "DONE")
    return {
        "total": total,
        "done": done,
        "percent": int(round(done / total * 100)) if total else 0,
        "estimated_minutes": sum(c.estimated_minutes or 0 for c in children),
        "actual_minutes": sum(c.actual_minutes or 0 for c in children),
        "by_status": {s: sum(1 for c in children if c.status == s) for s in {c.status for c in children}},
    }


# --------------------------------------------------------------------------- queries


def visible_query(ctx):
    q = scoped(WorkOrder, ctx)
    if authz.visible_work_scope(ctx) == "all":
        return q
    team_ids = ctx.member_team_ids()
    assigned = db.session.query(WorkOrderAssignee.work_order_id).filter(
        or_(WorkOrderAssignee.user_id == ctx.user.id, WorkOrderAssignee.team_id.in_(team_ids or ["-"]))
    )
    return q.filter(or_(WorkOrder.id.in_(assigned), WorkOrder.created_by_user_id == ctx.user.id))


def get_visible_or_404(ctx, work_order_id: str) -> WorkOrder:
    wo = visible_query(ctx).filter(WorkOrder.id == work_order_id).first()
    if wo is None:
        raise NotFound("Work order not found.")
    return wo


SORTS = {
    "priority_desc": lambda: (WorkOrder.priority_rank.desc(), WorkOrder.due_at.asc().nullslast(), WorkOrder.number.desc()),
    "due_asc": lambda: (WorkOrder.due_at.asc().nullslast(), WorkOrder.priority_rank.desc(), WorkOrder.number.desc()),
    "updated_desc": lambda: (WorkOrder.last_activity_at.desc(), WorkOrder.number.desc()),
    "created_desc": lambda: (WorkOrder.created_at.desc(), WorkOrder.number.desc()),
    "number_desc": lambda: (WorkOrder.number.desc(),),
    "number_asc": lambda: (WorkOrder.number.asc(),),
    "project": lambda: (OpsProject.name.asc().nullslast(), WorkOrder.number.desc()),
}


def list_work_orders(ctx, query):
    q = visible_query(ctx)
    tab = query.extra.get("tab", "todo")
    if tab == "done":
        q = q.filter(WorkOrder.status.in_(WO_CLOSED_STATUSES))
    elif tab == "todo":
        q = q.filter(WorkOrder.status.in_(WO_OPEN_STATUSES))
    f = query.filters
    if "status" in f:
        q = q.filter(WorkOrder.status.in_([s.upper() for s in f["status"]]))
    if "priority" in f:
        q = q.filter(WorkOrder.priority.in_([p.upper() for p in f["priority"]]))
    if "work_type" in f:
        q = q.filter(WorkOrder.work_type.in_([t.upper() for t in f["work_type"]]))
    if "project" in f:
        q = q.filter(WorkOrder.project_id.is_(None)) if f["project"] == ["none"] else q.filter(WorkOrder.project_id.in_(f["project"]))
    if "team" in f:
        q = q.filter(or_(WorkOrder.team_id.in_(f["team"]), WorkOrder.id.in_(db.session.query(WorkOrderAssignee.work_order_id).filter(WorkOrderAssignee.team_id.in_(f["team"])))))
    if "location" in f:
        q = q.filter(WorkOrder.location_id.in_(f["location"]))
    if "asset" in f:
        q = q.filter(or_(WorkOrder.primary_asset_id.in_(f["asset"]), WorkOrder.id.in_(db.session.query(WorkOrderAsset.work_order_id).filter(WorkOrderAsset.asset_id.in_(f["asset"])))))
    if "category" in f:
        q = q.filter(WorkOrder.id.in_(db.session.query(WorkOrderCategory.work_order_id).filter(WorkOrderCategory.category_id.in_(f["category"]))))
    if "assignee" in f:
        values = f["assignee"]
        if values == ["unassigned"]:
            q = q.filter(~WorkOrder.id.in_(db.session.query(WorkOrderAssignee.work_order_id)))
        else:
            ids = [ctx.user.id if v == "me" else int(v) for v in values if v == "me" or v.isdigit()]
            team_ids = ctx.member_team_ids() if "me" in values else set()
            cond = WorkOrderAssignee.user_id.in_(ids or [-1])
            if team_ids:
                cond = or_(cond, WorkOrderAssignee.team_id.in_(team_ids))
            q = q.filter(WorkOrder.id.in_(db.session.query(WorkOrderAssignee.work_order_id).filter(cond)))
    if "due" in f:
        now = datetime.utcnow()
        value = f["due"][0]
        if value == "overdue":
            q = q.filter(WorkOrder.due_at < now, WorkOrder.status.in_(WO_OPEN_STATUSES))
        elif value == "today":
            q = q.filter(WorkOrder.due_at >= now.replace(hour=0, minute=0, second=0, microsecond=0), WorkOrder.due_at < now.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1))
        elif value == "week":
            q = q.filter(WorkOrder.due_at >= now, WorkOrder.due_at <= now + timedelta(days=7))
        elif value == "month":
            q = q.filter(WorkOrder.due_at >= now, WorkOrder.due_at <= now + timedelta(days=30))
        elif value == "none":
            q = q.filter(WorkOrder.due_at.is_(None))
        elif value == "has":
            q = q.filter(WorkOrder.due_at.isnot(None))
    if "blocked" in f:
        q = q.filter(WorkOrder.is_blocked.is_(f["blocked"][0] in {"1", "true", "yes"}))
    if "parent" in f:
        q = q.filter(WorkOrder.parent_work_order_id.is_(None)) if f["parent"] == ["none"] else q.filter(WorkOrder.parent_work_order_id.in_(f["parent"]))
    if "created_by" in f:
        q = q.filter(WorkOrder.created_by_user_id.in_([int(v) for v in f["created_by"] if v.isdigit()] or [-1]))
    if query.q:
        like = f"%{query.q.lower()}%"
        cond = or_(func.lower(WorkOrder.title).like(like), func.lower(func.coalesce(WorkOrder.description, "")).like(like))
        if query.q.lstrip("#").isdigit():
            cond = or_(cond, WorkOrder.number == int(query.q.lstrip("#")))
        q = q.filter(cond)
    sort = query.sort if query.sort in SORTS else "due_asc"
    if sort == "project":
        q = q.outerjoin(OpsProject, OpsProject.id == WorkOrder.project_id)
    q = q.order_by(*SORTS[sort]())
    offset = int((query.cursor or {}).get("offset", 0))
    rows = q.offset(offset).limit(query.limit + 1).all()
    has_more = len(rows) > query.limit
    rows = rows[: query.limit]
    ids = [r.id for r in rows]
    comment_counts = comments_service.count_for(ctx, "work_order", ids)
    child_rows = db.session.query(WorkOrder.parent_work_order_id, func.count(WorkOrder.id)).filter(WorkOrder.parent_work_order_id.in_(ids or ["-"])).group_by(WorkOrder.parent_work_order_id).all()
    child_counts = {pid: int(n) for pid, n in child_rows}
    return rows, comment_counts, child_counts, ({"offset": offset + query.limit} if has_more else None)


def counts_for_tabs(ctx, project_id: str | None = None) -> dict:
    q = visible_query(ctx)
    if project_id:
        q = q.filter(WorkOrder.project_id == project_id)
    todo = q.filter(WorkOrder.status.in_(WO_OPEN_STATUSES)).count()
    done = q.filter(WorkOrder.status.in_(WO_CLOSED_STATUSES)).count()
    return {"todo": todo, "done": done}


# --------------------------------------------------------------------------- validation helpers


def _validate_links(ctx, *, project_id=None, team_id=None, location_id=None, primary_asset_id=None, asset_ids=None, category_ids=None, parent_id=None, wo_id=None):
    if project_id:
        get_or_404(OpsProject, project_id, ctx, "Project")
    if team_id:
        get_or_404(Team, team_id, ctx, "Team")
    if location_id:
        get_or_404(Location, location_id, ctx, "Location")
    if primary_asset_id:
        get_or_404(Asset, primary_asset_id, ctx, "Asset")
    for aid in asset_ids or []:
        get_or_404(Asset, aid, ctx, "Asset")
    for cid in category_ids or []:
        get_or_404(Category, cid, ctx, "Category")
    if parent_id:
        parent = get_or_404(WorkOrder, parent_id, ctx, "Parent work order")
        if wo_id and parent.id == wo_id:
            raise Validation("A work order cannot be its own parent.", field="parent_work_order_id")
        if parent.parent_work_order_id and wo_id and parent.parent_work_order_id == wo_id:
            raise Validation("That parent would create a cycle.", field="parent_work_order_id")


def _check_priority(ctx, priority: str, project_id=None, team_id=None):
    if priority == "CRITICAL" and not authz.can_create_for(ctx, "work_order.set_critical", project_id=project_id, team_id=team_id):
        raise Validation("Only leads and safety officers can set CRITICAL priority.", field="priority", code="critical_forbidden")


def _set_categories(wo: WorkOrder, category_ids):
    wanted = set(category_ids or [])
    existing = {c.category_id: c for c in wo.categories}
    for cid, row in existing.items():
        if cid not in wanted:
            db.session.delete(row)
    for cid in wanted:
        if cid not in existing:
            db.session.add(WorkOrderCategory(work_order_id=wo.id, category_id=cid))


def _set_related_assets(wo: WorkOrder, asset_ids):
    wanted = set(asset_ids or [])
    if wo.primary_asset_id:
        wanted.discard(wo.primary_asset_id)
    existing = {l.asset_id: l for l in wo.asset_links if l.relationship_type != "primary"}
    for aid, row in existing.items():
        if aid not in wanted:
            db.session.delete(row)
    for aid in wanted:
        if aid not in existing:
            db.session.add(WorkOrderAsset(work_order_id=wo.id, asset_id=aid, relationship_type="related"))
    primary_link = next((l for l in wo.asset_links if l.relationship_type == "primary"), None)
    if wo.primary_asset_id:
        if primary_link is None:
            db.session.add(WorkOrderAsset(work_order_id=wo.id, asset_id=wo.primary_asset_id, relationship_type="primary"))
        elif primary_link.asset_id != wo.primary_asset_id:
            primary_link.asset_id = wo.primary_asset_id
    elif primary_link is not None:
        db.session.delete(primary_link)


def _apply_assignees(ctx, wo: WorkOrder, user_ids, team_ids) -> tuple[set[int], set[str]]:
    users = resolve_users(ctx.organization.id, user_ids)
    for uid in user_ids or []:
        if int(uid) not in users:
            raise Validation(f"User {uid} is not an active member.", field="assignee_user_ids")
    for tid in team_ids or []:
        get_or_404(Team, tid, ctx, "Team")
    wanted_users = {int(u) for u in (user_ids or [])}
    wanted_teams = set(team_ids or [])
    existing_users = {a.user_id: a for a in wo.assignees if a.user_id}
    existing_teams = {a.team_id: a for a in wo.assignees if a.team_id}
    new_users = wanted_users - set(existing_users)
    new_teams = wanted_teams - set(existing_teams)
    for uid, row in existing_users.items():
        if uid not in wanted_users:
            db.session.delete(row)
    for tid, row in existing_teams.items():
        if tid not in wanted_teams:
            db.session.delete(row)
    for uid in new_users:
        db.session.add(WorkOrderAssignee(work_order_id=wo.id, user_id=uid, assigned_by_user_id=ctx.user.id))
    for tid in new_teams:
        db.session.add(WorkOrderAssignee(work_order_id=wo.id, team_id=tid, assigned_by_user_id=ctx.user.id))
    db.session.flush()
    return new_users, new_teams


def _set_watchers(ctx, wo: WorkOrder, user_ids):
    users = resolve_users(ctx.organization.id, user_ids)
    wanted = {int(u) for u in (user_ids or []) if int(u) in users}
    existing = {w.user_id: w for w in wo.watchers}
    for uid, row in existing.items():
        if uid not in wanted:
            db.session.delete(row)
    for uid in wanted:
        if uid not in existing:
            db.session.add(WorkOrderWatcher(work_order_id=wo.id, user_id=uid))
    db.session.flush()


def _notify_assignment(ctx, wo: WorkOrder, new_users: set[int], new_teams: set[str]):
    recipients = set(new_users)
    if new_teams:
        for (uid,) in db.session.query(TeamMember.user_id).filter(TeamMember.team_id.in_(new_teams), TeamMember.is_lead.is_(True)):
            recipients.add(uid)
    if recipients:
        notifications.notify(
            ctx.organization.id, recipients, type="work_order.assigned", title=f"Assigned: #{wo.number} {wo.title}",
            body=f"{ctx.user.name} assigned you to work order #{wo.number}.", entity_type="work_order", entity_id=wo.id, exclude_user_id=ctx.user.id,
        )


def _interested_user_ids(wo: WorkOrder) -> set[int]:
    ids = set(wo.assignee_user_ids) | {w.user_id for w in wo.watchers}
    if wo.created_by_user_id:
        ids.add(wo.created_by_user_id)
    return ids


def _record_status(wo: WorkOrder, from_status, to_status, actor, note=None):
    db.session.add(WorkOrderStatusHistory(work_order_id=wo.id, from_status=from_status, to_status=to_status, changed_by_user_id=actor.id if actor else None, note=note))
    wo.last_activity_at = datetime.utcnow()


# --------------------------------------------------------------------------- create / update


def create(ctx, data, *, parent: WorkOrder | None = None) -> WorkOrder:
    project_id = data.project_id or (parent.project_id if parent else None)
    team_id = data.team_id or (parent.team_id if parent else None)
    authz.require_create(ctx, "work_order.create", project_id=project_id, team_id=team_id)
    _validate_links(
        ctx, project_id=project_id, team_id=team_id, location_id=data.location_id, primary_asset_id=data.primary_asset_id,
        asset_ids=data.asset_ids, category_ids=data.category_ids, parent_id=(parent.id if parent else data.parent_work_order_id),
    )
    _check_priority(ctx, data.priority, project_id=project_id, team_id=team_id)
    if data.due_at and data.start_at and data.due_at < data.start_at:
        raise Validation("Due date cannot be before the start date.", field="due_at")
    if (data.assignee_user_ids or data.assignee_team_ids) and not authz.can_create_for(ctx, "work_order.assign", project_id=project_id, team_id=team_id):
        # members may only assign themselves
        if set(data.assignee_user_ids) - {ctx.user.id} or data.assignee_team_ids:
            raise Validation("You can only assign yourself.", field="assignee_user_ids", code="assign_forbidden")

    for attempt in range(3):
        try:
            with db.session.begin_nested():
                wo = WorkOrder(
                    organization_id=ctx.organization.id,
                    number=allocate_number(ctx.organization.id),
                    title=data.title,
                    description=data.description,
                    status=data.status,
                    priority=data.priority,
                    priority_rank=WO_PRIORITY_RANK[data.priority],
                    work_type=data.work_type,
                    project_id=project_id,
                    team_id=team_id,
                    location_id=data.location_id or (parent.location_id if parent else None),
                    primary_asset_id=data.primary_asset_id or (parent.primary_asset_id if parent else None),
                    parent_work_order_id=parent.id if parent else data.parent_work_order_id,
                    start_at=data.start_at,
                    due_at=data.due_at,
                    estimated_minutes=data.estimated_minutes,
                    budget_code=data.budget_code,
                    recurring_rule_json=json.dumps(data.recurrence.model_dump()) if data.recurrence else None,
                    parent_completion_policy=data.parent_completion_policy,
                    created_by_user_id=ctx.user.id,
                    updated_by_user_id=ctx.user.id,
                )
                db.session.add(wo)
                db.session.flush()
            break
        except IntegrityError:
            if attempt == 2:
                raise
            continue
    _set_categories(wo, data.category_ids)
    _set_related_assets(wo, data.asset_ids)
    new_users, new_teams = _apply_assignees(ctx, wo, data.assignee_user_ids, data.assignee_team_ids)
    _set_watchers(ctx, wo, set(data.watcher_user_ids) | {ctx.user.id})
    _record_status(wo, None, wo.status, ctx.user)
    _notify_assignment(ctx, wo, new_users, new_teams)
    audit.record_event("work_order.created", "work_order", wo.id, organization_id=ctx.organization.id, actor=ctx.user, after=audit.snapshot(wo, WO_FIELDS), metadata={"number": wo.number, "assignees": sorted(new_users), "teams": sorted(new_teams)})
    db.session.commit()
    events.emit(EVENT_CREATED, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id)
    return wo


def update(ctx, wo: WorkOrder, data, fields_set: set[str]) -> WorkOrder:
    authz.require(ctx, "work_order.edit", record=wo)
    if wo.status in WO_CLOSED_STATUSES:
        raise Conflict("Closed work orders cannot be edited. Duplicate it instead.", code="closed")
    project_id = data.project_id if "project_id" in fields_set else wo.project_id
    team_id = data.team_id if "team_id" in fields_set else wo.team_id
    _validate_links(
        ctx,
        project_id=project_id if "project_id" in fields_set else None,
        team_id=team_id if "team_id" in fields_set else None,
        location_id=data.location_id if "location_id" in fields_set else None,
        primary_asset_id=data.primary_asset_id if "primary_asset_id" in fields_set else None,
        asset_ids=data.asset_ids if "asset_ids" in fields_set else None,
        category_ids=data.category_ids if "category_ids" in fields_set else None,
        wo_id=wo.id,
    )
    if "priority" in fields_set and data.priority:
        _check_priority(ctx, data.priority, project_id=project_id, team_id=team_id)
    before = audit.snapshot(wo, WO_FIELDS)
    simple = {"title", "description", "project_id", "team_id", "location_id", "primary_asset_id", "start_at", "due_at", "estimated_minutes", "budget_code", "is_blocked", "parent_completion_policy", "work_type"}
    for key in fields_set & simple:
        setattr(wo, key, getattr(data, key))
    if "priority" in fields_set and data.priority:
        wo.priority = data.priority
        wo.priority_rank = WO_PRIORITY_RANK[data.priority]
    if "recurrence" in fields_set:
        wo.recurring_rule_json = json.dumps(data.recurrence.model_dump()) if data.recurrence else None
    if wo.due_at and wo.start_at and wo.due_at < wo.start_at:
        raise Validation("Due date cannot be before the start date.", field="due_at")
    if "category_ids" in fields_set and data.category_ids is not None:
        _set_categories(wo, data.category_ids)
    if "asset_ids" in fields_set or "primary_asset_id" in fields_set:
        _set_related_assets(wo, data.asset_ids if "asset_ids" in fields_set and data.asset_ids is not None else [l.asset_id for l in wo.asset_links if l.relationship_type != "primary"])
    touch(wo, ctx.user)
    wo.last_activity_at = datetime.utcnow()
    after = audit.snapshot(wo, WO_FIELDS)
    audit.record_event("work_order.updated", "work_order", wo.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"number": wo.number, "changed": audit.diff(before, after)})
    db.session.commit()
    events.emit(EVENT_UPDATED, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id)
    return wo


def set_assignees(ctx, wo: WorkOrder, user_ids, team_ids) -> WorkOrder:
    authz.require(ctx, "work_order.assign", record=wo)
    before = {"users": sorted(wo.assignee_user_ids), "teams": sorted(wo.assignee_team_ids)}
    new_users, new_teams = _apply_assignees(ctx, wo, user_ids, team_ids)
    db.session.refresh(wo)
    after = {"users": sorted(wo.assignee_user_ids), "teams": sorted(wo.assignee_team_ids)}
    wo.last_activity_at = datetime.utcnow()
    touch(wo, ctx.user)
    _notify_assignment(ctx, wo, new_users, new_teams)
    audit.record_event("work_order.assigned", "work_order", wo.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"number": wo.number})
    db.session.commit()
    events.emit(EVENT_ASSIGNED, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id, new_users=sorted(new_users))
    return wo


def set_watchers(ctx, wo: WorkOrder, user_ids) -> WorkOrder:
    authz.require(ctx, "work_order.edit", record=wo)
    _set_watchers(ctx, wo, user_ids)
    db.session.commit()
    return wo


# --------------------------------------------------------------------------- lifecycle


def _transition(ctx, wo: WorkOrder, to_status: str, permission: str, note: str | None = None):
    authz.require(ctx, permission, record=wo)
    if to_status not in WO_TRANSITIONS.get(wo.status, set()):
        raise Conflict(f"Cannot move a work order from {wo.status} to {to_status}.", code="invalid_transition", from_status=wo.status, to_status=to_status)
    from_status = wo.status
    wo.status = to_status
    touch(wo, ctx.user)
    _record_status(wo, from_status, to_status, ctx.user, note)
    audit.record_event("work_order.status_changed", "work_order", wo.id, organization_id=ctx.organization.id, actor=ctx.user, before={"status": from_status}, after={"status": to_status}, metadata={"number": wo.number, "note": note})
    notifications.notify(
        ctx.organization.id, _interested_user_ids(wo), type="work_order.status", title=f"#{wo.number} is now {to_status.replace('_', ' ').title()}",
        body=note, entity_type="work_order", entity_id=wo.id, exclude_user_id=ctx.user.id,
    )
    return from_status


def start(ctx, wo: WorkOrder, note=None) -> WorkOrder:
    if wo.status == "DRAFT":
        _transition(ctx, wo, "OPEN", "work_order.edit", note)
    from_status = _transition(ctx, wo, "IN_PROGRESS", "work_order.start", note)
    if wo.start_at is None:
        wo.start_at = datetime.utcnow()
    db.session.commit()
    events.emit(EVENT_STATUS, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id, from_status=from_status, to_status="IN_PROGRESS")
    return wo


def hold(ctx, wo: WorkOrder, note=None) -> WorkOrder:
    from_status = _transition(ctx, wo, "ON_HOLD", "work_order.start", note)
    db.session.commit()
    events.emit(EVENT_STATUS, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id, from_status=from_status, to_status="ON_HOLD")
    return wo


def resume(ctx, wo: WorkOrder, note=None) -> WorkOrder:
    from_status = _transition(ctx, wo, "IN_PROGRESS", "work_order.start", note)
    db.session.commit()
    events.emit(EVENT_STATUS, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id, from_status=from_status, to_status="IN_PROGRESS")
    return wo


def open_draft(ctx, wo: WorkOrder) -> WorkOrder:
    from_status = _transition(ctx, wo, "OPEN", "work_order.edit")
    db.session.commit()
    events.emit(EVENT_STATUS, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id, from_status=from_status, to_status="OPEN")
    return wo


def cancel(ctx, wo: WorkOrder, note=None) -> WorkOrder:
    from_status = _transition(ctx, wo, "CANCELED", "work_order.cancel", note)
    wo.canceled_at = datetime.utcnow()
    wo.cancel_reason = (note or "")[:500] or None
    db.session.commit()
    events.emit(EVENT_STATUS, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id, from_status=from_status, to_status="CANCELED")
    return wo


def recompute_actual_minutes(wo: WorkOrder):
    wo.actual_minutes = int(db.session.query(func.coalesce(func.sum(TimeEntry.minutes), 0)).filter(TimeEntry.work_order_id == wo.id).scalar() or 0)


def _next_occurrence(wo: WorkOrder) -> WorkOrder | None:
    rule = wo.recurring_rule
    if not rule:
        return None
    interval = int(rule.get("interval") or 1)
    freq = rule.get("frequency")
    step = {"daily": timedelta(days=interval), "weekly": timedelta(weeks=interval), "monthly": timedelta(days=30 * interval)}.get(freq)
    if step is None:
        return None
    anchor = wo.due_at if rule.get("mode") == "fixed" and wo.due_at else (wo.completed_at or datetime.utcnow())
    next_due = anchor + step
    key = f"recur:{wo.id}:{next_due.date().isoformat()}"
    if WorkOrder.query.filter_by(generation_key=key).first():
        return None
    nxt = WorkOrder(
        organization_id=wo.organization_id,
        number=allocate_number(wo.organization_id),
        title=wo.title,
        description=wo.description,
        status="OPEN",
        priority=wo.priority,
        priority_rank=wo.priority_rank,
        work_type=wo.work_type,
        project_id=wo.project_id,
        team_id=wo.team_id,
        location_id=wo.location_id,
        primary_asset_id=wo.primary_asset_id,
        due_at=next_due,
        start_at=None,
        estimated_minutes=wo.estimated_minutes,
        budget_code=wo.budget_code,
        recurring_rule_json=wo.recurring_rule_json,
        generation_key=key,
        maintenance_plan_id=wo.maintenance_plan_id,
        created_by_user_id=wo.created_by_user_id,
        updated_by_user_id=wo.created_by_user_id,
    )
    db.session.add(nxt)
    db.session.flush()
    for c in wo.categories:
        db.session.add(WorkOrderCategory(work_order_id=nxt.id, category_id=c.category_id))
    for a in wo.assignees:
        db.session.add(WorkOrderAssignee(work_order_id=nxt.id, user_id=a.user_id, team_id=a.team_id, assigned_by_user_id=a.assigned_by_user_id))
    for l in wo.asset_links:
        db.session.add(WorkOrderAsset(work_order_id=nxt.id, asset_id=l.asset_id, relationship_type=l.relationship_type))
    _record_status(nxt, None, "OPEN", None, f"Generated from #{wo.number}")
    return nxt


def complete(ctx, wo: WorkOrder, data) -> tuple[WorkOrder, WorkOrder | None, WorkOrder | None]:
    """Returns ``(work_order, follow_up, next_occurrence)``."""
    authz.require(ctx, "work_order.complete", record=wo)
    if wo.status == "DRAFT":
        raise Conflict("Open the draft before completing it.", code="invalid_transition")
    if wo.status == "ON_HOLD":
        _transition(ctx, wo, "IN_PROGRESS", "work_order.start", "Resumed to complete")
    open_children = [c for c in wo.children if c.status in WO_OPEN_STATUSES]
    if open_children:
        raise Conflict(f"{len(open_children)} sub-work order(s) are still open.", code="children_open", open_children=[c.number for c in open_children])
    from_status = _transition(ctx, wo, "DONE", "work_order.complete", data.completion_note)
    now = datetime.utcnow()
    wo.completed_at = now
    wo.completion_note = data.completion_note
    if data.time_minutes:
        db.session.add(TimeEntry(organization_id=ctx.organization.id, work_order_id=wo.id, user_id=ctx.user.id, minutes=data.time_minutes, note="Logged at completion", created_by_user_id=ctx.user.id, updated_by_user_id=ctx.user.id))
    for cost in data.costs:
        db.session.add(CostEntry(organization_id=ctx.organization.id, work_order_id=wo.id, type=cost.type, amount=cost.amount, description=cost.description, created_by_user_id=ctx.user.id, updated_by_user_id=ctx.user.id))
    db.session.flush()
    recompute_actual_minutes(wo)
    if data.asset_status and wo.primary_asset and data.asset_status != wo.primary_asset.status:
        from asme.ops.schemas import AssetStatusChange
        from asme.ops.services import assets as assets_service

        if authz.can(ctx, "asset.manage", record=wo.primary_asset) or authz.can(ctx, "work_order.complete", record=wo):
            assets_service.change_status(ctx, wo.primary_asset, AssetStatusChange(status=data.asset_status, note=f"Updated on completion of #{wo.number}"))
    follow_up = None
    if data.follow_up_title:
        from asme.ops.schemas import WorkOrderCreate

        follow_up = create(
            ctx,
            WorkOrderCreate(
                title=data.follow_up_title, description=f"Follow-up from #{wo.number}: {wo.title}", project_id=wo.project_id, team_id=wo.team_id,
                location_id=wo.location_id, primary_asset_id=wo.primary_asset_id, work_type=wo.work_type, priority=wo.priority if wo.priority != "CRITICAL" else "HIGH",
                category_ids=[c.category_id for c in wo.categories],
            ),
        )
        wo = db.session.get(WorkOrder, wo.id)
    next_occurrence = _next_occurrence(wo)
    # parent auto-completion policy
    parent = wo.parent
    if parent and parent.parent_completion_policy == "auto" and parent.status in WO_OPEN_STATUSES:
        siblings = [c for c in parent.children if c.id != wo.id]
        if all(c.status in WO_CLOSED_STATUSES for c in siblings):
            p_from = parent.status
            parent.status = "DONE"
            parent.completed_at = now
            _record_status(parent, p_from, "DONE", ctx.user, "Auto-completed: all sub-work orders done")
            audit.record_event("work_order.status_changed", "work_order", parent.id, organization_id=ctx.organization.id, actor=ctx.user, before={"status": p_from}, after={"status": "DONE"}, metadata={"number": parent.number, "auto": True})
    db.session.commit()
    events.emit(EVENT_STATUS, work_order_id=wo.id, user_id=ctx.user.id, organization_id=ctx.organization.id, from_status=from_status, to_status="DONE")
    return wo, follow_up, next_occurrence


def duplicate(ctx, wo: WorkOrder) -> WorkOrder:
    from asme.ops.schemas import WorkOrderCreate

    data = WorkOrderCreate(
        title=wo.title, description=wo.description, project_id=wo.project_id, location_id=wo.location_id, primary_asset_id=wo.primary_asset_id,
        asset_ids=[l.asset_id for l in wo.asset_links if l.relationship_type != "primary"], assignee_user_ids=sorted(wo.assignee_user_ids),
        assignee_team_ids=sorted(wo.assignee_team_ids), team_id=wo.team_id, estimated_minutes=wo.estimated_minutes, work_type=wo.work_type,
        priority=wo.priority if authz.can(ctx, "work_order.set_critical", record=wo) or wo.priority != "CRITICAL" else "HIGH",
        category_ids=[c.category_id for c in wo.categories], budget_code=wo.budget_code, parent_work_order_id=wo.parent_work_order_id,
        parent_completion_policy=wo.parent_completion_policy,
    )
    if not authz.can(ctx, "work_order.assign", record=wo):
        data.assignee_user_ids, data.assignee_team_ids = [], []
    copy = create(ctx, data)
    audit.record_event("work_order.duplicated", "work_order", copy.id, organization_id=ctx.organization.id, actor=ctx.user, metadata={"from_number": wo.number, "number": copy.number})
    db.session.commit()
    return copy


def create_sub(ctx, parent: WorkOrder, data) -> WorkOrder:
    from asme.ops.schemas import WorkOrderCreate

    if parent.status in WO_CLOSED_STATUSES:
        raise Conflict("Cannot add sub-work orders to a closed work order.", code="closed")
    return create(
        ctx,
        WorkOrderCreate(
            title=data.title, description=data.description, assignee_user_ids=data.assignee_user_ids, due_at=data.due_at, priority=data.priority,
            estimated_minutes=data.estimated_minutes, work_type=parent.work_type, category_ids=[c.category_id for c in parent.categories],
        ),
        parent=parent,
    )


# --------------------------------------------------------------------------- time & cost


def add_time_entry(ctx, wo: WorkOrder, data) -> TimeEntry:
    authz.require(ctx, "time_entry.create", record=wo)
    user_id = ctx.user.id
    if data.user_id and data.user_id != ctx.user.id:
        if not authz.can(ctx, "work_order.assign", record=wo):
            raise Validation("You can only log your own time.", field="user_id")
        if data.user_id not in resolve_users(ctx.organization.id, [data.user_id]):
            raise Validation("Unknown member.", field="user_id")
        user_id = data.user_id
    if data.started_at and data.ended_at and data.ended_at < data.started_at:
        raise Validation("End must be after start.", field="ended_at")
    entry = TimeEntry(organization_id=ctx.organization.id, work_order_id=wo.id, user_id=user_id, minutes=data.minutes, started_at=data.started_at, ended_at=data.ended_at, note=data.note, created_by_user_id=ctx.user.id, updated_by_user_id=ctx.user.id)
    db.session.add(entry)
    db.session.flush()
    recompute_actual_minutes(wo)
    wo.last_activity_at = datetime.utcnow()
    audit.record_event("work_order.time_logged", "work_order", wo.id, organization_id=ctx.organization.id, actor=ctx.user, after={"minutes": data.minutes, "user_id": user_id}, metadata={"number": wo.number})
    db.session.commit()
    return entry


def add_cost_entry(ctx, wo: WorkOrder, data) -> CostEntry:
    authz.require(ctx, "cost_entry.create", record=wo)
    entry = CostEntry(organization_id=ctx.organization.id, work_order_id=wo.id, type=data.type, amount=data.amount, description=data.description, created_by_user_id=ctx.user.id, updated_by_user_id=ctx.user.id)
    db.session.add(entry)
    db.session.flush()
    wo.last_activity_at = datetime.utcnow()
    audit.record_event("work_order.cost_logged", "work_order", wo.id, organization_id=ctx.organization.id, actor=ctx.user, after={"type": data.type, "amount": float(data.amount)}, metadata={"number": wo.number})
    db.session.commit()
    return entry


def add_comment(ctx, wo: WorkOrder, body: str, parent_comment_id=None):
    authz.require(ctx, "comment.create", record=wo)
    comment = comments_service.create(ctx, "work_order", wo.id, body, parent_comment_id)
    wo.last_activity_at = datetime.utcnow()
    recipients = _interested_user_ids(wo) | comments_service.mentioned_user_ids(body)
    notifications.notify(ctx.organization.id, recipients, type="work_order.comment", title=f"New comment on #{wo.number}", body=body[:200], entity_type="work_order", entity_id=wo.id, exclude_user_id=ctx.user.id)
    db.session.commit()
    return comment
