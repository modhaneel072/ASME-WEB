"""Operations dashboard aggregates. Formulas are documented in docs/reporting-metrics.md
and tested in tests/ops/test_dashboard.py - keep the three in step."""

from __future__ import annotations

from datetime import datetime, timedelta

from sqlalchemy import func

from asme.extensions import db
from asme.ops.models import WO_OPEN_STATUSES, WO_PRIORITIES, WO_WORK_TYPES, CostEntry, Team, TimeEntry, WorkOrder
from asme.models import User


def _range(range_key: str) -> tuple[datetime, datetime, int]:
    days = {"7d": 7, "30d": 30, "90d": 90, "365d": 365}.get(range_key, 30)
    end = datetime.utcnow()
    start = end - timedelta(days=days)
    return start, end, days


def operations(ctx, range_key: str = "30d", project_id: str | None = None) -> dict:
    org_id = ctx.organization.id
    start, end, days = _range(range_key)
    base = WorkOrder.query.filter(WorkOrder.organization_id == org_id)
    if project_id:
        base = base.filter(WorkOrder.project_id == project_id)

    def _q():
        q = WorkOrder.query.filter(WorkOrder.organization_id == org_id)
        return q.filter(WorkOrder.project_id == project_id) if project_id else q

    by_status = {s: 0 for s in ("DRAFT", "OPEN", "IN_PROGRESS", "ON_HOLD", "DONE", "CANCELED", "SKIPPED")}
    for status, n in db.session.query(WorkOrder.status, func.count(WorkOrder.id)).filter(WorkOrder.organization_id == org_id, *( [WorkOrder.project_id == project_id] if project_id else [] )).group_by(WorkOrder.status).all():
        by_status[status] = int(n)
    open_total = sum(by_status[s] for s in WO_OPEN_STATUSES)
    overdue = _q().filter(WorkOrder.status.in_(WO_OPEN_STATUSES), WorkOrder.due_at < end).count()
    blocked = _q().filter(WorkOrder.status.in_(WO_OPEN_STATUSES), WorkOrder.is_blocked.is_(True)).count()
    due_soon = _q().filter(WorkOrder.status.in_(WO_OPEN_STATUSES), WorkOrder.due_at >= end, WorkOrder.due_at <= end + timedelta(days=7)).count()

    created_in_range = _q().filter(WorkOrder.created_at >= start, WorkOrder.created_at <= end)
    completed_in_range = _q().filter(WorkOrder.status == "DONE", WorkOrder.completed_at >= start, WorkOrder.completed_at <= end)
    created_count = created_in_range.count()
    completed_count = completed_in_range.count()

    # weekly series: created vs completed
    buckets = max(1, days // 7)
    series = []
    for i in range(buckets):
        b_start = start + timedelta(days=7 * i)
        b_end = end if i == buckets - 1 else b_start + timedelta(days=7)
        series.append(
            {
                "start": b_start.date().isoformat(),
                "created": _q().filter(WorkOrder.created_at >= b_start, WorkOrder.created_at < b_end).count(),
                "completed": _q().filter(WorkOrder.status == "DONE", WorkOrder.completed_at >= b_start, WorkOrder.completed_at < b_end).count(),
            }
        )

    completed_rows = completed_in_range.all()
    on_time = sum(1 for wo in completed_rows if not wo.due_at or (wo.completed_at and wo.completed_at <= wo.due_at))
    on_time_rate = round(on_time / completed_count * 100) if completed_count else None
    durations = [(wo.completed_at - wo.created_at).total_seconds() / 3600 for wo in completed_rows if wo.completed_at and wo.created_at]
    avg_completion_hours = round(sum(durations) / len(durations), 1) if durations else None

    by_priority = {p: 0 for p in WO_PRIORITIES}
    for p, n in db.session.query(WorkOrder.priority, func.count(WorkOrder.id)).filter(WorkOrder.organization_id == org_id, WorkOrder.status.in_(WO_OPEN_STATUSES), *( [WorkOrder.project_id == project_id] if project_id else [] )).group_by(WorkOrder.priority).all():
        by_priority[p] = int(n)
    by_type = {t: 0 for t in WO_WORK_TYPES}
    for t, n in db.session.query(WorkOrder.work_type, func.count(WorkOrder.id)).filter(WorkOrder.organization_id == org_id, WorkOrder.created_at >= start, *( [WorkOrder.project_id == project_id] if project_id else [] )).group_by(WorkOrder.work_type).all():
        by_type[t] = int(n)
    repeating = _q().filter(WorkOrder.created_at >= start, WorkOrder.recurring_rule_json.isnot(None)).count()
    non_repeating = created_count - repeating

    workload = [
        {"team_id": tid, "team": name, "open": int(n)}
        for tid, name, n in db.session.query(Team.id, Team.name, func.count(WorkOrder.id))
        .join(WorkOrder, WorkOrder.team_id == Team.id)
        .filter(WorkOrder.organization_id == org_id, WorkOrder.status.in_(WO_OPEN_STATUSES), *( [WorkOrder.project_id == project_id] if project_id else [] ))
        .group_by(Team.id, Team.name)
        .order_by(func.count(WorkOrder.id).desc())
        .limit(10)
        .all()
    ]
    from asme.ops.models import WorkOrderAssignee

    user_load = [
        {"user_id": uid, "name": name, "open": int(n)}
        for uid, name, n in db.session.query(User.id, User.name, func.count(WorkOrder.id))
        .join(WorkOrderAssignee, WorkOrderAssignee.user_id == User.id)
        .join(WorkOrder, WorkOrder.id == WorkOrderAssignee.work_order_id)
        .filter(WorkOrder.organization_id == org_id, WorkOrder.status.in_(WO_OPEN_STATUSES), *( [WorkOrder.project_id == project_id] if project_id else [] ))
        .group_by(User.id, User.name)
        .order_by(func.count(WorkOrder.id).desc())
        .limit(10)
        .all()
    ]
    minutes = db.session.query(func.coalesce(func.sum(TimeEntry.minutes), 0)).join(WorkOrder, WorkOrder.id == TimeEntry.work_order_id).filter(WorkOrder.organization_id == org_id, TimeEntry.created_at >= start, *( [WorkOrder.project_id == project_id] if project_id else [] )).scalar() or 0
    cost_rows = db.session.query(CostEntry.type, func.coalesce(func.sum(CostEntry.amount), 0)).join(WorkOrder, WorkOrder.id == CostEntry.work_order_id).filter(WorkOrder.organization_id == org_id, CostEntry.created_at >= start, *( [WorkOrder.project_id == project_id] if project_id else [] )).group_by(CostEntry.type).all()
    costs = {t: float(v) for t, v in cost_rows}

    return {
        "range": {"key": range_key, "start": start.isoformat(), "end": end.isoformat(), "days": days},
        "project_id": project_id,
        "totals": {"open": open_total, "overdue": overdue, "blocked": blocked, "due_soon": due_soon, "created": created_count, "completed": completed_count},
        "by_status": by_status,
        "by_priority": by_priority,
        "by_work_type": by_type,
        "repeating": {"repeating": repeating, "non_repeating": max(non_repeating, 0)},
        "created_vs_completed": series,
        "on_time_completion_rate": on_time_rate,
        "average_completion_hours": avg_completion_hours,
        "workload_by_team": workload,
        "workload_by_user": user_load,
        "hours_logged": round(int(minutes) / 60, 1),
        "costs": {"parts": costs.get("part", 0.0), "labor": costs.get("labor", 0.0), "vendor": costs.get("vendor", 0.0), "other": costs.get("other", 0.0), "total": round(sum(costs.values()), 2)},
        "generated_at": end.isoformat(),
    }
