"""Operational projects: list/detail, membership, milestones, health, activity."""

from __future__ import annotations

import re
from datetime import date, datetime

from sqlalchemy import func

from asme.extensions import db
from asme.models import Project as PublicProject
from asme.ops import audit
from asme.ops.models import (
    WO_CLOSED_STATUSES,
    WO_OPEN_STATUSES,
    AuditEvent,
    Milestone,
    OpsProject,
    OpsProjectMember,
    Team,
    WorkOrder,
)
from asme.ops.services.common import apply_patch, iso, ref, resolve_users, touch, user_ref
from asme.ops.tenancy import get_or_404, scoped
from asme.services.errors import Conflict, NotFound, Validation

PROJECT_FIELDS = (
    "name", "code", "description", "status", "risk_level", "lead_user_id", "faculty_advisor_user_id", "competition",
    "academic_year", "start_date", "target_date", "budget_amount", "budget_code", "repository_url", "cad_url",
    "requirements_url", "visibility", "public_project_id",
)
MILESTONE_FIELDS = ("name", "description", "due_date", "status", "owner_user_id", "weight", "order")


def next_code(ctx, name: str) -> str:
    base = "".join(w[0] for w in re.findall(r"[A-Za-z0-9]+", name or ""))[:6].upper() or "PRJ"
    candidate = base
    n = 2
    while scoped(OpsProject, ctx).filter(func.lower(OpsProject.code) == candidate.lower()).first():
        candidate = f"{base}{n}"
        n += 1
    return candidate


# --------------------------------------------------------------------------- metrics


def _wo_counts(org_id: str, project_ids: list[str]) -> dict[str, dict]:
    if not project_ids:
        return {}
    now = datetime.utcnow()
    rows = db.session.query(WorkOrder.project_id, WorkOrder.status, func.count(WorkOrder.id)).filter(
        WorkOrder.organization_id == org_id, WorkOrder.project_id.in_(project_ids)
    ).group_by(WorkOrder.project_id, WorkOrder.status).all()
    overdue = db.session.query(WorkOrder.project_id, func.count(WorkOrder.id)).filter(
        WorkOrder.organization_id == org_id, WorkOrder.project_id.in_(project_ids), WorkOrder.status.in_(WO_OPEN_STATUSES), WorkOrder.due_at < now
    ).group_by(WorkOrder.project_id).all()
    out: dict[str, dict] = {pid: {"open": 0, "done": 0, "total": 0, "overdue": 0, "blocked": 0} for pid in project_ids}
    for pid, status, n in rows:
        bucket = out.setdefault(pid, {"open": 0, "done": 0, "total": 0, "overdue": 0, "blocked": 0})
        n = int(n)
        if status in WO_OPEN_STATUSES:
            bucket["open"] += n
        if status == "DONE":
            bucket["done"] += n
        if status not in {"CANCELED", "SKIPPED", "DRAFT"}:
            bucket["total"] += n
    for pid, n in overdue:
        out[pid]["overdue"] = int(n)
    blocked = db.session.query(WorkOrder.project_id, func.count(WorkOrder.id)).filter(
        WorkOrder.organization_id == org_id, WorkOrder.project_id.in_(project_ids), WorkOrder.status.in_(WO_OPEN_STATUSES), WorkOrder.is_blocked.is_(True)
    ).group_by(WorkOrder.project_id).all()
    for pid, n in blocked:
        out[pid]["blocked"] = int(n)
    return out


def completion_percent(project: OpsProject, counts: dict | None) -> int:
    """Weighted milestone completion; falls back to done/total work orders when no milestones."""
    milestones = list(project.milestones)
    if milestones:
        total = sum(max(m.weight or 1, 1) for m in milestones)
        done = sum(max(m.weight or 1, 1) for m in milestones if m.status == "complete")
        return int(round(done / total * 100)) if total else 0
    if counts and counts.get("total"):
        return int(round(counts["done"] / counts["total"] * 100))
    return 0


def next_milestone(project: OpsProject) -> dict | None:
    today = date.today()
    upcoming = [m for m in project.milestones if m.status not in {"complete"} and m.due_date]
    if not upcoming:
        return None
    upcoming.sort(key=lambda m: m.due_date)
    m = next((x for x in upcoming if x.due_date >= today), upcoming[0])
    return {"id": m.id, "name": m.name, "due_date": m.due_date.isoformat(), "status": m.status, "days_left": (m.due_date - today).days}


def serialize_project_list_item(project: OpsProject, counts: dict, team_counts: dict, memberships_by_user=None) -> dict:
    c = counts.get(project.id, {"open": 0, "done": 0, "total": 0, "overdue": 0, "blocked": 0})
    return {
        "id": project.id,
        "name": project.name,
        "code": project.code,
        "status": project.status,
        "risk_level": project.risk_level,
        "lead": user_ref(project.lead, (memberships_by_user or {}).get(project.lead_user_id)),
        "completion_percent": completion_percent(project, c),
        "open_work_orders": c["open"],
        "overdue_work_orders": c["overdue"],
        "blocked_work_orders": c["blocked"],
        "next_milestone": next_milestone(project),
        "target_date": iso(project.target_date),
        "academic_year": project.academic_year,
        "competition": project.competition,
        "team_count": team_counts.get(project.id, 0),
        "visibility": project.visibility,
        "archived_at": iso(project.archived_at),
        "updated_at": iso(project.updated_at),
    }


def serialize_project(project: OpsProject, ctx, memberships_by_user=None) -> dict:
    counts = _wo_counts(ctx.organization.id, [project.id])
    team_counts = {project.id: scoped(Team, ctx).filter(Team.project_id == project.id, Team.archived_at.is_(None)).count()}
    payload = serialize_project_list_item(project, counts, team_counts, memberships_by_user)
    cost_total = db.session.query(func.coalesce(func.sum(CostEntryAmount.amount), 0)).select_from(CostEntryAmount).join(WorkOrder, WorkOrder.id == CostEntryAmount.work_order_id).filter(WorkOrder.project_id == project.id).scalar() or 0
    hours = db.session.query(func.coalesce(func.sum(WorkOrder.actual_minutes), 0)).filter(WorkOrder.project_id == project.id).scalar() or 0
    payload.update(
        {
            "description": project.description,
            "faculty_advisor": user_ref(project.faculty_advisor),
            "start_date": iso(project.start_date),
            "budget_amount": float(project.budget_amount) if project.budget_amount is not None else None,
            "budget_code": project.budget_code,
            "budget_used": float(cost_total),
            "hours_logged": round(int(hours) / 60, 1),
            "repository_url": project.repository_url,
            "cad_url": project.cad_url,
            "requirements_url": project.requirements_url,
            "public_project_id": project.public_project_id,
            "public_slug": project.public_project.slug if project.public_project else None,
            "members": [
                {**user_ref(m.user, (memberships_by_user or {}).get(m.user_id)), "project_role": m.project_role, "team_id": m.team_id, "team": ref(m.team)}
                for m in sorted(project.members, key=lambda m: (m.project_role != "lead", (m.user.name or "").lower()))
            ],
            "milestones": [serialize_milestone(m) for m in project.milestones],
            "teams": [ref(t) | {"member_count": len(t.members), "color": t.color} for t in scoped(Team, ctx).filter(Team.project_id == project.id, Team.archived_at.is_(None)).order_by(Team.name).all()],
            "created_at": iso(project.created_at),
        }
    )
    return payload


from asme.ops.models import CostEntry as CostEntryAmount  # noqa: E402  (alias keeps the query above readable)


def serialize_milestone(m: Milestone) -> dict:
    return {
        "id": m.id,
        "project_id": m.project_id,
        "name": m.name,
        "description": m.description,
        "due_date": iso(m.due_date),
        "status": m.status,
        "owner": user_ref(m.owner),
        "weight": m.weight,
        "order": m.order,
        "completed_at": iso(m.completed_at),
        "is_overdue": bool(m.due_date and m.status != "complete" and m.due_date < date.today()),
    }


# --------------------------------------------------------------------------- queries


def list_projects(ctx, query):
    q = scoped(OpsProject, ctx)
    view = query.extra.get("view", "active")
    if view == "active":
        q = q.filter(OpsProject.archived_at.is_(None), OpsProject.status.in_(["planning", "active", "on_hold"]))
    elif view == "archived":
        q = q.filter((OpsProject.archived_at.isnot(None)) | (OpsProject.status == "archived"))
    filters = query.filters
    if "lead" in filters:
        q = q.filter(OpsProject.lead_user_id.in_([int(v) for v in filters["lead"] if v.isdigit()] or [-1]))
    if "status" in filters:
        q = q.filter(OpsProject.status.in_(filters["status"]))
    if "risk" in filters:
        q = q.filter(OpsProject.risk_level.in_(filters["risk"]))
    if "year" in filters:
        q = q.filter(OpsProject.academic_year.in_(filters["year"]))
    if "team" in filters:
        q = q.filter(OpsProject.id.in_(db.session.query(Team.project_id).filter(Team.id.in_(filters["team"]))))
    if query.q:
        like = f"%{query.q.lower()}%"
        q = q.filter(func.lower(OpsProject.name).like(like) | func.lower(OpsProject.code).like(like))
    order = {
        "name_asc": (OpsProject.name.asc(), OpsProject.id.asc()),
        "updated_desc": (OpsProject.updated_at.desc(), OpsProject.id.desc()),
        "target_asc": (OpsProject.target_date.asc().nullslast(), OpsProject.id.asc()),
        "risk_desc": (OpsProject.risk_level.desc(), OpsProject.name.asc()),
    }
    q = q.order_by(*order.get(query.sort, order["name_asc"]))
    offset = int((query.cursor or {}).get("offset", 0))
    rows = q.offset(offset).limit(query.limit + 1).all()
    has_more = len(rows) > query.limit
    rows = rows[: query.limit]
    counts = _wo_counts(ctx.organization.id, [p.id for p in rows])
    team_rows = db.session.query(Team.project_id, func.count(Team.id)).filter(Team.organization_id == ctx.organization.id, Team.archived_at.is_(None), Team.project_id.isnot(None)).group_by(Team.project_id).all()
    team_counts = {pid: int(n) for pid, n in team_rows}
    return rows, counts, team_counts, ({"offset": offset + query.limit} if has_more else None)


# --------------------------------------------------------------------------- writes


def _validate_people(ctx, lead_user_id, advisor_user_id):
    ids = [i for i in (lead_user_id, advisor_user_id) if i]
    users = resolve_users(ctx.organization.id, ids)
    if lead_user_id and lead_user_id not in users:
        raise Validation("Project lead must be an active member.", field="lead_user_id")
    if advisor_user_id and advisor_user_id not in users:
        raise Validation("Faculty advisor must be an active member.", field="faculty_advisor_user_id")
    return users


def _validate_public(public_project_id, project_id=None):
    if not public_project_id:
        return
    if db.session.get(PublicProject, public_project_id) is None:
        raise NotFound("Public project not found.")
    taken = OpsProject.query.filter(OpsProject.public_project_id == public_project_id, OpsProject.id != (project_id or "")).first()
    if taken:
        raise Conflict("That public project is already linked to another ops project.", code="public_project_taken")


def _ensure_lead_membership(project: OpsProject):
    if project.lead_user_id:
        row = OpsProjectMember.query.filter_by(project_id=project.id, user_id=project.lead_user_id).first()
        if row is None:
            db.session.add(OpsProjectMember(project_id=project.id, user_id=project.lead_user_id, project_role="lead"))
        else:
            row.project_role = "lead"
        db.session.flush()


def create_project(ctx, data) -> OpsProject:
    _validate_people(ctx, data.lead_user_id, data.faculty_advisor_user_id)
    _validate_public(data.public_project_id)
    code = (data.code or "").strip().upper() or next_code(ctx, data.name)
    if scoped(OpsProject, ctx).filter(func.lower(OpsProject.code) == code.lower()).first():
        raise Conflict("A project with that code already exists.", code="code_taken")
    project = OpsProject(organization_id=ctx.organization.id, created_by_user_id=ctx.user.id, updated_by_user_id=ctx.user.id)
    for key in PROJECT_FIELDS:
        setattr(project, key, getattr(data, key))
    project.code = code
    db.session.add(project)
    db.session.flush()
    _ensure_lead_membership(project)
    audit.record_event("project.created", "project", project.id, organization_id=ctx.organization.id, actor=ctx.user, after=audit.snapshot(project, PROJECT_FIELDS))
    db.session.commit()
    from asme import events

    events.emit(events.CHAPTER_CHANGED, reason="project_created")
    return project


def update_project(ctx, project: OpsProject, data, fields_set: set[str]) -> OpsProject:
    _validate_people(ctx, getattr(data, "lead_user_id", None) if "lead_user_id" in fields_set else None, getattr(data, "faculty_advisor_user_id", None) if "faculty_advisor_user_id" in fields_set else None)
    if "public_project_id" in fields_set:
        _validate_public(data.public_project_id, project.id)
    if "code" in fields_set and data.code:
        code = data.code.strip().upper()
        if scoped(OpsProject, ctx).filter(func.lower(OpsProject.code) == code.lower(), OpsProject.id != project.id).first():
            raise Conflict("A project with that code already exists.", code="code_taken")
        data.code = code
    before = audit.snapshot(project, PROJECT_FIELDS)
    apply_patch(project, data, fields_set, set(PROJECT_FIELDS))
    if "status" in fields_set:
        project.archived_at = datetime.utcnow() if data.status == "archived" else None
    touch(project, ctx.user)
    _ensure_lead_membership(project)
    after = audit.snapshot(project, PROJECT_FIELDS)
    audit.record_event("project.updated", "project", project.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"changed": audit.diff(before, after)})
    db.session.commit()
    return project


def archive_project(ctx, project: OpsProject) -> OpsProject:
    project.status = "archived"
    project.archived_at = datetime.utcnow()
    touch(project, ctx.user)
    audit.record_event("project.archived", "project", project.id, organization_id=ctx.organization.id, actor=ctx.user)
    db.session.commit()
    return project


def set_members(ctx, project: OpsProject, members) -> OpsProject:
    users = resolve_users(ctx.organization.id, [m.user_id for m in members])
    wanted = {}
    for m in members:
        if m.user_id not in users:
            raise Validation(f"User {m.user_id} is not an active member.", field="members")
        if m.team_id:
            get_or_404(Team, m.team_id, ctx, "Team")
        wanted[m.user_id] = m
    if project.lead_user_id and project.lead_user_id not in wanted:
        raise Validation("The project lead must remain a member; change the lead first.", field="members")
    before = {"members": sorted(project.member_user_ids)}
    existing = {m.user_id: m for m in project.members}
    for user_id, row in existing.items():
        if user_id not in wanted:
            db.session.delete(row)
        else:
            row.project_role = wanted[user_id].project_role
            row.team_id = wanted[user_id].team_id
    for user_id, m in wanted.items():
        if user_id not in existing:
            db.session.add(OpsProjectMember(project_id=project.id, user_id=user_id, project_role=m.project_role, team_id=m.team_id))
    db.session.flush()
    db.session.refresh(project)
    after = {"members": sorted(project.member_user_ids)}
    audit.record_event("project.members_set", "project", project.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after)
    db.session.commit()
    return project


def create_milestone(ctx, project: OpsProject, data) -> Milestone:
    if data.owner_user_id and data.owner_user_id not in resolve_users(ctx.organization.id, [data.owner_user_id]):
        raise Validation("Owner must be an active member.", field="owner_user_id")
    order = (db.session.query(func.coalesce(func.max(Milestone.order), 0)).filter(Milestone.project_id == project.id).scalar() or 0) + 1
    m = Milestone(organization_id=ctx.organization.id, project_id=project.id, created_by_user_id=ctx.user.id, updated_by_user_id=ctx.user.id, order=order)
    for key in ("name", "description", "due_date", "status", "owner_user_id", "weight"):
        setattr(m, key, getattr(data, key))
    if m.status == "complete":
        m.completed_at = datetime.utcnow()
    db.session.add(m)
    db.session.flush()
    audit.record_event("milestone.created", "milestone", m.id, organization_id=ctx.organization.id, actor=ctx.user, after=audit.snapshot(m, MILESTONE_FIELDS), metadata={"project_id": project.id})
    touch(project, ctx.user)
    db.session.commit()
    return m


def update_milestone(ctx, project: OpsProject, milestone: Milestone, data, fields_set: set[str]) -> Milestone:
    if milestone.project_id != project.id:
        raise NotFound("Milestone not found.")
    if "owner_user_id" in fields_set and data.owner_user_id and data.owner_user_id not in resolve_users(ctx.organization.id, [data.owner_user_id]):
        raise Validation("Owner must be an active member.", field="owner_user_id")
    before = audit.snapshot(milestone, MILESTONE_FIELDS)
    apply_patch(milestone, data, fields_set, set(MILESTONE_FIELDS))
    if "status" in fields_set:
        milestone.completed_at = datetime.utcnow() if data.status == "complete" else None
    touch(milestone, ctx.user)
    touch(project, ctx.user)
    after = audit.snapshot(milestone, MILESTONE_FIELDS)
    audit.record_event("milestone.updated", "milestone", milestone.id, organization_id=ctx.organization.id, actor=ctx.user, before=before, after=after, metadata={"project_id": project.id, "changed": audit.diff(before, after)})
    db.session.commit()
    return milestone


def delete_milestone(ctx, project: OpsProject, milestone: Milestone):
    if milestone.project_id != project.id:
        raise NotFound("Milestone not found.")
    audit.record_event("milestone.deleted", "milestone", milestone.id, organization_id=ctx.organization.id, actor=ctx.user, before=audit.snapshot(milestone, MILESTONE_FIELDS), metadata={"project_id": project.id})
    db.session.delete(milestone)
    db.session.commit()


# --------------------------------------------------------------------------- health & activity


def health(ctx, project: OpsProject) -> dict:
    counts = _wo_counts(ctx.organization.id, [project.id]).get(project.id, {"open": 0, "done": 0, "total": 0, "overdue": 0, "blocked": 0})
    now = datetime.utcnow()
    by_status = dict(db.session.query(WorkOrder.status, func.count(WorkOrder.id)).filter(WorkOrder.project_id == project.id).group_by(WorkOrder.status).all())
    by_priority = dict(db.session.query(WorkOrder.priority, func.count(WorkOrder.id)).filter(WorkOrder.project_id == project.id, WorkOrder.status.in_(WO_OPEN_STATUSES)).group_by(WorkOrder.priority).all())
    workload = db.session.query(Team.id, Team.name, func.count(WorkOrder.id)).join(WorkOrder, WorkOrder.team_id == Team.id).filter(WorkOrder.project_id == project.id, WorkOrder.status.in_(WO_OPEN_STATUSES)).group_by(Team.id, Team.name).all()
    milestones = [serialize_milestone(m) for m in project.milestones]
    at_risk = [m for m in milestones if m["is_overdue"] or (m["status"] != "complete" and m["due_date"] and (date.fromisoformat(m["due_date"]) - date.today()).days <= 14)]
    cost_total = db.session.query(func.coalesce(func.sum(CostEntryAmount.amount), 0)).join(WorkOrder, WorkOrder.id == CostEntryAmount.work_order_id).filter(WorkOrder.project_id == project.id).scalar() or 0
    return {
        "project_id": project.id,
        "completion_percent": completion_percent(project, counts),
        "work_orders": {"open": counts["open"], "done": counts["done"], "total": counts["total"], "overdue": counts["overdue"], "blocked": counts["blocked"], "by_status": {k: int(v) for k, v in by_status.items()}, "by_priority": {k: int(v) for k, v in by_priority.items()}},
        "workload_by_team": [{"team_id": tid, "team": name, "open_work_orders": int(n)} for tid, name, n in workload],
        "milestones": {"total": len(milestones), "complete": sum(1 for m in milestones if m["status"] == "complete"), "at_risk": at_risk, "next": next_milestone(project)},
        "budget": {"amount": float(project.budget_amount) if project.budget_amount is not None else None, "used": float(cost_total), "remaining": (float(project.budget_amount) - float(cost_total)) if project.budget_amount is not None else None},
        "hours_logged": round((db.session.query(func.coalesce(func.sum(WorkOrder.actual_minutes), 0)).filter(WorkOrder.project_id == project.id).scalar() or 0) / 60, 1),
        "risk_level": project.risk_level,
        "generated_at": now.isoformat(),
    }


def activity(ctx, project: OpsProject, limit=50) -> list[dict]:
    wo_ids = [r[0] for r in db.session.query(WorkOrder.id).filter(WorkOrder.project_id == project.id).all()]
    milestone_ids = [m.id for m in project.milestones]
    q = AuditEvent.query.filter(AuditEvent.organization_id == ctx.organization.id).filter(
        ((AuditEvent.entity_type == "project") & (AuditEvent.entity_id == project.id))
        | ((AuditEvent.entity_type == "work_order") & (AuditEvent.entity_id.in_(wo_ids or ["-"])))
        | ((AuditEvent.entity_type == "milestone") & (AuditEvent.entity_id.in_(milestone_ids or ["-"])))
    )
    rows = q.order_by(AuditEvent.occurred_at.desc()).limit(limit).all()
    from asme.ops.services.audit_log import serialize_event

    return [serialize_event(r) for r in rows]
