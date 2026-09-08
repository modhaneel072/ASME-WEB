"""Content & community: projects, team membership, work logs, announcements, contact inbox."""

from __future__ import annotations

import json
from datetime import date, datetime
from uuid import uuid4

from sqlalchemy import func

from asme import events
from asme.constants import CONTACT_STATUSES
from asme.extensions import db
from asme.models import Announcement, ContactMessage, Project, ProjectMembership, WorkLog
from asme.services import audit
from asme.services.errors import Conflict, NotFound, Validation
from asme.utils import parse_float, parse_json_list, parse_positive_int, slugify

# --------------------------------------------------------------------------- projects


def save_project(form, actor) -> Project:
    project_id = parse_positive_int(form.get("project_id"), default=0)
    title = (form.get("title") or "").strip()
    if not title:
        raise Validation("Project title is required.", field="title")
    project = db.session.get(Project, project_id) if project_id else None
    if not project:
        project = Project(
            slug=slugify(title) or f"project-{uuid4().hex[:8]}",
            title=title[:200],
            summary="",
            description="",
            status="Active",
        )
        db.session.add(project)

    project.title = title[:200]
    project.slug = slugify(form.get("slug") or project.title) or project.slug
    project.summary = (form.get("summary") or "").strip()[:320] or "Project summary pending."
    project.description = (form.get("description") or "").strip() or "Project details pending."
    project.project_type = (form.get("project_type") or "").strip()[:80] or project.project_type or "General"
    project.status = (form.get("status") or "").strip()[:80] or "Active"
    project.timeline = (form.get("timeline") or "").strip() or None
    gallery_text = (form.get("gallery_json") or "").strip()
    if gallery_text:
        project.gallery_json = json.dumps(parse_json_list(gallery_text))
    elif project.gallery_json is None:
        project.gallery_json = json.dumps([])
    project.lead_name = (form.get("lead_name") or "").strip()[:160] or None
    project.image_url = (form.get("image_url") or "").strip()[:500] or None
    project.external_link = (form.get("external_link") or "").strip()[:500] or None
    if form.get("is_joinable") is not None:
        project.is_joinable = (form.get("is_joinable") or "1").strip() in {"1", "true", "yes", "on"}
    audit.record("save_project", f"project_id={project.id or 'new'} title={project.title}", actor=actor)
    db.session.commit()
    events.emit(events.CHAPTER_CHANGED, reason="project_saved")
    return project


def joinable_projects():
    return Project.query.filter(Project.is_joinable.is_(True)).order_by(Project.title.asc(), Project.id.asc()).all()


def active_memberships(user):
    return (
        ProjectMembership.query.filter(ProjectMembership.user_id == user.id, ProjectMembership.left_at.is_(None))
        .order_by(ProjectMembership.joined_at.asc())
        .all()
    )


def join_project(user, project) -> ProjectMembership:
    if not project.is_joinable:
        raise Validation("That team is not accepting members right now.", code="not_joinable")
    existing = ProjectMembership.query.filter_by(project_id=project.id, user_id=user.id).first()
    if existing and existing.left_at is None:
        raise Conflict("You are already on that team.", code="already_member")
    if existing:
        existing.left_at = None
        existing.joined_at = datetime.utcnow()
        row = existing
    else:
        row = ProjectMembership(project_id=project.id, user_id=user.id, role="member")
        db.session.add(row)
    db.session.commit()
    events.emit(events.TEAM_JOINED, user_id=user.id, project_id=project.id)
    return row


def leave_project(user, project):
    row = ProjectMembership.query.filter_by(project_id=project.id, user_id=user.id).first()
    if not row or row.left_at is not None:
        raise NotFound("You are not on that team.")
    row.left_at = datetime.utcnow()
    db.session.commit()
    events.emit(events.USER_UPDATED, user_id=user.id)


def log_hours(user, project, hours, note=None, logged_for=None) -> WorkLog:
    hours = parse_float(hours, default=0.0)
    if hours <= 0 or hours > 24:
        raise Validation("Hours must be between 0 and 24.", field="hours")
    row = WorkLog(
        user_id=user.id,
        project_id=project.id if project else None,
        hours=round(hours, 2),
        note=(note or "").strip()[:300] or None,
        logged_for=logged_for or date.today(),
    )
    db.session.add(row)
    db.session.commit()
    events.emit(events.HOURS_LOGGED, user_id=user.id, project_id=project.id if project else None, hours=row.hours)
    return row


def approve_hours(row, actor):
    row.approved_by_user_id = actor.id
    row.approved_at = datetime.utcnow()
    db.session.commit()
    events.emit(events.HOURS_LOGGED, user_id=row.user_id, project_id=row.project_id, hours=row.hours)


def total_hours(user, since=None, approved_only=False) -> float:
    query = db.session.query(func.coalesce(func.sum(WorkLog.hours), 0.0)).filter(WorkLog.user_id == user.id)
    if since:
        query = query.filter(WorkLog.logged_for >= since)
    if approved_only:
        query = query.filter(WorkLog.approved_at.isnot(None))
    return float(query.scalar() or 0.0)


# --------------------------------------------------------------------------- announcements


def update_announcement(row, form, actor) -> Announcement:
    row.title = (form.get("title") or row.title).strip()[:220]
    row.body = (form.get("body") or row.body).strip()[:10000]
    row.is_published = (form.get("is_published") or "1").strip() == "1"
    row.show_on_public = (form.get("show_on_public") or "1").strip() == "1"
    row.show_on_member = (form.get("show_on_member") or "1").strip() == "1"
    row.published_at = datetime.utcnow() if row.is_published else None
    audit.record("update_announcement", f"announcement_id={row.id}", actor=actor)
    db.session.commit()
    return row


def delete_announcement(row, actor):
    audit.record("delete_announcement", f"announcement_id={row.id}", actor=actor)
    db.session.delete(row)
    db.session.commit()


def public_announcements(limit=5):
    return (
        Announcement.query.filter_by(is_published=True, show_on_public=True)
        .order_by(func.coalesce(Announcement.published_at, Announcement.created_at).desc(), Announcement.id.desc())
        .limit(limit)
        .all()
    )


def member_announcements(limit=8):
    return (
        Announcement.query.filter_by(is_published=True, show_on_member=True)
        .order_by(func.coalesce(Announcement.published_at, Announcement.created_at).desc(), Announcement.id.desc())
        .limit(limit)
        .all()
    )


# --------------------------------------------------------------------------- contact inbox


def create_contact_message(*, name, email, message, kind="contact", subject=None, target=None, user=None) -> ContactMessage:
    name = (name or "").strip()
    email = (email or "").strip()
    message = (message or "").strip()
    if not name or not email or not message:
        raise Validation("Name, email, and message are required.")
    row = ContactMessage(
        name=name[:160],
        email=email[:160],
        kind=kind,
        user_id=user.id if user else None,
        target=(target or "")[:80] or None,
        subject=(subject or "")[:220] or None,
        message=message[:5000],
        status="new",
    )
    db.session.add(row)
    db.session.commit()
    return row


def update_contact_status(row, status, admin_reply, actor) -> ContactMessage:
    status = (status or "new").strip().lower()
    if status not in CONTACT_STATUSES:
        status = "new"
    row.status = status
    row.admin_reply = (admin_reply or "").strip()[:10000] or None
    audit.record("update_contact_status", f"message_id={row.id} status={status}", actor=actor)
    db.session.commit()
    return row
