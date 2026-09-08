"""The Launchpad engine: evaluate a subject's tracks, advance phases, grant entitlements.

Progress is *derived*. Nothing here trusts a stored percentage - every call
re-reads the domain tables through the rule evaluators.
"""

from __future__ import annotations

import logging
from datetime import datetime, timedelta

from asme.auth.session import normalize_role, role_allows
from asme.constants import SUBJECT_CHAPTER, SUBJECT_USER
from asme.extensions import db
from asme.models import Phase, PhaseState, Task, TaskState, Track, TrainingCompletion, TrainingModule, User
from asme.services import audit
from asme.services.errors import Forbidden, NotFound, Validation
from asme.services.onboarding import entitlements as ent
from asme.services.onboarding.rules import HUMAN_RULES, RuleResult, Subject, evaluate_rule
from asme.utils import academic_year_key

log = logging.getLogger("asme.onboarding")


# --------------------------------------------------------------------------- subjects


def subject_for_user(user) -> Subject:
    return Subject(type=SUBJECT_USER, id=str(user.id), user=user)


def subject_for_chapter(year_key: str | None = None) -> Subject:
    return Subject(type=SUBJECT_CHAPTER, id=year_key or academic_year_key(), user=None)


def tracks_for(subject: Subject):
    return Track.query.filter(Track.audience == subject.type, Track.is_active.is_(True)).order_by(Track.id.asc()).all()


def _get_or_create_task_state(task, subject) -> TaskState:
    state = TaskState.query.filter_by(task_id=task.id, subject_type=subject.type, subject_id=subject.id).first()
    if state is None:
        state = TaskState(task_id=task.id, subject_type=subject.type, subject_id=subject.id, status="pending", current=0, target=1)
        db.session.add(state)
        db.session.flush()
    return state


def _get_or_create_phase_state(phase, subject) -> PhaseState:
    state = PhaseState.query.filter_by(phase_id=phase.id, subject_type=subject.type, subject_id=subject.id).first()
    if state is None:
        state = PhaseState(phase_id=phase.id, subject_type=subject.type, subject_id=subject.id, status="locked")
        db.session.add(state)
        db.session.flush()
    return state


def _apply_result(state: TaskState, result: RuleResult):
    state.current = int(result.current)
    state.target = max(int(result.target), 1)
    if result.complete:
        if state.status != "complete":
            state.completed_at = datetime.utcnow()
        state.status = "complete"
    else:
        state.status = "in_progress" if state.current > 0 else "pending"
        state.completed_at = None
    state.evidence = result.evidence


def _apply_grants(phase, track, subject, granted_by=None):
    if not subject.user:
        return
    source_ref = f"{track.key}:{phase.key}"
    for key in phase.grants:
        if key.startswith("role:"):
            wanted = normalize_role(key.split(":", 1)[1])
            if not role_allows(subject.user.role, wanted):
                subject.user.role = wanted
                audit.record("promote_role", f"user_id={subject.user.id} role={wanted} via={source_ref}")
            continue
        ent.grant(subject.user, key, source="phase", source_ref=source_ref, granted_by=granted_by)


def _revoke_grants(phase, track, subject):
    if not subject.user:
        return
    ent.revoke_phase_grants(subject.user, f"{track.key}:{phase.key}")


# --------------------------------------------------------------------------- evaluation


def evaluate(subject: Subject, commit: bool = True) -> dict:
    """Re-evaluate every active track for ``subject``. Returns the progress payload."""
    now = datetime.utcnow()
    for track in tracks_for(subject):
        previous_complete = True
        for phase in track.phases:
            phase_state = _get_or_create_phase_state(phase, subject)
            all_required_done = True
            for task in phase.tasks:
                task_state = _get_or_create_task_state(task, subject)
                result = evaluate_rule(task, subject, task_state)
                if task.rule_type in HUMAN_RULES:
                    # Human-marked tasks keep their stored status; refresh counters only.
                    task_state.current = int(result.current)
                    task_state.target = max(int(result.target), 1)
                else:
                    _apply_result(task_state, result)
                if task.is_required and not task_state.is_complete:
                    all_required_done = False

            if phase_state.status == "locked" and previous_complete:
                phase_state.status = "active"
                phase_state.started_at = phase_state.started_at or now

            if phase_state.status == "complete" and not all_required_done:
                # Regression (e.g. training expired, tool overdue): drop back and revoke.
                phase_state.status = "active"
                phase_state.completed_at = None
                _revoke_grants(phase, track, subject)
                log.info("phase regressed subject=%s:%s phase=%s", subject.type, subject.id, phase.key)
            elif phase_state.status == "active" and all_required_done:
                phase_state.status = "complete"
                phase_state.completed_at = now
                _apply_grants(phase, track, subject)
                log.info("phase completed subject=%s:%s phase=%s", subject.type, subject.id, phase.key)
            elif phase_state.status == "complete" and all_required_done:
                _apply_grants(phase, track, subject)  # idempotent re-assert (e.g. after admin revoke expiry)

            previous_complete = phase_state.status == "complete"
            if not previous_complete:
                # Everything after an incomplete phase is locked.
                pass
        # Lock phases that follow the first incomplete one.
        seen_incomplete = False
        for phase in track.phases:
            phase_state = _get_or_create_phase_state(phase, subject)
            if seen_incomplete and phase_state.status != "complete":
                phase_state.status = "locked"
            if phase_state.status != "complete":
                seen_incomplete = True
    if commit:
        db.session.commit()
    return progress(subject)


def evaluate_user(user, commit=True):
    return evaluate(subject_for_user(user), commit=commit)


def evaluate_chapter(year_key=None, commit=True):
    return evaluate(subject_for_chapter(year_key), commit=commit)


# --------------------------------------------------------------------------- progress payload


def progress(subject: Subject) -> dict:
    tracks_payload = []
    for track in tracks_for(subject):
        phases_payload = []
        required_total = required_done = 0
        current_phase = None
        for phase in track.phases:
            phase_state = PhaseState.query.filter_by(phase_id=phase.id, subject_type=subject.type, subject_id=subject.id).first()
            tasks_payload = []
            phase_required = phase_required_done = 0
            for task in phase.tasks:
                task_state = TaskState.query.filter_by(task_id=task.id, subject_type=subject.type, subject_id=subject.id).first()
                status = task_state.status if task_state else "pending"
                done = status == "complete"
                if task.is_required:
                    phase_required += 1
                    phase_required_done += 1 if done else 0
                tasks_payload.append(
                    {
                        "key": task.key,
                        "name": task.name,
                        "description": task.description,
                        "required": bool(task.is_required),
                        "status": status,
                        "current": task_state.current if task_state else 0,
                        "target": task_state.target if task_state else 1,
                        "est_minutes": task.est_minutes,
                        "rule_type": task.rule_type,
                        "rule_config": task.rule_config,
                        "cta_route": task.cta_route,
                        "cta_label": task.cta_label,
                        "help_url": task.help_url,
                        "evidence": task_state.evidence if task_state else {},
                        "completed_at": task_state.completed_at.isoformat() if task_state and task_state.completed_at else None,
                        "note": task_state.note if task_state else None,
                        "human": task.rule_type in HUMAN_RULES,
                    }
                )
            status = phase_state.status if phase_state else "locked"
            payload = {
                "key": phase.key,
                "name": phase.name,
                "tagline": phase.tagline,
                "description": phase.description,
                "order": phase.order,
                "status": status,
                "est_minutes": phase.est_minutes,
                "grants": phase.grants,
                "tasks": tasks_payload,
                "required_total": phase_required,
                "required_done": phase_required_done,
                "steps_left": max(phase_required - phase_required_done, 0),
                "started_at": phase_state.started_at.isoformat() if phase_state and phase_state.started_at else None,
                "completed_at": phase_state.completed_at.isoformat() if phase_state and phase_state.completed_at else None,
            }
            phases_payload.append(payload)
            required_total += phase_required
            required_done += phase_required_done
            if current_phase is None and status == "active":
                current_phase = payload
        percent = int(round((required_done / required_total) * 100)) if required_total else 0
        tracks_payload.append(
            {
                "key": track.key,
                "name": track.name,
                "audience": track.audience,
                "description": track.description,
                "percent": percent,
                "required_total": required_total,
                "required_done": required_done,
                "steps_left": current_phase["steps_left"] if current_phase else 0,
                "current_phase": current_phase,
                "complete": bool(phases_payload) and all(p["status"] == "complete" for p in phases_payload),
                "phases": phases_payload,
            }
        )
    return {
        "subject": {"type": subject.type, "id": subject.id},
        "tracks": tracks_payload,
        "percent": tracks_payload[0]["percent"] if tracks_payload else 0,
        "steps_left": tracks_payload[0]["steps_left"] if tracks_payload else 0,
        "entitlements": sorted(ent.user_entitlements(subject.user)) if subject.user else [],
    }


def summary_for_user(user) -> dict:
    """Cheap read for dashboards: current phase + percent. Evaluates once for a
    subject the engine has never seen (e.g. accounts that predate Launchpad)."""
    subject = subject_for_user(user)
    has_state = PhaseState.query.filter_by(subject_type=subject.type, subject_id=subject.id).first() is not None
    payload = progress(subject) if has_state else evaluate(subject)
    track = payload["tracks"][0] if payload["tracks"] else None
    return {
        "percent": payload["percent"],
        "steps_left": payload["steps_left"],
        "current_phase": track["current_phase"] if track else None,
        "complete": track["complete"] if track else False,
        "entitlements": payload["entitlements"],
    }


# --------------------------------------------------------------------------- human-marked tasks


def find_task(task_key: str, audience: str) -> Task | None:
    return (
        Task.query.join(Phase, Task.phase_id == Phase.id)
        .join(Track, Phase.track_id == Track.id)
        .filter(Task.key == task_key, Track.audience == audience, Track.is_active.is_(True))
        .order_by(Task.id.asc())
        .first()
    )


def complete_task(task_key: str, subject: Subject, actor, note: str | None = None) -> TaskState:
    """Mark a ``manual``/``signoff`` task complete. Rule-driven tasks are rejected."""
    task = find_task(task_key, subject.type)
    if task is None:
        raise NotFound("Task not found.", code="task_not_found")
    if task.rule_type not in HUMAN_RULES:
        raise Validation(
            "That task is completed automatically when its requirement is met.",
            code="task_not_manual",
            rule_type=task.rule_type,
        )
    if task.rule_type == "signoff":
        by_role = task.rule_config.get("by_role") or "team_leader"
        if not actor or not role_allows(actor.role, by_role):
            raise Forbidden(f"Only a {by_role.replace('_', ' ')} or above can sign this off.", code="signoff_role")
        if subject.user and actor.id == subject.user.id and not role_allows(actor.role, "admin"):
            raise Forbidden("You cannot sign off your own task.", code="self_signoff")
    else:  # manual
        if subject.is_user and not (actor and (actor.id == subject.user.id or role_allows(actor.role, "team_leader"))):
            raise Forbidden("You can only complete your own tasks.", code="not_owner")
        if subject.is_chapter and not (actor and role_allows(actor.role, "admin")):
            raise Forbidden("Only an admin can complete chapter setup tasks.", code="admin_only")
        if task.rule_config.get("note_required") and not (note or "").strip():
            raise Validation("A note is required to complete this task.", field="note")

    state = _get_or_create_task_state(task, subject)
    state.status = "complete"
    state.current = 1
    state.target = 1
    state.completed_at = datetime.utcnow()
    state.completed_by_user_id = actor.id if actor else None
    state.note = (note or "").strip()[:300] or None
    audit.record("complete_onboarding_task", f"task={task.key} subject={subject.type}:{subject.id}", actor=actor)
    db.session.commit()
    evaluate(subject)
    return state


def reopen_task(task_key: str, subject: Subject, actor, note=None) -> TaskState:
    task = find_task(task_key, subject.type)
    if task is None:
        raise NotFound("Task not found.", code="task_not_found")
    if not (actor and role_allows(actor.role, "admin")):
        raise Forbidden("Only an admin can reopen a task.", code="admin_only")
    state = _get_or_create_task_state(task, subject)
    state.status = "pending"
    state.current = 0
    state.completed_at = None
    state.completed_by_user_id = None
    state.note = (note or "").strip()[:300] or None
    audit.record("reopen_onboarding_task", f"task={task.key} subject={subject.type}:{subject.id}", actor=actor)
    db.session.commit()
    evaluate(subject)
    return state


# --------------------------------------------------------------------------- training


def record_training(user, module_key, score=None, actor=None, note=None) -> TrainingCompletion:
    module = TrainingModule.query.filter_by(key=module_key).first()
    if module is None:
        raise NotFound("Training module not found.", code="module_not_found")
    if actor is not None and actor.id != user.id and not role_allows(actor.role, "team_leader"):
        raise Forbidden("Only a team lead or admin can record training for someone else.", code="signoff_role")
    if actor is not None and actor.id == user.id and module.min_score and not role_allows(actor.role, "admin"):
        raise Forbidden("Scored modules must be recorded by a team lead or admin.", code="signoff_role")
    if score is not None:
        score = int(score)
        if score < 0 or score > 100:
            raise Validation("Score must be between 0 and 100.", field="score")
    expires_at = None
    if module.validity_days:
        expires_at = datetime.utcnow() + timedelta(days=int(module.validity_days))
    row = TrainingCompletion(
        module_id=module.id,
        user_id=user.id,
        score=score,
        expires_at=expires_at,
        recorded_by_user_id=actor.id if actor else None,
        note=(note or "").strip()[:300] or None,
    )
    db.session.add(row)
    audit.record("record_training", f"user_id={user.id} module={module.key} score={score}", actor=actor)
    db.session.commit()
    from asme import events

    events.emit(events.TRAINING_COMPLETED, user_id=user.id, module=module.key, score=score)
    return row


# --------------------------------------------------------------------------- admin overview


def member_overview():
    """Per-user phase summary for the admin Launchpad page."""
    users = (
        User.query.filter(User.is_active.is_(True), User.role.in_(["member", "team_leader"]))
        .order_by(User.name.asc(), User.id.asc())
        .all()
    )
    rows = []
    for user in users:
        summary = summary_for_user(user)
        rows.append({"user": user, **summary})
    return rows
