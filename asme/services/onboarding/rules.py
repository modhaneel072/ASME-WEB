"""Task rule evaluators.

A rule never stores a copy of domain state: it queries the tables that already
hold the truth and returns a ``RuleResult``. Adding a rule type means adding
one function here and registering it in ``RULES``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

from sqlalchemy import func

from asme.constants import LOAN_STATE_OUT, LOAN_STATE_RETURNED, SUBJECT_CHAPTER, SUBJECT_USER
from asme.extensions import db
from asme.models import (
    AttendanceRecord,
    Event,
    Item,
    Loan,
    NFCTag,
    Phase,
    PhaseState,
    PrintRequest,
    Project,
    ProjectMembership,
    TrainingCompletion,
    TrainingModule,
    User,
    WorkLog,
)
from asme.utils import semester_start


@dataclass
class Subject:
    type: str
    id: str
    user: User | None = None

    @property
    def is_user(self) -> bool:
        return self.type == SUBJECT_USER

    @property
    def is_chapter(self) -> bool:
        return self.type == SUBJECT_CHAPTER


@dataclass
class RuleResult:
    current: int = 0
    target: int = 1
    complete: bool = False
    evidence: dict = field(default_factory=dict)


def _window_start(config) -> date | None:
    window = (config.get("window") or "").strip().lower()
    if not window:
        return None
    if window == "semester":
        return semester_start()
    if window.startswith("days:"):
        try:
            return date.today() - timedelta(days=int(window.split(":", 1)[1]))
        except ValueError:
            return None
    return None


def _stored(state) -> RuleResult:
    """Rules that rely on a human marking them done just reflect stored state."""
    if state is not None and state.status == "complete":
        return RuleResult(current=1, target=1, complete=True, evidence={"completed_by": state.completed_by_user_id})
    return RuleResult(current=0, target=1, complete=False)


# --------------------------------------------------------------------------- user rules


def rule_manual(subject, config, state):
    return _stored(state)


def rule_signoff(subject, config, state):
    return _stored(state)


def rule_existence(subject, config, state):
    if not subject.user:
        return RuleResult()
    user = subject.user
    fields = list(config.get("fields") or [])
    if config.get("field"):
        fields.append(config["field"])
    missing = []
    for name in fields:
        if name == "nfc":
            has_tag = bool(user.nfc_uid) or NFCTag.query.filter_by(user_id=user.id, active=True).first() is not None
            if not has_tag:
                missing.append("nfc")
        elif name == "team":
            on_team = (
                ProjectMembership.query.filter(ProjectMembership.user_id == user.id, ProjectMembership.left_at.is_(None)).first()
                is not None
            )
            if not on_team:
                missing.append("team")
        else:
            value = getattr(user, name, None)
            if value in (None, "", 0):
                missing.append(name)
    total = max(len(fields), 1)
    return RuleResult(current=total - len(missing), target=total, complete=not missing, evidence={"missing": missing})


def rule_email_domain(subject, config, state):
    if not subject.user:
        return RuleResult()
    domains = [d.strip().lower().lstrip("@") for d in (config.get("domains") or []) if d.strip()]
    email = (subject.user.email or "").strip().lower()
    domain = email.split("@", 1)[1] if "@" in email else ""
    ok = bool(domains) and any(domain == d or domain.endswith("." + d) for d in domains)
    return RuleResult(current=1 if ok else 0, target=1, complete=ok, evidence={"domain": domain})


def _count_query(model, subject, config):
    user = subject.user
    since = _window_start(config)
    where = config.get("where") or {}
    if model is AttendanceRecord:
        query = AttendanceRecord.query.filter(AttendanceRecord.user_id == user.id)
        if config.get("exclude_rsvp", True):
            query = query.filter(AttendanceRecord.checkin_method != "rsvp")
        if since:
            query = query.filter(AttendanceRecord.checkin_time >= datetime.combine(since, datetime.min.time()))
        return query
    if model is Loan:
        query = Loan.query.filter(Loan.user_id == user.id)
        if where.get("status"):
            query = query.filter(Loan.status == where["status"])
        if since:
            query = query.filter(Loan.timestamp >= datetime.combine(since, datetime.min.time()))
        return query
    if model is PrintRequest:
        query = PrintRequest.query.filter(PrintRequest.user_id == user.id)
        if where.get("status"):
            query = query.filter(PrintRequest.status == where["status"])
        if since:
            query = query.filter(PrintRequest.created_at >= datetime.combine(since, datetime.min.time()))
        return query
    if model is ProjectMembership:
        return ProjectMembership.query.filter(ProjectMembership.user_id == user.id, ProjectMembership.left_at.is_(None))
    raise ValueError(f"count_threshold does not support model {model.__name__}")


USER_COUNT_MODELS = {
    "AttendanceRecord": AttendanceRecord,
    "Loan": Loan,
    "Transaction": Loan,
    "PrintRequest": PrintRequest,
    "ProjectMembership": ProjectMembership,
}


def rule_count_threshold(subject, config, state):
    if not subject.user:
        return RuleResult()
    model = USER_COUNT_MODELS.get(config.get("model") or "")
    if model is None:
        return RuleResult()
    minimum = max(int(config.get("min") or 1), 1)
    count = _count_query(model, subject, config).count()
    return RuleResult(current=min(count, minimum), target=minimum, complete=count >= minimum, evidence={"count": count})


def rule_hours_threshold(subject, config, state):
    if not subject.user:
        return RuleResult()
    minimum = max(int(config.get("min") or 1), 1)
    since = _window_start(config)
    query = db.session.query(func.coalesce(func.sum(WorkLog.hours), 0.0)).filter(WorkLog.user_id == subject.user.id)
    if since:
        query = query.filter(WorkLog.logged_for >= since)
    if config.get("approved_only"):
        query = query.filter(WorkLog.approved_at.isnot(None))
    hours = float(query.scalar() or 0.0)
    return RuleResult(current=int(min(hours, minimum)), target=minimum, complete=hours >= minimum, evidence={"hours": round(hours, 2)})


def rule_clean_streak(subject, config, state):
    """N on-time returns and nothing currently overdue."""
    if not subject.user:
        return RuleResult()
    minimum = max(int(config.get("min") or 1), 1)
    since = _window_start(config)
    query = Loan.query.filter(Loan.user_id == subject.user.id, Loan.status == LOAN_STATE_RETURNED)
    if since:
        query = query.filter(Loan.timestamp >= datetime.combine(since, datetime.min.time()))
    on_time = sum(1 for loan in query.all() if loan.returned_on_time)
    overdue_open = Loan.query.filter(
        Loan.user_id == subject.user.id,
        Loan.status == LOAN_STATE_OUT,
        Loan.due_date.isnot(None),
        Loan.due_date < date.today(),
    ).count()
    complete = on_time >= minimum and overdue_open == 0
    return RuleResult(
        current=min(on_time, minimum),
        target=minimum,
        complete=complete,
        evidence={"on_time_returns": on_time, "overdue_open": overdue_open},
    )


def rule_loan_signed_off(subject, config, state):
    if not subject.user:
        return RuleResult()
    minimum = max(int(config.get("min") or 1), 1)
    count = Loan.query.filter(Loan.user_id == subject.user.id, Loan.signed_off_by_user_id.isnot(None)).count()
    return RuleResult(current=min(count, minimum), target=minimum, complete=count >= minimum, evidence={"signed_off_loans": count})


def rule_training(subject, config, state):
    if not subject.user:
        return RuleResult()
    module = TrainingModule.query.filter_by(key=config.get("module") or "").first()
    if not module:
        return RuleResult(evidence={"error": "module not found"})
    min_score = config.get("min_score", module.min_score)
    rows = (
        TrainingCompletion.query.filter(TrainingCompletion.module_id == module.id, TrainingCompletion.user_id == subject.user.id)
        .order_by(TrainingCompletion.completed_at.desc())
        .all()
    )
    for row in rows:
        if not row.is_valid:
            continue
        if min_score is not None and (row.score is None or row.score < int(min_score)):
            continue
        return RuleResult(
            current=1,
            target=1,
            complete=True,
            evidence={"score": row.score, "completed_at": row.completed_at, "expires_at": row.expires_at},
        )
    return RuleResult(current=0, target=1, complete=False, evidence={"attempts": len(rows)})


# --------------------------------------------------------------------------- chapter rules


def _chapter_count(config) -> int:
    model_name = config.get("model") or ""
    where = config.get("where") or {}
    if model_name == "User":
        query = User.query.filter(User.is_active.is_(True))
        roles = where.get("role_in")
        if roles:
            query = query.filter(User.role.in_(list(roles)))
        return query.count()
    if model_name == "Item":
        query = Item.query
        if where.get("active", True):
            query = query.filter(Item.active.is_(True))
        return query.count()
    if model_name == "ItemLocation":
        return (
            db.session.query(func.count(func.distinct(func.lower(Item.location))))
            .filter(Item.location.isnot(None), Item.location != "")
            .scalar()
            or 0
        )
    if model_name == "Project":
        query = Project.query
        if where.get("joinable"):
            query = query.filter(Project.is_joinable.is_(True))
        return query.count()
    if model_name == "Event":
        query = Event.query.filter(Event.status == "scheduled")
        if where.get("future", True):
            query = query.filter(Event.start_time >= datetime.now())
        if where.get("kind"):
            query = query.filter(Event.kind == where["kind"])
        return query.count()
    if model_name == "NFCTag":
        return NFCTag.query.filter(NFCTag.active.is_(True)).count()
    return 0


def rule_chapter_count(subject, config, state):
    minimum = max(int(config.get("min") or 1), 1)
    count = _chapter_count(config)
    return RuleResult(current=min(count, minimum), target=minimum, complete=count >= minimum, evidence={"count": count})


def rule_config_set(subject, config, state):
    from asme.config import settings
    from asme.services import scheduling

    which = (config.get("setting") or "").strip().lower()
    ok = False
    detail = ""
    if which == "calendar":
        status = scheduling.provider_status()
        ok = bool(status.enabled or scheduling.calendar_embed_url())
        detail = "; ".join(status.errors) if not ok else "connected"
    elif which == "smtp":
        cfg = settings()
        ok = bool(cfg.smtp_user and cfg.smtp_pass)
        detail = "configured" if ok else "ASME_SMTP_USER/ASME_SMTP_PASS missing"
    elif which == "printers":
        cfg = settings()
        ok = bool(cfg.h2s_print_cmd and cfg.p1s_print_cmd)
        detail = "configured" if ok else "print commands missing"
    return RuleResult(current=1 if ok else 0, target=1, complete=ok, evidence={"detail": detail})


def rule_member_phase_ratio(subject, config, state):
    phase_key = config.get("phase") or ""
    min_ratio = float(config.get("min_ratio") or 0.5)
    members = User.query.filter(User.is_active.is_(True), User.role.in_(["member", "team_leader"])).count()
    if members == 0:
        return RuleResult(current=0, target=100, complete=False, evidence={"members": 0})
    phase_ids = [row.id for row in Phase.query.filter(Phase.key == phase_key).all()]
    if not phase_ids:
        return RuleResult(current=0, target=100, complete=False, evidence={"error": "phase not found"})
    done = (
        PhaseState.query.filter(
            PhaseState.phase_id.in_(phase_ids),
            PhaseState.subject_type == SUBJECT_USER,
            PhaseState.status == "complete",
        )
        .with_entities(PhaseState.subject_id)
        .distinct()
        .count()
    )
    ratio = done / members
    target_pct = int(round(min_ratio * 100))
    return RuleResult(
        current=min(int(round(ratio * 100)), target_pct),
        target=target_pct,
        complete=ratio >= min_ratio,
        evidence={"members": members, "complete": done, "ratio": round(ratio, 3)},
    )


RULES = {
    "manual": rule_manual,
    "signoff": rule_signoff,
    "existence": rule_existence,
    "email_domain": rule_email_domain,
    "count_threshold": rule_count_threshold,
    "hours_threshold": rule_hours_threshold,
    "clean_streak": rule_clean_streak,
    "loan_signed_off": rule_loan_signed_off,
    "training": rule_training,
    "chapter_count": rule_chapter_count,
    "config_set": rule_config_set,
    "member_phase_ratio": rule_member_phase_ratio,
}

HUMAN_RULES = {"manual", "signoff"}


def evaluate_rule(task, subject, state) -> RuleResult:
    handler = RULES.get(task.rule_type)
    if handler is None:
        return RuleResult(evidence={"error": f"unknown rule_type {task.rule_type}"})
    return handler(subject, task.rule_config, state)
