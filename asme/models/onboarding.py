"""Launchpad - the onboarding engine's tables.

``TaskState`` / ``PhaseState`` key on ``(subject_type, subject_id)`` rather than a
user id so the same engine drives both the per-member track and the per-chapter
(academic year) track.
"""

import json
from datetime import datetime

from asme.extensions import db


class Track(db.Model):
    __tablename__ = "onboarding_tracks"
    id = db.Column(db.Integer, primary_key=True)

    key = db.Column(db.String(60), nullable=False, unique=True, index=True)
    name = db.Column(db.String(160), nullable=False)
    audience = db.Column(db.String(20), nullable=False, index=True)  # user / chapter
    description = db.Column(db.Text, nullable=True)
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    phases = db.relationship("Phase", back_populates="track", order_by="Phase.order")


class Phase(db.Model):
    __tablename__ = "onboarding_phases"
    __table_args__ = (db.UniqueConstraint("track_id", "key", name="uq_phase_track_key"),)
    id = db.Column(db.Integer, primary_key=True)

    track_id = db.Column(db.Integer, db.ForeignKey("onboarding_tracks.id"), nullable=False, index=True)
    key = db.Column(db.String(60), nullable=False)
    name = db.Column(db.String(160), nullable=False)
    tagline = db.Column(db.String(200), nullable=True)
    description = db.Column(db.Text, nullable=True)
    order = db.Column(db.Integer, nullable=False, default=0)
    est_minutes = db.Column(db.Integer, nullable=True)
    grants_json = db.Column(db.Text, nullable=False, default="[]")

    track = db.relationship("Track", back_populates="phases")
    tasks = db.relationship("Task", back_populates="phase", order_by="Task.order")

    @property
    def grants(self) -> list[str]:
        try:
            value = json.loads(self.grants_json or "[]")
        except Exception:
            return []
        return [str(item) for item in value if str(item).strip()]

    @grants.setter
    def grants(self, value):
        self.grants_json = json.dumps(list(value or []))


class Task(db.Model):
    __tablename__ = "onboarding_tasks"
    __table_args__ = (db.UniqueConstraint("phase_id", "key", name="uq_task_phase_key"),)
    id = db.Column(db.Integer, primary_key=True)

    phase_id = db.Column(db.Integer, db.ForeignKey("onboarding_phases.id"), nullable=False, index=True)
    key = db.Column(db.String(60), nullable=False, index=True)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    is_required = db.Column(db.Boolean, nullable=False, default=True)
    est_minutes = db.Column(db.Integer, nullable=True)
    rule_type = db.Column(db.String(40), nullable=False, default="manual")
    rule_config_json = db.Column(db.Text, nullable=False, default="{}")
    cta_route = db.Column(db.String(200), nullable=True)
    cta_label = db.Column(db.String(80), nullable=True)
    help_url = db.Column(db.String(500), nullable=True)
    order = db.Column(db.Integer, nullable=False, default=0)

    phase = db.relationship("Phase", back_populates="tasks")

    @property
    def rule_config(self) -> dict:
        try:
            value = json.loads(self.rule_config_json or "{}")
        except Exception:
            return {}
        return value if isinstance(value, dict) else {}

    @rule_config.setter
    def rule_config(self, value):
        self.rule_config_json = json.dumps(dict(value or {}))


class TaskState(db.Model):
    __tablename__ = "onboarding_task_states"
    __table_args__ = (
        db.UniqueConstraint("task_id", "subject_type", "subject_id", name="uq_task_state_subject"),
        db.Index("ix_task_state_subject", "subject_type", "subject_id"),
    )
    id = db.Column(db.Integer, primary_key=True)

    task_id = db.Column(db.Integer, db.ForeignKey("onboarding_tasks.id"), nullable=False, index=True)
    subject_type = db.Column(db.String(20), nullable=False)
    subject_id = db.Column(db.String(60), nullable=False)
    status = db.Column(db.String(20), nullable=False, default="pending")  # pending / in_progress / complete
    current = db.Column(db.Integer, nullable=False, default=0)
    target = db.Column(db.Integer, nullable=False, default=1)
    completed_at = db.Column(db.DateTime, nullable=True)
    completed_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    evidence_json = db.Column(db.Text, nullable=True)
    note = db.Column(db.String(300), nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    task = db.relationship("Task")
    completed_by = db.relationship("User", foreign_keys=[completed_by_user_id])

    @property
    def evidence(self) -> dict:
        try:
            value = json.loads(self.evidence_json or "{}")
        except Exception:
            return {}
        return value if isinstance(value, dict) else {}

    @evidence.setter
    def evidence(self, value):
        self.evidence_json = json.dumps(dict(value or {}), default=str)

    @property
    def is_complete(self) -> bool:
        return self.status == "complete"


class PhaseState(db.Model):
    __tablename__ = "onboarding_phase_states"
    __table_args__ = (
        db.UniqueConstraint("phase_id", "subject_type", "subject_id", name="uq_phase_state_subject"),
        db.Index("ix_phase_state_subject", "subject_type", "subject_id"),
    )
    id = db.Column(db.Integer, primary_key=True)

    phase_id = db.Column(db.Integer, db.ForeignKey("onboarding_phases.id"), nullable=False, index=True)
    subject_type = db.Column(db.String(20), nullable=False)
    subject_id = db.Column(db.String(60), nullable=False)
    status = db.Column(db.String(20), nullable=False, default="locked")  # locked / active / complete
    started_at = db.Column(db.DateTime, nullable=True)
    completed_at = db.Column(db.DateTime, nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    phase = db.relationship("Phase")


class Entitlement(db.Model):
    __tablename__ = "entitlements"
    __table_args__ = (db.Index("ix_entitlement_user_key", "user_id", "key"),)
    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    key = db.Column(db.String(60), nullable=False)
    source = db.Column(db.String(20), nullable=False, default="phase")  # phase / override
    source_ref = db.Column(db.String(120), nullable=True)  # e.g. "member:shop_ready"
    granted_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    granted_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    expires_at = db.Column(db.DateTime, nullable=True)
    revoked_at = db.Column(db.DateTime, nullable=True)
    revoked_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    reason = db.Column(db.String(300), nullable=True)

    user = db.relationship("User", foreign_keys=[user_id])
    granted_by = db.relationship("User", foreign_keys=[granted_by_user_id])
    revoked_by = db.relationship("User", foreign_keys=[revoked_by_user_id])

    @property
    def is_active(self) -> bool:
        if self.revoked_at is not None:
            return False
        if self.expires_at is not None and self.expires_at <= datetime.utcnow():
            return False
        return True


class TrainingModule(db.Model):
    __tablename__ = "training_modules"
    id = db.Column(db.Integer, primary_key=True)

    key = db.Column(db.String(60), nullable=False, unique=True, index=True)
    title = db.Column(db.String(200), nullable=False)
    provider = db.Column(db.String(120), nullable=True)
    url = db.Column(db.String(500), nullable=True)
    duration_min = db.Column(db.Integer, nullable=True)
    is_optional = db.Column(db.Boolean, nullable=False, default=False)
    min_score = db.Column(db.Integer, nullable=True)
    validity_days = db.Column(db.Integer, nullable=True)  # None = never expires
    description = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class TrainingCompletion(db.Model):
    __tablename__ = "training_completions"
    id = db.Column(db.Integer, primary_key=True)

    module_id = db.Column(db.Integer, db.ForeignKey("training_modules.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    score = db.Column(db.Integer, nullable=True)
    completed_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    expires_at = db.Column(db.DateTime, nullable=True)
    recorded_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    note = db.Column(db.String(300), nullable=True)

    module = db.relationship("TrainingModule")
    user = db.relationship("User", foreign_keys=[user_id])
    recorded_by = db.relationship("User", foreign_keys=[recorded_by_user_id])

    @property
    def is_valid(self) -> bool:
        return self.expires_at is None or self.expires_at > datetime.utcnow()
