from datetime import datetime

from asme.extensions import db


class AuditLog(db.Model):
    __tablename__ = "audit_logs"
    id = db.Column(db.Integer, primary_key=True)

    admin_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    action = db.Column(db.String(160), nullable=False, index=True)
    details = db.Column(db.Text, nullable=True)
    ip_address = db.Column(db.String(120), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)

    admin_user = db.relationship("User", foreign_keys=[admin_user_id])


class OutboxJob(db.Model):
    """Durable background work: committed in the same transaction as the domain
    write that needs it, then picked up by the in-process worker."""

    __tablename__ = "outbox_jobs"
    id = db.Column(db.Integer, primary_key=True)

    kind = db.Column(db.String(60), nullable=False, index=True)
    payload_json = db.Column(db.Text, nullable=False, default="{}")
    status = db.Column(db.String(20), nullable=False, default="pending", index=True)  # pending/running/done/failed
    attempts = db.Column(db.Integer, nullable=False, default=0)
    max_attempts = db.Column(db.Integer, nullable=False, default=5)
    run_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    locked_at = db.Column(db.DateTime, nullable=True)
    last_error = db.Column(db.String(1000), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    completed_at = db.Column(db.DateTime, nullable=True)
