from datetime import datetime

from asme.extensions import db


class PrintRequest(db.Model):
    """A member's intent to print something. Execution attempts live in ``PrintRun``."""

    __tablename__ = "print_requests"
    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=True, index=True)
    printer_type = db.Column(db.String(20), nullable=False)  # P1S_1..P1S_4 / H2S
    file_path = db.Column(db.String(500), nullable=True)
    file_link = db.Column(db.String(500), nullable=True)
    filament = db.Column(db.String(160), nullable=True)
    material = db.Column(db.String(120), nullable=True)
    color = db.Column(db.String(120), nullable=True)
    infill_percent = db.Column(db.Integer, nullable=True)
    priority = db.Column(db.String(40), nullable=False, default="normal")
    deadline = db.Column(db.Date, nullable=True)
    notes = db.Column(db.String(500), nullable=True)
    admin_notes = db.Column(db.String(500), nullable=True)
    status = db.Column(db.String(30), nullable=False, default="submitted", index=True)
    reviewed_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    reviewed_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    user = db.relationship("User", foreign_keys=[user_id])
    member = db.relationship("Member", foreign_keys=[member_id])
    reviewed_by = db.relationship("User", foreign_keys=[reviewed_by_user_id])
    runs = db.relationship("PrintRun", back_populates="request", order_by="PrintRun.id")


class PrintRun(db.Model):
    """One execution attempt of a request on a physical printer."""

    __tablename__ = "print_runs"
    id = db.Column(db.Integer, primary_key=True)

    request_id = db.Column(db.Integer, db.ForeignKey("print_requests.id"), nullable=False, index=True)
    printer_type = db.Column(db.String(20), nullable=False)
    status = db.Column(db.String(20), nullable=False, default="queued", index=True)  # queued/printing/done/failed
    started_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    started_at = db.Column(db.DateTime, nullable=True)
    completed_at = db.Column(db.DateTime, nullable=True)
    error = db.Column(db.String(500), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    request = db.relationship("PrintRequest", back_populates="runs")
    started_by = db.relationship("User", foreign_keys=[started_by_user_id])


class PrintJob(db.Model):
    """Legacy kiosk print queue (H2S / P1S) driven by shell commands. Retained for
    ``ASME_ENABLE_LEGACY_OPS=1`` deployments only."""

    __tablename__ = "print_jobs"
    id = db.Column(db.Integer, primary_key=True)

    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False)
    printer_type = db.Column(db.String(20), nullable=False)  # "H2S" or "P1S"
    file_name = db.Column(db.String(260), nullable=False)
    file_path = db.Column(db.String(500), nullable=False)
    notes = db.Column(db.String(500), nullable=True)

    status = db.Column(db.String(20), nullable=False, default="queued")  # queued/printing/done/failed
    submitted_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    started_at = db.Column(db.DateTime, nullable=True)
    completed_at = db.Column(db.DateTime, nullable=True)

    member = db.relationship("Member")
