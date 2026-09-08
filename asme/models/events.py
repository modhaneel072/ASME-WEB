from datetime import date, datetime

from asme.extensions import db


class Event(db.Model):
    __tablename__ = "events"
    id = db.Column(db.Integer, primary_key=True)

    title = db.Column(db.String(220), nullable=False)
    description = db.Column(db.Text, nullable=True)
    kind = db.Column(db.String(40), nullable=False, default="meeting")
    location = db.Column(db.String(220), nullable=True)
    status = db.Column(db.String(40), nullable=False, default="scheduled", index=True)
    requested_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    start_time = db.Column(db.DateTime, nullable=False, index=True)
    end_time = db.Column(db.DateTime, nullable=False)
    calendar_event_link = db.Column(db.String(500), nullable=True)
    # Legacy inline provider ids; new syncs are recorded in ``CalendarSync``.
    google_event_id = db.Column(db.String(220), nullable=True)
    google_calendar_id = db.Column(db.String(260), nullable=True)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    created_by = db.relationship("User", foreign_keys=[created_by_user_id])
    requested_by = db.relationship("User", foreign_keys=[requested_by_user_id])


class AttendanceRecord(db.Model):
    __tablename__ = "attendance_records"
    id = db.Column(db.Integer, primary_key=True)

    event_id = db.Column(db.Integer, db.ForeignKey("events.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=True, index=True)
    tag_uid = db.Column(db.String(160), nullable=True)
    checkin_method = db.Column(db.String(40), nullable=False, default="nfc")
    checkin_time = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)

    event = db.relationship("Event", foreign_keys=[event_id])
    user = db.relationship("User", foreign_keys=[user_id])
    member = db.relationship("Member", foreign_keys=[member_id])


class CalendarSync(db.Model):
    """External calendar mirror of a domain object (currently ``Event``)."""

    __tablename__ = "calendar_syncs"
    id = db.Column(db.Integer, primary_key=True)

    subject_type = db.Column(db.String(40), nullable=False, default="event")
    subject_id = db.Column(db.Integer, nullable=False, index=True)
    provider = db.Column(db.String(20), nullable=False)  # google / outlook
    calendar_id = db.Column(db.String(260), nullable=True)
    external_id = db.Column(db.String(260), nullable=True)
    link = db.Column(db.String(500), nullable=True)
    status = db.Column(db.String(20), nullable=False, default="pending")  # pending / synced / failed / deleted
    last_synced_at = db.Column(db.DateTime, nullable=True)
    last_error = db.Column(db.String(500), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class AttendanceScan(db.Model):
    """Legacy day-based kiosk scan (pre-Event attendance). Read-only going forward."""

    __tablename__ = "attendance_scans"
    id = db.Column(db.Integer, primary_key=True)

    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False)
    scanned_uid = db.Column(db.String(120), nullable=False)
    attendance_date = db.Column(db.Date, nullable=False, default=date.today)
    scanned_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    member = db.relationship("Member")


class Meeting(db.Model):
    """Legacy room booking with inline Google/Outlook ids. Used only by the legacy ops
    calendar page; the portal books through ``Event`` + ``CalendarSync``."""

    __tablename__ = "meetings"
    id = db.Column(db.Integer, primary_key=True)

    team_name = db.Column(db.String(160), nullable=False)
    requester_email = db.Column(db.String(160), nullable=True)
    room = db.Column(db.String(80), nullable=False)
    meeting_date = db.Column(db.Date, nullable=False)
    start_time = db.Column(db.Time, nullable=False)
    end_time = db.Column(db.Time, nullable=False)
    notes = db.Column(db.String(500), nullable=True)
    google_event_id = db.Column(db.String(180), nullable=True)
    google_calendar_id = db.Column(db.String(240), nullable=True)
    outlook_event_id = db.Column(db.String(180), nullable=True)
    outlook_calendar_id = db.Column(db.String(240), nullable=True)
    cancel_request_token = db.Column(db.String(120), nullable=True, unique=True)
    cancel_requested_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
