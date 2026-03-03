from flask_sqlalchemy import SQLAlchemy
from datetime import date, datetime

db = SQLAlchemy()

class Member(db.Model):
    __tablename__ = "members"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(160), nullable=False, unique=True)
    member_class = db.Column(db.String(80), nullable=False)  # e.g., "Freshman", "Sophomore", "ME Junior", etc.

    # Optional NFC tag for member card
    nfc_tag = db.Column(db.String(120), nullable=True, unique=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class Item(db.Model):
    __tablename__ = "items"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(160), nullable=False)
    description = db.Column(db.String(300), nullable=True)
    category = db.Column(db.String(80), nullable=True)
    location = db.Column(db.String(120), nullable=True)
    item_condition = db.Column(db.String(120), nullable=True)
    notes = db.Column(db.String(500), nullable=True)
    photo_url = db.Column(db.String(500), nullable=True)
    item_type = db.Column(db.String(20), nullable=False, default="tool")  # tool / consumable
    is_consumable = db.Column(db.Boolean, nullable=False, default=False)
    active = db.Column(db.Boolean, nullable=False, default=True)
    min_stock_threshold = db.Column(db.Integer, nullable=False, default=0)

    total_qty = db.Column(db.Integer, nullable=False, default=0)
    available_qty = db.Column(db.Integer, nullable=False, default=0)

    # NFC tag on a bin/tool group
    nfc_tag = db.Column(db.String(120), nullable=True, unique=True)
    item_code = db.Column(db.String(40), nullable=True, unique=True, index=True)
    treasury_tracker_id = db.Column(db.String(80), nullable=True, unique=True, index=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class Transaction(db.Model):
    __tablename__ = "transactions"
    id = db.Column(db.Integer, primary_key=True)

    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False)

    action = db.Column(db.String(20), nullable=False)  # "checkout" or "return"
    qty = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), nullable=True)  # "OUT" or "RETURNED"

    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    checkout_time = db.Column(db.DateTime, nullable=True)
    return_time = db.Column(db.DateTime, nullable=True)
    due_date = db.Column(db.Date, nullable=True)
    notes = db.Column(db.String(300), nullable=True)
    checkout_notes = db.Column(db.String(300), nullable=True)
    return_condition = db.Column(db.String(120), nullable=True)
    return_notes = db.Column(db.String(300), nullable=True)
    return_photo_path = db.Column(db.String(500), nullable=True)

    member = db.relationship("Member")
    user = db.relationship("User", foreign_keys=[user_id])
    item = db.relationship("Item")


class ItemTag(db.Model):
    __tablename__ = "item_tags"
    id = db.Column(db.Integer, primary_key=True)

    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False, index=True)
    tag_value = db.Column(db.String(160), nullable=False, unique=True, index=True)
    source = db.Column(db.String(40), nullable=False, default="uid")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    item = db.relationship("Item")


class AttendanceScan(db.Model):
    __tablename__ = "attendance_scans"
    id = db.Column(db.Integer, primary_key=True)

    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False)
    scanned_uid = db.Column(db.String(120), nullable=False)
    attendance_date = db.Column(db.Date, nullable=False, default=date.today)
    scanned_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    member = db.relationship("Member")


class PrintJob(db.Model):
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


class Meeting(db.Model):
    __tablename__ = "meetings"
    id = db.Column(db.Integer, primary_key=True)

    team_name = db.Column(db.String(160), nullable=False)
    requester_email = db.Column(db.String(160), nullable=True)
    room = db.Column(db.String(80), nullable=False)
    meeting_date = db.Column(db.Date, nullable=False)
    start_time = db.Column(db.Time, nullable=False)
    end_time = db.Column(db.Time, nullable=False)
    notes = db.Column(db.String(500), nullable=True)
    # Legacy Google fields (kept for backward compatibility).
    google_event_id = db.Column(db.String(180), nullable=True)
    google_calendar_id = db.Column(db.String(240), nullable=True)
    # Active Outlook fields.
    outlook_event_id = db.Column(db.String(180), nullable=True)
    outlook_calendar_id = db.Column(db.String(240), nullable=True)
    cancel_request_token = db.Column(db.String(120), nullable=True, unique=True)
    cancel_requested_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(160), nullable=False)
    email = db.Column(db.String(160), nullable=False, unique=True, index=True)
    username = db.Column(db.String(80), nullable=True, unique=True, index=True)
    password_hash = db.Column(db.String(260), nullable=False)
    role = db.Column(db.String(30), nullable=False, default="member")
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    nfc_uid = db.Column(db.String(160), nullable=True, unique=True, index=True)
    major = db.Column(db.String(120), nullable=True)
    graduation_year = db.Column(db.Integer, nullable=True)
    exec_title = db.Column(db.String(160), nullable=True)
    exec_message = db.Column(db.String(500), nullable=True)
    headshot_url = db.Column(db.String(500), nullable=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=True, unique=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    last_login_at = db.Column(db.DateTime, nullable=True)

    member = db.relationship("Member", foreign_keys=[member_id])


class NFCTag(db.Model):
    __tablename__ = "nfc_tags"
    id = db.Column(db.Integer, primary_key=True)

    tag_uid = db.Column(db.String(160), nullable=False, unique=True, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    active = db.Column(db.Boolean, nullable=False, default=True)
    assigned_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    unassigned_at = db.Column(db.DateTime, nullable=True)
    assigned_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    notes = db.Column(db.String(260), nullable=True)

    user = db.relationship("User", foreign_keys=[user_id])
    assigned_by = db.relationship("User", foreign_keys=[assigned_by_user_id])


class Project(db.Model):
    __tablename__ = "projects"
    id = db.Column(db.Integer, primary_key=True)

    slug = db.Column(db.String(160), nullable=False, unique=True, index=True)
    title = db.Column(db.String(200), nullable=False)
    project_type = db.Column(db.String(80), nullable=True)
    summary = db.Column(db.String(320), nullable=False)
    description = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(80), nullable=False, default="Active")
    timeline = db.Column(db.Text, nullable=True)
    gallery_json = db.Column(db.Text, nullable=True)
    lead_name = db.Column(db.String(160), nullable=True)
    image_url = db.Column(db.String(500), nullable=True)
    external_link = db.Column(db.String(500), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class ContactMessage(db.Model):
    __tablename__ = "contact_messages"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(160), nullable=False)
    email = db.Column(db.String(160), nullable=False)
    kind = db.Column(db.String(40), nullable=False, default="contact")
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    target = db.Column(db.String(80), nullable=True)
    subject = db.Column(db.String(220), nullable=True)
    message = db.Column(db.Text, nullable=False)
    admin_reply = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(40), nullable=False, default="new")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    user = db.relationship("User", foreign_keys=[user_id])


class PrintRequest(db.Model):
    __tablename__ = "print_requests"
    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=True, index=True)
    printer_type = db.Column(db.String(20), nullable=False)  # H2S / P1S
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
    status = db.Column(db.String(30), nullable=False, default="submitted")
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


class Event(db.Model):
    __tablename__ = "events"
    id = db.Column(db.Integer, primary_key=True)

    title = db.Column(db.String(220), nullable=False)
    description = db.Column(db.Text, nullable=True)
    location = db.Column(db.String(220), nullable=True)
    status = db.Column(db.String(40), nullable=False, default="scheduled")
    requested_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    start_time = db.Column(db.DateTime, nullable=False)
    end_time = db.Column(db.DateTime, nullable=False)
    calendar_event_link = db.Column(db.String(500), nullable=True)
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
    checkin_time = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    event = db.relationship("Event", foreign_keys=[event_id])
    user = db.relationship("User", foreign_keys=[user_id])
    member = db.relationship("Member", foreign_keys=[member_id])


class Announcement(db.Model):
    __tablename__ = "announcements"
    id = db.Column(db.Integer, primary_key=True)

    title = db.Column(db.String(220), nullable=False)
    body = db.Column(db.Text, nullable=False)
    is_published = db.Column(db.Boolean, nullable=False, default=True)
    show_on_public = db.Column(db.Boolean, nullable=False, default=True)
    show_on_member = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    published_at = db.Column(db.DateTime, nullable=True)


class AuditLog(db.Model):
    __tablename__ = "audit_logs"
    id = db.Column(db.Integer, primary_key=True)

    admin_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    action = db.Column(db.String(160), nullable=False)
    details = db.Column(db.Text, nullable=True)
    ip_address = db.Column(db.String(120), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    admin_user = db.relationship("User", foreign_keys=[admin_user_id])


class PasswordResetToken(db.Model):
    __tablename__ = "password_reset_tokens"
    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    token = db.Column(db.String(200), nullable=False, unique=True, index=True)
    expires_at = db.Column(db.DateTime, nullable=False)
    used_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    user = db.relationship("User", foreign_keys=[user_id])
