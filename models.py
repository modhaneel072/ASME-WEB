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
    category = db.Column(db.String(80), nullable=True)
    location = db.Column(db.String(120), nullable=True)

    total_qty = db.Column(db.Integer, nullable=False, default=0)
    available_qty = db.Column(db.Integer, nullable=False, default=0)

    # NFC tag on a bin/tool group
    nfc_tag = db.Column(db.String(120), nullable=True, unique=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class Transaction(db.Model):
    __tablename__ = "transactions"
    id = db.Column(db.Integer, primary_key=True)

    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False)
    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False)

    action = db.Column(db.String(20), nullable=False)  # "checkout" or "return"
    qty = db.Column(db.Integer, nullable=False)

    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    due_date = db.Column(db.Date, nullable=True)
    notes = db.Column(db.String(300), nullable=True)

    member = db.relationship("Member")
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
    google_event_id = db.Column(db.String(180), nullable=True)
    google_calendar_id = db.Column(db.String(240), nullable=True)
    cancel_request_token = db.Column(db.String(120), nullable=True, unique=True)
    cancel_requested_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
