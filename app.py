import os
import secrets
import smtplib
import subprocess
from datetime import date, datetime, time, timedelta
from email.message import EmailMessage
from mimetypes import guess_type
from pathlib import Path
from urllib.parse import quote_plus
from uuid import uuid4

import pandas as pd
from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from sqlalchemy import and_, inspect, or_, text
from werkzeug.utils import secure_filename

from models import AttendanceScan, Item, Meeting, Member, PrintJob, Transaction, db

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("ASME_DATABASE_URL", "sqlite:///inventory.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SECRET_KEY"] = os.environ.get("ASME_SECRET_KEY", "asme-dev-secret")
db.init_app(app)

PRINTER_TYPES = ("H2S", "P1S")
MEETING_ROOMS = ("Robotics Room", "Fluids Lab")
# Supports standard G-code plus Bambu Studio 3MF containers.
ALLOWED_GCODE_EXTENSIONS = {"gcode", "gco", "3mf"}
UPLOAD_DIR = Path(app.instance_path) / "gcode_uploads"
PRINT_COMMANDS_ENV_FILE = Path(app.instance_path) / "print_commands.env"
DEFAULT_GCAL_SERVICE_ACCOUNT_FILE = Path(app.instance_path) / "google_service_account.json"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


def load_local_print_command_env():
    if not PRINT_COMMANDS_ENV_FILE.exists():
        return

    try:
        for raw_line in PRINT_COMMANDS_ENV_FILE.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if not key:
                continue
            os.environ[key] = value
    except Exception:
        # Keep app startup resilient if env file has bad formatting.
        pass


load_local_print_command_env()


def default_due_date(days=7):
    return date.today() + timedelta(days=days)


def parse_int(value, default=1):
    try:
        parsed = int(value)
        return parsed if parsed > 0 else default
    except Exception:
        return default


def parse_due_date(value):
    raw = (value or "").strip()
    if not raw:
        return None
    try:
        year, month, day = [int(part) for part in raw.split("-")]
        return date(year, month, day)
    except Exception:
        return None


def parse_clock_time(value):
    raw = (value or "").strip()
    if not raw:
        return None
    try:
        hour, minute = [int(part) for part in raw.split(":")]
        return time(hour=hour, minute=minute)
    except Exception:
        return None


def google_calendar_embed_context():
    embed_url = (os.environ.get("ASME_GOOGLE_CALENDAR_EMBED_URL") or "").strip()
    if embed_url:
        return {"url": embed_url, "placeholder": False}

    calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_ID") or "").strip()
    if not calendar_id:
        calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_ROBOTICS_ID") or "").strip()
    if not calendar_id:
        calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_FLUIDS_ID") or "").strip()
    timezone = (os.environ.get("ASME_GOOGLE_CALENDAR_TZ") or "America/Chicago").strip()

    if calendar_id:
        generated_url = (
            "https://calendar.google.com/calendar/embed"
            f"?src={quote_plus(calendar_id)}"
            f"&ctz={quote_plus(timezone)}"
            "&mode=WEEK&showTabs=0&showPrint=0&showCalendars=0&showTz=0"
        )
        return {"url": generated_url, "placeholder": False}

    # Public fallback so the page always shows a live Google Calendar embed.
    fallback_url = (
        "https://calendar.google.com/calendar/embed"
        "?src=en.usa%23holiday%40group.v.calendar.google.com"
        "&ctz=America%2FChicago&mode=WEEK&showTabs=0&showPrint=0&showCalendars=0&showTz=0"
    )
    return {"url": fallback_url, "placeholder": True}


MEETING_SCHEMA_READY = False


def ensure_meeting_schema_columns():
    global MEETING_SCHEMA_READY
    if MEETING_SCHEMA_READY:
        return

    inspector = inspect(db.engine)
    table_names = inspector.get_table_names()
    if "meetings" not in table_names:
        db.create_all()
        inspector = inspect(db.engine)
        table_names = inspector.get_table_names()
    if "meetings" not in table_names:
        return

    existing_columns = {col["name"] for col in inspector.get_columns("meetings")}
    alters = []
    if "requester_email" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN requester_email VARCHAR(160)")
    if "google_event_id" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN google_event_id VARCHAR(180)")
    if "google_calendar_id" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN google_calendar_id VARCHAR(240)")
    if "cancel_request_token" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN cancel_request_token VARCHAR(120)")
    if "cancel_requested_at" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN cancel_requested_at DATETIME")

    if alters:
        with db.engine.begin() as conn:
            for statement in alters:
                conn.execute(text(statement))

    MEETING_SCHEMA_READY = True


def get_calendar_timezone():
    return (os.environ.get("ASME_GOOGLE_CALENDAR_TZ") or "America/Chicago").strip()


def get_google_calendar_id_for_room(room):
    robotics_calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_ROBOTICS_ID") or "").strip()
    fluids_calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_FLUIDS_ID") or "").strip()
    default_calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_ID") or "").strip()

    if room == "Robotics Room" and robotics_calendar_id:
        return robotics_calendar_id
    if room == "Fluids Lab" and fluids_calendar_id:
        return fluids_calendar_id
    return default_calendar_id


def get_google_calendar_service():
    service_account_file = (
        os.environ.get("ASME_GCAL_SERVICE_ACCOUNT_FILE") or str(DEFAULT_GCAL_SERVICE_ACCOUNT_FILE)
    ).strip()
    if not service_account_file:
        return None, "ASME_GCAL_SERVICE_ACCOUNT_FILE is not set."
    if not Path(service_account_file).exists():
        return None, f"Google service account file not found: {service_account_file}"

    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except Exception:
        return None, "Google API libraries missing. Install google-api-python-client and google-auth."

    try:
        credentials = service_account.Credentials.from_service_account_file(
            service_account_file,
            scopes=["https://www.googleapis.com/auth/calendar"],
        )
        service = build("calendar", "v3", credentials=credentials, cache_discovery=False)
        return service, None
    except Exception as exc:
        return None, f"Failed to create Google Calendar client: {str(exc)[:250]}"


def create_google_calendar_event(meeting):
    calendar_id = get_google_calendar_id_for_room(meeting.room)
    if not calendar_id:
        return None, None, (
            "Google Calendar ID is not configured. Set ASME_GOOGLE_CALENDAR_ID or room-specific IDs."
        )

    service, service_error = get_google_calendar_service()
    if service_error:
        return None, None, service_error

    timezone = get_calendar_timezone()
    start_dt = datetime.combine(meeting.meeting_date, meeting.start_time)
    end_dt = datetime.combine(meeting.meeting_date, meeting.end_time)
    description_lines = []
    if meeting.requester_email:
        description_lines.append(f"Requested by: {meeting.requester_email}")
    if meeting.notes:
        description_lines.append(f"Notes: {meeting.notes}")

    event_body = {
        "summary": f"{meeting.team_name} - {meeting.room}",
        "description": "\n".join(description_lines).strip() or None,
        "start": {"dateTime": start_dt.isoformat(), "timeZone": timezone},
        "end": {"dateTime": end_dt.isoformat(), "timeZone": timezone},
    }

    try:
        created_event = service.events().insert(calendarId=calendar_id, body=event_body).execute()
        return calendar_id, created_event.get("id"), None
    except Exception as exc:
        return None, None, f"Google Calendar event create failed: {str(exc)[:250]}"


def delete_google_calendar_event(meeting):
    if not meeting.google_event_id:
        return None

    calendar_id = (meeting.google_calendar_id or "").strip() or get_google_calendar_id_for_room(meeting.room)
    if not calendar_id:
        return "Missing Google Calendar ID for cancellation."

    service, service_error = get_google_calendar_service()
    if service_error:
        return service_error

    try:
        service.events().delete(calendarId=calendar_id, eventId=meeting.google_event_id).execute()
        return None
    except Exception as exc:
        status_code = getattr(getattr(exc, "resp", None), "status", None)
        if status_code == 404:
            return None
        return f"Google Calendar event delete failed: {str(exc)[:250]}"


def send_meeting_cancel_confirmation_email(meeting, confirm_url, reject_url):
    smtp_host = (os.environ.get("ASME_SMTP_HOST") or "smtp.gmail.com").strip()
    smtp_port = int((os.environ.get("ASME_SMTP_PORT") or "587").strip())
    smtp_user = (os.environ.get("ASME_SMTP_USER") or "").strip()
    smtp_pass = (os.environ.get("ASME_SMTP_PASS") or "").strip()
    notify_to = (os.environ.get("ASME_CANCEL_NOTIFY_TO") or smtp_user).strip()

    if not smtp_user or not smtp_pass:
        return "SMTP is not configured. Set ASME_SMTP_USER and ASME_SMTP_PASS."
    if not notify_to:
        return "ASME_CANCEL_NOTIFY_TO is not configured."

    message = EmailMessage()
    message["From"] = smtp_user
    message["To"] = notify_to
    message["Subject"] = f"ASME Meeting Cancellation Request: {meeting.team_name} ({meeting.room})"
    message.set_content(
        "\n".join(
            [
                "A meeting cancellation was requested from the website.",
                "",
                f"Team: {meeting.team_name}",
                f"Room: {meeting.room}",
                f"Date: {meeting.meeting_date.isoformat()}",
                f"Time: {meeting.start_time.strftime('%H:%M')} - {meeting.end_time.strftime('%H:%M')}",
                f"Requested by: {meeting.requester_email or 'Not provided'}",
                f"Notes: {meeting.notes or ''}",
                "",
                f"Confirm cancellation: {confirm_url}",
                f"Reject cancellation:  {reject_url}",
            ]
        )
    )

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as smtp:
            smtp.starttls()
            smtp.login(smtp_user, smtp_pass)
            smtp.send_message(message)
        return None
    except Exception as exc:
        return f"Failed to send cancellation email: {str(exc)[:250]}"


def calendar_automation_status():
    default_calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_ID") or "").strip()
    robotics_calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_ROBOTICS_ID") or "").strip()
    fluids_calendar_id = (os.environ.get("ASME_GOOGLE_CALENDAR_FLUIDS_ID") or "").strip()
    service_account_file = (
        os.environ.get("ASME_GCAL_SERVICE_ACCOUNT_FILE") or str(DEFAULT_GCAL_SERVICE_ACCOUNT_FILE)
    ).strip()
    smtp_user = (os.environ.get("ASME_SMTP_USER") or "").strip()
    smtp_pass = (os.environ.get("ASME_SMTP_PASS") or "").strip()
    cancel_notify_to = (os.environ.get("ASME_CANCEL_NOTIFY_TO") or "").strip()

    has_any_calendar_id = bool(default_calendar_id or robotics_calendar_id or fluids_calendar_id)
    return {
        "service_account_file": service_account_file,
        "service_account_file_exists": Path(service_account_file).exists(),
        "default_calendar_id": default_calendar_id,
        "robotics_calendar_id": robotics_calendar_id,
        "fluids_calendar_id": fluids_calendar_id,
        "has_any_calendar_id": has_any_calendar_id,
        "smtp_ready": bool(smtp_user and smtp_pass),
        "smtp_user": smtp_user,
        "cancel_notify_to": cancel_notify_to,
        "timezone": get_calendar_timezone(),
    }


def allowed_gcode(filename):
    if "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_GCODE_EXTENSIONS


def delete_print_job_with_file(job):
    file_path = job.file_path
    file_name = job.file_name
    printer_type = job.printer_type
    file_removed = False
    file_error = None

    if file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
            file_removed = True
        except Exception as exc:
            file_error = str(exc)

    db.session.delete(job)
    db.session.commit()
    dispatch_next_job(printer_type)

    return {
        "file_name": file_name,
        "file_removed": file_removed,
        "file_error": file_error,
    }


def resolve_member(member_tag, member_id):
    member_tag = (member_tag or "").strip()
    member_id = (member_id or "").strip()

    if member_tag:
        member = Member.query.filter_by(nfc_tag=member_tag).first()
        if member:
            return member

    if member_id:
        try:
            return db.session.get(Member, int(member_id))
        except Exception:
            return None
    return None


def resolve_item(item_tag, item_id):
    item_tag = (item_tag or "").strip()
    item_id = (item_id or "").strip()

    if item_tag:
        item = Item.query.filter_by(nfc_tag=item_tag).first()
        if item:
            return item

    if item_id:
        try:
            return db.session.get(Item, int(item_id))
        except Exception:
            return None
    return None


def dispatch_next_job(printer_type):
    active = PrintJob.query.filter_by(printer_type=printer_type, status="printing").first()
    if active:
        return None

    next_job = (
        PrintJob.query.filter_by(printer_type=printer_type, status="queued")
        .order_by(PrintJob.submitted_at.asc(), PrintJob.id.asc())
        .first()
    )
    if not next_job:
        return None

    next_job.status = "printing"
    next_job.started_at = datetime.utcnow()
    db.session.commit()

    dispatch_error = launch_print_command(next_job)
    if dispatch_error:
        next_job.status = "failed"
        next_job.completed_at = datetime.utcnow()
        next_job.notes = f"{next_job.notes} | {dispatch_error}" if next_job.notes else dispatch_error
        db.session.commit()
        return dispatch_next_job(printer_type)

    return next_job


def launch_print_command(job):
    env_name = f"ASME_{job.printer_type}_PRINT_CMD"
    cmd_template = os.environ.get(env_name)
    if not cmd_template:
        return (
            f"{env_name} is not configured. Add it to {PRINT_COMMANDS_ENV_FILE} "
            "and restart the app."
        )

    command = cmd_template.format(file=job.file_path, filename=job.file_name, job_id=job.id)
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    if result.returncode == 0:
        return None

    error_text = (result.stderr or result.stdout or "print command failed").strip()
    return f"{env_name} failed: {error_text[:300]}"


def iso_or_none(value):
    if value is None:
        return None
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def serialize_member(member):
    return {
        "id": member.id,
        "name": member.name,
        "email": member.email,
        "member_class": member.member_class,
        "nfc_tag": member.nfc_tag,
        "created_at": iso_or_none(member.created_at),
    }


def serialize_item(item):
    return {
        "id": item.id,
        "name": item.name,
        "category": item.category,
        "location": item.location,
        "total_qty": item.total_qty,
        "available_qty": item.available_qty,
        "nfc_tag": item.nfc_tag,
        "low_stock": item.available_qty <= 2,
        "out_of_stock": item.available_qty <= 0,
        "created_at": iso_or_none(item.created_at),
    }


def serialize_transaction(tx):
    return {
        "id": tx.id,
        "timestamp": iso_or_none(tx.timestamp),
        "member_id": tx.member_id,
        "member_name": tx.member.name,
        "item_id": tx.item_id,
        "item_name": tx.item.name,
        "action": tx.action,
        "qty": tx.qty,
        "due_date": iso_or_none(tx.due_date),
        "notes": tx.notes,
    }


def serialize_attendance_scan(scan):
    return {
        "id": scan.id,
        "member_id": scan.member_id,
        "member_name": scan.member.name,
        "uid": scan.scanned_uid,
        "attendance_date": iso_or_none(scan.attendance_date),
        "scanned_at": iso_or_none(scan.scanned_at),
    }


def serialize_print_job(job):
    return {
        "id": job.id,
        "member_id": job.member_id,
        "member_name": job.member.name,
        "printer_type": job.printer_type,
        "file_name": job.file_name,
        "status": job.status,
        "notes": job.notes,
        "submitted_at": iso_or_none(job.submitted_at),
        "started_at": iso_or_none(job.started_at),
        "completed_at": iso_or_none(job.completed_at),
        "open_url": url_for("open_print_job", job_id=job.id),
        "download_url": url_for("download_print_job", job_id=job.id),
    }


def get_today_attendance_unique():
    scans = (
        AttendanceScan.query.filter_by(attendance_date=date.today())
        .order_by(AttendanceScan.scanned_at.desc())
        .all()
    )
    unique_scans = []
    seen = set()
    for scan in scans:
        if scan.member_id in seen:
            continue
        seen.add(scan.member_id)
        unique_scans.append(scan)
    return unique_scans


def get_queue_snapshot():
    payload = {}
    for printer in PRINTER_TYPES:
        active = (
            PrintJob.query.filter_by(printer_type=printer, status="printing")
            .order_by(PrintJob.started_at.asc(), PrintJob.id.asc())
            .first()
        )
        queued = (
            PrintJob.query.filter_by(printer_type=printer, status="queued")
            .order_by(PrintJob.submitted_at.asc(), PrintJob.id.asc())
            .all()
        )
        finished = (
            PrintJob.query.filter(
                PrintJob.printer_type == printer,
                PrintJob.status.in_(["done", "failed"]),
            )
            .order_by(PrintJob.completed_at.desc(), PrintJob.id.desc())
            .limit(8)
            .all()
        )
        payload[printer] = {
            "active": serialize_print_job(active) if active else None,
            "queued": [serialize_print_job(job) for job in queued],
            "recent_finished": [serialize_print_job(job) for job in finished],
        }
    return payload


def build_bootstrap_payload():
    members = Member.query.order_by(Member.name.asc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    attendance = get_today_attendance_unique()
    recent_transactions = Transaction.query.order_by(Transaction.timestamp.desc()).limit(15).all()
    queues = get_queue_snapshot()

    return {
        "today": str(date.today()),
        "default_due": str(default_due_date()),
        "members": [serialize_member(member) for member in members],
        "items": [serialize_item(item) for item in items],
        "attendance_today": [serialize_attendance_scan(scan) for scan in attendance],
        "attendance_count": len(attendance),
        "recent_transactions": [serialize_transaction(tx) for tx in recent_transactions],
        "queues": queues,
    }


def value_from_request(key, default=None):
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        return payload.get(key, default)
    return request.form.get(key, default)


def api_error(message, status=400):
    return jsonify({"ok": False, "error": message}), status


def api_success(message=None, status=200):
    return jsonify({"ok": True, "message": message, "payload": build_bootstrap_payload()}), status


def dashboard_context(transaction_limit=15):
    members = Member.query.order_by(Member.name.asc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    recent_transactions = (
        Transaction.query.order_by(Transaction.timestamp.desc())
        .limit(transaction_limit)
        .all()
    )
    today_scans = (
        AttendanceScan.query.filter_by(attendance_date=date.today())
        .order_by(AttendanceScan.scanned_at.desc())
        .all()
    )

    today_attendance = []
    seen_member_ids = set()
    for scan in today_scans:
        if scan.member_id in seen_member_ids:
            continue
        seen_member_ids.add(scan.member_id)
        today_attendance.append(scan)

    queues = {}
    for printer in PRINTER_TYPES:
        active = (
            PrintJob.query.filter_by(printer_type=printer, status="printing")
            .order_by(PrintJob.started_at.asc(), PrintJob.id.asc())
            .first()
        )
        queued = (
            PrintJob.query.filter_by(printer_type=printer, status="queued")
            .order_by(PrintJob.submitted_at.asc(), PrintJob.id.asc())
            .all()
        )
        recent_finished = (
            PrintJob.query.filter(
                PrintJob.printer_type == printer,
                PrintJob.status.in_(["done", "failed"]),
            )
            .order_by(PrintJob.completed_at.desc(), PrintJob.id.desc())
            .limit(8)
            .all()
        )
        queues[printer] = {
            "active": active,
            "queued": queued,
            "recent_finished": recent_finished,
        }

    low_stock_count = len([item for item in items if item.available_qty <= 2])
    active_prints_count = len([printer for printer in PRINTER_TYPES if queues[printer]["active"]])

    return {
        "members": members,
        "items": items,
        "recent_transactions": recent_transactions,
        "today_attendance": today_attendance,
        "attendance_count": len(today_attendance),
        "queues": queues,
        "default_due": str(default_due_date()),
        "today": str(date.today()),
        "low_stock_count": low_stock_count,
        "active_prints_count": active_prints_count,
        "h2s_waiting_count": len(queues["H2S"]["queued"]),
        "p1s_waiting_count": len(queues["P1S"]["queued"]),
    }


def get_upcoming_meetings():
    ensure_meeting_schema_columns()
    now = datetime.now()
    return (
        Meeting.query.filter(
            or_(
                Meeting.meeting_date > now.date(),
                and_(Meeting.meeting_date == now.date(), Meeting.end_time >= now.time()),
            )
        )
        .filter(Meeting.cancel_request_token.is_(None))
        .order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc())
        .all()
    )


def render_ops_page(template_name, active_page, page_title, page_subtitle, transaction_limit=15):
    h2s_print_cmd = (os.environ.get("ASME_H2S_PRINT_CMD") or "").strip()
    p1s_print_cmd = (os.environ.get("ASME_P1S_PRINT_CMD") or "").strip()

    context = dashboard_context(transaction_limit=transaction_limit)
    context.update(
        {
            "active_page": active_page,
            "page_title": page_title,
            "page_subtitle": page_subtitle,
            "h2s_print_cmd_configured": bool(h2s_print_cmd),
            "p1s_print_cmd_configured": bool(p1s_print_cmd),
            "h2s_print_cmd_value": h2s_print_cmd,
            "p1s_print_cmd_value": p1s_print_cmd,
            "print_commands_env_file": str(PRINT_COMMANDS_ENV_FILE),
        }
    )
    return render_template(template_name, **context)


def redirect_home(page):
    next_url = (request.form.get("next") or "").strip()
    if next_url.startswith("/"):
        return redirect(next_url)

    page_endpoints = {
        "dashboard": "dashboard_page",
        "attendance": "attendance_page",
        "inventory": "inventory_page",
        "prints": "prints_page",
        "activity": "activity_page",
        "calendar": "calendar_page",
    }
    endpoint = page_endpoints.get(page, "dashboard_page")
    return redirect(url_for(endpoint))


@app.get("/")
def index():
    return render_ops_page(
        template_name="ops/dashboard.html",
        active_page="dashboard",
        page_title="Dashboard",
        page_subtitle="Overview of attendance, stock, and printer queue status.",
        transaction_limit=12,
    )


@app.get("/dashboard")
def dashboard_page():
    return render_ops_page(
        template_name="ops/dashboard.html",
        active_page="dashboard",
        page_title="Dashboard",
        page_subtitle="Overview of attendance, stock, and printer queue status.",
        transaction_limit=12,
    )


@app.get("/attendance")
def attendance_page():
    return render_ops_page(
        template_name="ops/attendance.html",
        active_page="attendance",
        page_title="Attendance",
        page_subtitle="Scan member NFC UIDs and track who is present today.",
        transaction_limit=10,
    )


@app.get("/inventory")
def inventory_page():
    return render_ops_page(
        template_name="ops/inventory.html",
        active_page="inventory",
        page_title="Inventory",
        page_subtitle="Checkout, return, and monitor stock health across all items.",
        transaction_limit=12,
    )


@app.get("/prints")
def prints_page():
    return render_ops_page(
        template_name="ops/prints.html",
        active_page="prints",
        page_title="Printing",
        page_subtitle="Submit jobs and run separate H2S and P1S print queues.",
        transaction_limit=12,
    )


@app.get("/activity")
def activity_page():
    return render_ops_page(
        template_name="ops/activity.html",
        active_page="activity",
        page_title="Activity Feed",
        page_subtitle="Recent checkout and return history.",
        transaction_limit=80,
    )


@app.get("/calendar")
def calendar_page():
    ensure_meeting_schema_columns()
    upcoming_meetings = get_upcoming_meetings()
    meetings_by_room = {room: [] for room in MEETING_ROOMS}
    for meeting in upcoming_meetings:
        meetings_by_room.setdefault(meeting.room, []).append(meeting)
    google_calendar = google_calendar_embed_context()
    automation_status = calendar_automation_status()
    pending_cancellations = (
        Meeting.query.filter(Meeting.cancel_request_token.isnot(None))
        .order_by(Meeting.cancel_requested_at.desc(), Meeting.id.desc())
        .all()
    )

    context = dashboard_context(transaction_limit=20)
    context.update(
        {
            "active_page": "calendar",
            "page_title": "Calendar",
            "page_subtitle": "Schedule team meetings in the Robotics Room or Fluids Lab.",
            "meeting_rooms": MEETING_ROOMS,
            "meetings_by_room": meetings_by_room,
            "calendar_default_date": str(date.today()),
            "google_calendar_embed_url": google_calendar["url"],
            "google_calendar_placeholder": google_calendar["placeholder"],
            "calendar_automation_status": automation_status,
            "pending_cancellations": pending_cancellations,
        }
    )
    return render_template("ops/calendar.html", **context)


@app.post("/calendar/book")
def book_meeting():
    ensure_meeting_schema_columns()
    team_name = (request.form.get("team_name") or "").strip()
    requester_email = (request.form.get("requester_email") or "").strip() or None
    room = (request.form.get("room") or "").strip()
    meeting_date = parse_due_date(request.form.get("meeting_date"))
    start_time = parse_clock_time(request.form.get("start_time"))
    end_time = parse_clock_time(request.form.get("end_time"))
    notes = (request.form.get("notes") or "").strip() or None

    if not team_name:
        flash("Team name is required.", "error")
        return redirect_home("calendar")
    if room not in MEETING_ROOMS:
        flash("Please select Robotics Room or Fluids Lab.", "error")
        return redirect_home("calendar")
    if not meeting_date:
        flash("Meeting date is required.", "error")
        return redirect_home("calendar")
    if meeting_date < date.today():
        flash("Meeting date cannot be in the past.", "error")
        return redirect_home("calendar")
    if not start_time or not end_time:
        flash("Start and end times are required.", "error")
        return redirect_home("calendar")
    if end_time <= start_time:
        flash("End time must be after start time.", "error")
        return redirect_home("calendar")

    conflicting = (
        Meeting.query.filter(
            Meeting.room == room,
            Meeting.meeting_date == meeting_date,
            Meeting.start_time < end_time,
            Meeting.end_time > start_time,
            Meeting.cancel_request_token.is_(None),
        )
        .order_by(Meeting.start_time.asc(), Meeting.id.asc())
        .first()
    )
    if conflicting:
        flash(
            f"{room} is already booked by {conflicting.team_name} from "
            f"{conflicting.start_time.strftime('%H:%M')} to {conflicting.end_time.strftime('%H:%M')}.",
            "error",
        )
        return redirect_home("calendar")

    meeting = Meeting(
        team_name=team_name,
        requester_email=requester_email,
        room=room,
        meeting_date=meeting_date,
        start_time=start_time,
        end_time=end_time,
        notes=notes,
    )

    calendar_id, event_id, calendar_error = create_google_calendar_event(meeting)
    if calendar_error:
        flash(f"Meeting was not saved: {calendar_error}", "error")
        return redirect_home("calendar")

    meeting.google_calendar_id = calendar_id
    meeting.google_event_id = event_id

    db.session.add(meeting)
    db.session.commit()
    flash(
        f"Booked {room} for {team_name} on {meeting_date.isoformat()} "
        f"({start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}) and synced to Google Calendar.",
        "success",
    )
    return redirect_home("calendar")


@app.post("/calendar/meeting/<int:meeting_id>/cancel")
def request_meeting_cancel(meeting_id):
    ensure_meeting_schema_columns()
    meeting = db.session.get(Meeting, meeting_id)
    if not meeting:
        flash("Meeting not found.", "error")
        return redirect_home("calendar")
    if meeting.cancel_request_token:
        flash("Cancellation already requested and waiting for email confirmation.", "info")
        return redirect_home("calendar")

    token = secrets.token_urlsafe(32)
    meeting.cancel_request_token = token
    meeting.cancel_requested_at = datetime.utcnow()
    db.session.commit()

    confirm_url = url_for("confirm_meeting_cancel", token=token, _external=True)
    reject_url = url_for("reject_meeting_cancel", token=token, _external=True)
    email_error = send_meeting_cancel_confirmation_email(meeting, confirm_url=confirm_url, reject_url=reject_url)
    if email_error:
        meeting.cancel_request_token = None
        meeting.cancel_requested_at = None
        db.session.commit()
        flash(f"Cancellation email could not be sent: {email_error}", "error")
        return redirect_home("calendar")

    flash(
        "Cancellation request sent. An approval email was sent to the club Gmail for confirmation.",
        "info",
    )
    return redirect_home("calendar")


@app.get("/calendar/cancel/confirm/<token>")
def confirm_meeting_cancel(token):
    ensure_meeting_schema_columns()
    meeting = Meeting.query.filter_by(cancel_request_token=token).first()
    if not meeting:
        return (
            "<h3>Cancellation link is invalid or already used.</h3>"
            "<p>You can close this tab.</p>",
            404,
        )

    calendar_error = delete_google_calendar_event(meeting)
    if calendar_error:
        return (
            "<h3>Cancellation could not be completed.</h3>"
            f"<p>{calendar_error}</p>"
            "<p>Please fix configuration and retry.</p>",
            500,
        )

    team_name = meeting.team_name
    room = meeting.room
    meeting_date = meeting.meeting_date.isoformat()
    db.session.delete(meeting)
    db.session.commit()

    return (
        "<h3>Cancellation confirmed.</h3>"
        f"<p>{team_name} in {room} on {meeting_date} was removed from Google Calendar.</p>"
        "<p>You can close this tab.</p>"
    )


@app.get("/calendar/cancel/reject/<token>")
def reject_meeting_cancel(token):
    ensure_meeting_schema_columns()
    meeting = Meeting.query.filter_by(cancel_request_token=token).first()
    if not meeting:
        return (
            "<h3>Rejection link is invalid or already used.</h3>"
            "<p>You can close this tab.</p>",
            404,
        )

    meeting.cancel_request_token = None
    meeting.cancel_requested_at = None
    db.session.commit()
    return (
        "<h3>Cancellation request rejected.</h3>"
        "<p>The meeting remains on the schedule and in Google Calendar.</p>"
        "<p>You can close this tab.</p>"
    )


@app.get("/api/bootstrap")
def api_bootstrap():
    return jsonify({"ok": True, "payload": build_bootstrap_payload()})


@app.post("/api/attendance/scan")
def api_attendance_scan():
    uid = str(value_from_request("uid", "")).strip()
    if not uid:
        return api_error("Scan failed: UID was empty.")

    member = Member.query.filter_by(nfc_tag=uid).first()
    if not member:
        return api_error("UID not recognized. Pair this UID to a member first.", status=404)

    scan = AttendanceScan(member_id=member.id, scanned_uid=uid, attendance_date=date.today())
    db.session.add(scan)
    db.session.commit()

    scans_today = AttendanceScan.query.filter_by(member_id=member.id, attendance_date=date.today()).count()
    if scans_today == 1:
        message = f"Attendance marked for {member.name}."
    else:
        message = f"{member.name} scanned again. Attendance already marked for today."
    return api_success(message=message)


@app.post("/api/inventory/transact")
def api_inventory_transact():
    member = resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
    item = resolve_item(value_from_request("item_tag"), value_from_request("item_id"))
    action = str(value_from_request("action", "")).strip().lower()
    qty = parse_int(value_from_request("qty"), default=1)
    notes = (str(value_from_request("notes", "")).strip() or None)
    due = parse_due_date(value_from_request("due_date"))

    if not member:
        return api_error("Could not find member. Scan a member UID or select one.")
    if not item:
        return api_error("Could not find item. Scan an item UID or select one.")
    if action not in {"checkout", "return"}:
        return api_error("Invalid inventory action.")

    if action == "checkout":
        if item.available_qty < qty:
            return api_error(f"Not enough stock. {item.name} has {item.available_qty} available.")
        item.available_qty -= qty
    else:
        item.available_qty = min(item.total_qty, item.available_qty + qty)

    tx = Transaction(
        member_id=member.id,
        item_id=item.id,
        action=action,
        qty=qty,
        due_date=due if action == "checkout" else None,
        notes=notes,
    )
    db.session.add(tx)
    db.session.commit()

    return api_success(message=f"{action.title()} saved: {qty} x {item.name} for {member.name}.")


@app.post("/api/print/submit")
def api_print_submit():
    member = resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
    printer_type = str(value_from_request("printer_type", "")).strip().upper()
    notes = (str(value_from_request("notes", "")).strip() or None)
    file = request.files.get("gcode_file")

    if not member:
        return api_error("Could not find member for print job. Scan/select member first.")
    if printer_type not in PRINTER_TYPES:
        return api_error("Invalid printer queue selected.")
    if not file or not file.filename:
        return api_error("Please upload a print file.")
    if not allowed_gcode(file.filename):
        return api_error("Invalid file type. Use .gcode, .gco, or .3mf.")

    original_name = secure_filename(file.filename)
    stored_name = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{original_name}"
    file_path = UPLOAD_DIR / stored_name
    file.save(file_path)

    job = PrintJob(
        member_id=member.id,
        printer_type=printer_type,
        file_name=original_name,
        file_path=str(file_path),
        notes=notes,
        status="queued",
    )
    db.session.add(job)
    db.session.commit()

    started = dispatch_next_job(printer_type)
    if started and started.id == job.id:
        message = f"{printer_type} job submitted and auto-started."
    else:
        message = f"{printer_type} job submitted to queue."
    return api_success(message=message, status=201)


@app.post("/api/print/job/<int:job_id>/complete")
def api_complete_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        return api_error("Print job not found.", status=404)

    job.status = "done"
    job.completed_at = datetime.utcnow()
    db.session.commit()

    dispatch_next_job(job.printer_type)
    return api_success(message=f"Marked job #{job.id} done. Next {job.printer_type} job auto-started if available.")


@app.post("/api/print/job/<int:job_id>/fail")
def api_fail_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        return api_error("Print job not found.", status=404)

    job.status = "failed"
    job.completed_at = datetime.utcnow()
    db.session.commit()

    dispatch_next_job(job.printer_type)
    return api_success(message=f"Marked job #{job.id} failed. Next {job.printer_type} job auto-started if available.")


@app.post("/api/print/job/<int:job_id>/delete")
def api_delete_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        return api_error("Print job not found.", status=404)
    if job.status == "printing":
        return api_error("Cannot delete an active printing job. Mark it done or failed first.", status=409)

    result = delete_print_job_with_file(job)
    if result["file_error"]:
        message = f"Deleted job for {result['file_name']}, but file removal failed: {result['file_error'][:200]}"
    elif result["file_removed"]:
        message = f"Deleted job and removed {result['file_name']} from storage."
    else:
        message = f"Deleted job for {result['file_name']}. File was already missing."
    return api_success(message=message)


@app.post("/attendance/scan")
def attendance_scan():
    uid = (request.form.get("uid") or "").strip()
    if not uid:
        flash("Scan failed: UID was empty.", "error")
        return redirect_home("attendance")

    member = Member.query.filter_by(nfc_tag=uid).first()
    if not member:
        flash("UID not recognized. Pair this UID to a member first.", "error")
        return redirect_home("attendance")

    scan = AttendanceScan(member_id=member.id, scanned_uid=uid, attendance_date=date.today())
    db.session.add(scan)
    db.session.commit()

    today_count = AttendanceScan.query.filter_by(member_id=member.id, attendance_date=date.today()).count()
    if today_count == 1:
        flash(f"Attendance marked for {member.name}.", "success")
    else:
        flash(f"{member.name} scanned again. Attendance already marked for today.", "info")
    return redirect_home("attendance")


@app.post("/transact")
def transact():
    member = resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    item = resolve_item(request.form.get("item_tag"), request.form.get("item_id"))
    action = (request.form.get("action") or "").strip().lower()
    qty = parse_int(request.form.get("qty"), default=1)
    notes = (request.form.get("notes") or "").strip() or None
    due = parse_due_date(request.form.get("due_date"))

    if not member:
        flash("Could not find member. Scan a member UID or select one.", "error")
        return redirect_home("inventory")
    if not item:
        flash("Could not find item. Scan an item UID or select one.", "error")
        return redirect_home("inventory")
    if action not in {"checkout", "return"}:
        flash("Invalid inventory action.", "error")
        return redirect_home("inventory")

    if action == "checkout":
        if item.available_qty < qty:
            flash(f"Not enough stock. {item.name} has {item.available_qty} available.", "error")
            return redirect_home("inventory")
        item.available_qty -= qty
    else:
        item.available_qty = min(item.total_qty, item.available_qty + qty)

    tx = Transaction(
        member_id=member.id,
        item_id=item.id,
        action=action,
        qty=qty,
        due_date=due if action == "checkout" else None,
        notes=notes,
    )
    db.session.add(tx)
    db.session.commit()
    flash(f"{action.title()} saved: {qty} x {item.name} for {member.name}.", "success")
    return redirect_home("inventory")


@app.post("/print/submit")
def submit_print_job():
    member = resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    printer_type = (request.form.get("printer_type") or "").strip().upper()
    notes = (request.form.get("notes") or "").strip() or None
    file = request.files.get("gcode_file")

    if not member:
        flash("Could not find member for print job. Scan/select member first.", "error")
        return redirect_home("prints")
    if printer_type not in PRINTER_TYPES:
        flash("Invalid printer queue selected.", "error")
        return redirect_home("prints")
    if not file or not file.filename:
        flash("Please upload a print file.", "error")
        return redirect_home("prints")
    if not allowed_gcode(file.filename):
        flash("Invalid file type. Use .gcode, .gco, or .3mf.", "error")
        return redirect_home("prints")

    original_name = secure_filename(file.filename)
    stored_name = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{original_name}"
    file_path = UPLOAD_DIR / stored_name
    file.save(file_path)

    job = PrintJob(
        member_id=member.id,
        printer_type=printer_type,
        file_name=original_name,
        file_path=str(file_path),
        notes=notes,
        status="queued",
    )
    db.session.add(job)
    db.session.commit()

    started = dispatch_next_job(printer_type)
    if started and started.id == job.id:
        flash(f"{printer_type} job submitted and auto-started.", "success")
    else:
        flash(f"{printer_type} job submitted to queue.", "success")
    return redirect_home("prints")


@app.post("/print/job/<int:job_id>/complete")
def complete_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
        return redirect_home("prints")

    job.status = "done"
    job.completed_at = datetime.utcnow()
    db.session.commit()

    dispatch_next_job(job.printer_type)
    flash(f"Marked job #{job.id} done. Next {job.printer_type} job auto-started if available.", "success")
    return redirect_home("prints")


@app.post("/print/job/<int:job_id>/fail")
def fail_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
        return redirect_home("prints")

    job.status = "failed"
    job.completed_at = datetime.utcnow()
    db.session.commit()

    dispatch_next_job(job.printer_type)
    flash(f"Marked job #{job.id} failed. Next {job.printer_type} job auto-started if available.", "info")
    return redirect_home("prints")


@app.get("/print/job/<int:job_id>/download")
def download_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job or not os.path.exists(job.file_path):
        flash("Print file not found for that job.", "error")
        return redirect_home("prints")
    return send_file(job.file_path, as_attachment=True, download_name=job.file_name)


@app.get("/print/job/<int:job_id>/open")
def open_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job or not os.path.exists(job.file_path):
        flash("Print file not found for that job.", "error")
        return redirect_home("prints")

    mime_type, _ = guess_type(job.file_name)
    return send_file(
        job.file_path,
        as_attachment=False,
        download_name=job.file_name,
        mimetype=mime_type or "application/octet-stream",
    )


@app.post("/print/job/<int:job_id>/delete")
def delete_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
        return redirect_home("prints")
    if job.status == "printing":
        flash("Cannot delete an active printing job. Mark it done or failed first.", "error")
        return redirect_home("prints")

    result = delete_print_job_with_file(job)
    if result["file_error"]:
        flash(f"Deleted job, but file removal failed: {result['file_error'][:200]}", "error")
    elif result["file_removed"]:
        flash(f"Deleted print job and removed {result['file_name']}.", "info")
    else:
        flash(f"Deleted print job for {result['file_name']}. File was already missing.", "info")
    return redirect_home("prints")


@app.route("/pair/member", methods=["GET", "POST"])
def pair_member():
    if request.method == "POST":
        member_id_raw = (request.form.get("member_id") or "").strip()
        tag = (request.form.get("tag") or "").strip()

        if not member_id_raw or not tag:
            flash("Member and UID are required.", "error")
            return redirect(url_for("pair_member"))

        try:
            member_id = int(member_id_raw)
        except Exception:
            flash("Invalid member selection.", "error")
            return redirect(url_for("pair_member"))

        existing = Member.query.filter_by(nfc_tag=tag).first()
        if existing and existing.id != member_id:
            flash("That UID is already assigned to another member.", "error")
            return redirect(url_for("pair_member"))

        member = db.session.get(Member, member_id)
        if not member:
            flash("Member not found.", "error")
            return redirect(url_for("pair_member"))

        member.nfc_tag = tag
        db.session.commit()
        flash(f"Saved UID for {member.name}.", "success")
        return redirect(url_for("pair_member"))

    members = Member.query.order_by(Member.name.asc()).all()
    paired = Member.query.filter(Member.nfc_tag.isnot(None)).order_by(Member.name.asc()).all()
    return render_template("pair_member.html", members=members, paired=paired)


@app.route("/pair/item", methods=["GET", "POST"])
def pair_item():
    if request.method == "POST":
        item_id_raw = (request.form.get("item_id") or "").strip()
        tag = (request.form.get("tag") or "").strip()

        if not item_id_raw or not tag:
            flash("Item and UID are required.", "error")
            return redirect(url_for("pair_item"))

        try:
            item_id = int(item_id_raw)
        except Exception:
            flash("Invalid item selection.", "error")
            return redirect(url_for("pair_item"))

        existing = Item.query.filter_by(nfc_tag=tag).first()
        if existing and existing.id != item_id:
            flash("That UID is already assigned to another item.", "error")
            return redirect(url_for("pair_item"))

        item = db.session.get(Item, item_id)
        if not item:
            flash("Item not found.", "error")
            return redirect(url_for("pair_item"))

        item.nfc_tag = tag
        db.session.commit()
        flash(f"Saved UID for {item.name}.", "success")
        return redirect(url_for("pair_item"))

    items = Item.query.order_by(Item.name.asc()).all()
    paired = Item.query.filter(Item.nfc_tag.isnot(None)).order_by(Item.name.asc()).all()
    return render_template("pair_item.html", items=items, paired=paired)


@app.get("/export")
def export_excel():
    ensure_meeting_schema_columns()
    members = Member.query.order_by(Member.id.asc()).all()
    items = Item.query.order_by(Item.id.asc()).all()
    transactions = Transaction.query.order_by(Transaction.timestamp.desc()).all()
    scans = AttendanceScan.query.order_by(AttendanceScan.scanned_at.desc()).all()
    jobs = PrintJob.query.order_by(PrintJob.submitted_at.desc()).all()
    meetings = (
        Meeting.query.order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc()).all()
    )

    members_df = pd.DataFrame(
        [
            {
                "id": m.id,
                "name": m.name,
                "email": m.email,
                "class": m.member_class,
                "nfc_tag": m.nfc_tag,
                "created_at": m.created_at,
            }
            for m in members
        ]
    )
    items_df = pd.DataFrame(
        [
            {
                "id": i.id,
                "name": i.name,
                "category": i.category,
                "location": i.location,
                "total_qty": i.total_qty,
                "available_qty": i.available_qty,
                "nfc_tag": i.nfc_tag,
                "created_at": i.created_at,
            }
            for i in items
        ]
    )
    tx_df = pd.DataFrame(
        [
            {
                "id": t.id,
                "timestamp": t.timestamp,
                "member": t.member.name,
                "member_email": t.member.email,
                "item": t.item.name,
                "action": t.action,
                "qty": t.qty,
                "due_date": t.due_date,
                "notes": t.notes,
            }
            for t in transactions
        ]
    )
    attendance_df = pd.DataFrame(
        [
            {
                "id": s.id,
                "member": s.member.name,
                "member_email": s.member.email,
                "uid": s.scanned_uid,
                "attendance_date": s.attendance_date,
                "scanned_at": s.scanned_at,
            }
            for s in scans
        ]
    )
    jobs_df = pd.DataFrame(
        [
            {
                "id": j.id,
                "member": j.member.name,
                "member_email": j.member.email,
                "printer_type": j.printer_type,
                "file_name": j.file_name,
                "file_path": j.file_path,
                "status": j.status,
                "notes": j.notes,
                "submitted_at": j.submitted_at,
                "started_at": j.started_at,
                "completed_at": j.completed_at,
            }
            for j in jobs
        ]
    )
    meetings_df = pd.DataFrame(
        [
            {
                "id": m.id,
                "team_name": m.team_name,
                "requester_email": m.requester_email,
                "room": m.room,
                "meeting_date": m.meeting_date,
                "start_time": m.start_time,
                "end_time": m.end_time,
                "notes": m.notes,
                "google_event_id": m.google_event_id,
                "google_calendar_id": m.google_calendar_id,
                "cancel_request_token": m.cancel_request_token,
                "cancel_requested_at": m.cancel_requested_at,
                "created_at": m.created_at,
            }
            for m in meetings
        ]
    )

    export_path = Path(app.instance_path) / "inventory_export.xlsx"
    with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
        members_df.to_excel(writer, index=False, sheet_name="Members")
        items_df.to_excel(writer, index=False, sheet_name="Items")
        tx_df.to_excel(writer, index=False, sheet_name="Transactions")
        attendance_df.to_excel(writer, index=False, sheet_name="Attendance")
        jobs_df.to_excel(writer, index=False, sheet_name="PrintJobs")
        meetings_df.to_excel(writer, index=False, sheet_name="Meetings")

    return send_file(export_path, as_attachment=True, download_name="inventory_export.xlsx")


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        ensure_meeting_schema_columns()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
