import os
import secrets
import smtplib
import subprocess
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from email.message import EmailMessage
from mimetypes import guess_type
from pathlib import Path
from urllib.parse import quote_plus
from uuid import uuid4

import pandas as pd
from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, session, url_for
from flask_login import current_user, login_user, logout_user
from sqlalchemy import and_, inspect, or_, text
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from auth import admin_required, elevated_required, init_auth, member_required
from models import AttendanceScan, Item, Meeting, Member, PrintJob, Transaction, db

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("ASME_DATABASE_URL", "sqlite:///inventory.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SECRET_KEY"] = os.environ.get("ASME_SECRET_KEY", "asme-dev-secret")
app.config["SESSION_PERMANENT"] = False
db.init_app(app)
init_auth(app)

PRINTER_TYPES = ("H2S", "P1S")
MEETING_ROOMS = ("Robotics Room", "Fluids Lab")
ROLE_CHOICES = ("member", "team_lead", "project_manager", "admin")
ELEVATED_ROLES = ("team_lead", "project_manager", "admin")
ROLE_LABELS = {
    "member": "Member",
    "team_lead": "Team Lead",
    "project_manager": "Project Manager",
    "admin": "Admin",
}
EXEC_TEAM_PLACEHOLDERS = (
    {
        "title": "President",
        "name": "Brayden Nagra",
        "summary": "Sets chapter direction, leads the exec board, and helps connect members with projects and long-term opportunities.",
        "image": "images/executive/brayden-nagra.jpg",
        "alt": "Brayden Nagra, President of ASME at Iowa",
        "featured": True,
    },
    {
        "title": "Vice President",
        "name": "Paul Conover",
        "summary": "Builds industry relationships, supports chapter direction, and helps oversee project teams across the organization.",
        "image": "images/executive/paul-conover.jpg",
        "alt": "Paul Conover, Vice President of ASME at Iowa",
    },
    {
        "title": "Treasurer",
        "name": "Adelai Kaiser",
        "summary": "Manages chapter finances, coordinates with SOBO, and keeps the club's budget and expenses organized.",
        "image": "images/executive/adelai-kaiser.jpg",
        "alt": "Adelai Kaiser, Treasurer of ASME at Iowa",
    },
    {
        "title": "Secretary",
        "name": "Picture Coming Soon",
        "summary": "Secretary role listed in the executive packet. Photo and short details can be added as soon as they are available.",
        "image": None,
        "alt": "Secretary role at ASME at Iowa, photo coming soon",
        "placeholder": True,
    },
    {
        "title": "Executive Coordinator",
        "name": "Jet Etnyre",
        "summary": "Keeps schedules, room reservations, and chapter events moving smoothly behind the scenes.",
        "image": "images/executive/jet-etnyre.jpg",
        "alt": "Jet Etnyre, Executive Coordinator of ASME at Iowa",
    },
    {
        "title": "Event Coordinator",
        "name": "Heriberto Salgado",
        "summary": "Organizes chapter events and creates more engaging ways for members to connect, learn, and get involved.",
        "image": "images/executive/heri-salgado.jpg",
        "alt": "Heriberto Salgado, Event Coordinator of ASME at Iowa",
    },
)
PUBLIC_SOCIAL_LINKS = (
    {
        "title": "Instagram",
        "handle": "@uiowaasme",
        "href": "https://www.instagram.com/uiowaasme/",
        "summary": "Chapter highlights, build photos, and event updates.",
    },
    {
        "title": "X",
        "handle": "@UIowaASME",
        "href": "https://x.com/UIowaASME",
        "summary": "Announcements, chapter visibility, and quick updates.",
    },
    {
        "title": "Facebook",
        "handle": "UIowaASME",
        "href": "https://www.facebook.com/UIowaASME",
        "summary": "Club presence, outreach, and public-facing updates.",
    },
    {
        "title": "Email",
        "handle": "studorg-asme@uiowa.edu",
        "href": "mailto:studorg-asme@uiowa.edu",
        "summary": "Reach the chapter directly about joining, sponsorship, or questions.",
    },
)
PUBLIC_NAV_ITEMS = (
    {"key": "home", "label": "Home", "endpoint": "home"},
    {"key": "about", "label": "About", "endpoint": "about_page"},
    {"key": "team", "label": "Executive Team", "endpoint": "exec_team"},
    {"key": "projects", "label": "Projects", "endpoint": "projects_page"},
    {"key": "join", "label": "Join", "endpoint": "join_page"},
    {"key": "contact", "label": "Contact", "endpoint": "contact_page"},
    {"key": "login", "label": "Login", "endpoint": "login"},
)
FEATURED_MESSAGE = {
    "eyebrow": "Executive Message",
    "title": "Hear where the chapter is headed.",
    "summary": (
        "Get a quick look at the chapter's direction, the work members are building, and what leadership is focused on this year."
    ),
    "href": "https://www.instagram.com/uiowaasme/",
    "cta": "Hear from Leadership",
    "poster": "media/asme-nationals-2025.jpeg",
}
HOME_PREVIEW_CARDS = (
    {
        "title": "About",
        "summary": "Learn what ASME at Iowa stands for and how the chapter is organized.",
        "cta": "View About",
        "endpoint": "about_page",
    },
    {
        "title": "Projects",
        "summary": "See the competition builds, fabrication work, and systems projects driving the chapter.",
        "cta": "Explore Projects",
        "endpoint": "projects_page",
    },
    {
        "title": "Join",
        "summary": "Find the clearest path into meetings, projects, and the member experience.",
        "cta": "Join ASME",
        "endpoint": "join_page",
    },
)
ABOUT_FEATURES = (
    {
        "label": "Mission",
        "title": "Turn classroom learning into visible, hands-on engineering work.",
        "summary": (
            "ASME at Iowa gives students a place to design, build, test, organize, and lead together so their "
            "engineering experience feels active before graduation."
        ),
    },
    {
        "label": "Culture",
        "title": "Build a chapter identity around momentum, accountability, and peer support.",
        "summary": (
            "Members learn faster when projects, operations, and leadership feel connected. The chapter is designed "
            "to keep those pieces moving together."
        ),
    },
    {
        "label": "Outcome",
        "title": "Graduate with stronger stories, better instincts, and more public-facing experience.",
        "summary": (
            "Students leave with technical reps, leadership signals, and a clearer sense of how engineering happens "
            "on real teams."
        ),
    },
)
ABOUT_RHYTHM = (
    {
        "title": "Recruit and orient",
        "summary": "New members get a clear introduction to the chapter, active work, and where they can contribute first.",
    },
    {
        "title": "Build and collaborate",
        "summary": "Teams move through design, fabrication, testing, logistics, and communication with support from peers and officers.",
    },
    {
        "title": "Lead and hand off",
        "summary": "Experienced members grow into leadership, improve systems, and leave a better chapter for the next cycle.",
    },
)
PROJECT_SHOWCASE = (
    {
        "label": "Competition Build",
        "title": "Design Build Fly",
        "summary": (
            "Student teams design and fabricate an unmanned electric aircraft to hit a demanding mission profile with "
            "strong flight performance and disciplined manufacturing."
        ),
        "focus": ("Airframe design", "Manufacturing strategy", "Flight test iteration"),
    },
    {
        "label": "Robotics + Manufacturing",
        "title": "Additive Manufacturing Mars Rover",
        "summary": (
            "Members create a resource-handling rover that blends mechanical systems, fabrication, and mission-driven "
            "problem solving on a competition-style terrain."
        ),
        "focus": ("Subsystem design", "Rapid prototyping", "Controls integration"),
    },
    {
        "label": "Systems Design",
        "title": "Automated Garbage Truck",
        "summary": (
            "A student-built automated collection concept focused on terrain handling, material sorting, and moving "
            "waste to the correct destination."
        ),
        "focus": ("Mechanism design", "Automation logic", "System reliability"),
    },
    {
        "label": "Chapter Infrastructure",
        "title": "Operations, Tooling, and Fabrication Support",
        "summary": (
            "The chapter also grows through process work such as print services, inventory systems, workshop support, "
            "and other member-facing engineering operations."
        ),
        "focus": ("Shop readiness", "Member support", "Operational systems"),
    },
)
PROJECT_STAGES = (
    {"label": "Design", "summary": "Translate ideas into feasible concepts, requirements, and technical direction."},
    {"label": "Build", "summary": "Prototype, fabricate, assemble, and solve the real constraints that appear in hardware."},
    {"label": "Test", "summary": "Evaluate what works, adjust quickly, and learn from results instead of guessing."},
    {"label": "Lead", "summary": "Coordinate teammates, communicate clearly, and keep momentum high across the chapter."},
)
JOIN_STEPS = (
    {
        "step": "01",
        "title": "Start with a meeting, event, or direct message",
        "summary": "The easiest entry point is simply to show up, introduce yourself, and see which work feels interesting.",
    },
    {
        "step": "02",
        "title": "Find a project lane or support role",
        "summary": "Members can plug into active builds, chapter operations, or areas where the team needs extra hands.",
    },
    {
        "step": "03",
        "title": "Learn fast by contributing early",
        "summary": "You do not need to arrive as an expert. You grow by helping with real work alongside returning members.",
    },
)
JOIN_BENEFITS = (
    "Hands-on engineering work beyond coursework",
    "Leadership experience that feels credible and visible",
    "A stronger engineering network on campus",
    "Project stories that matter when interviewing",
)
LOGIN_CHOICES = (
    {
        "label": "Member Access",
        "title": "Enter the member workspace.",
        "summary": "Use your chapter credentials to manage checkouts, submit print jobs, and access the club schedule.",
        "cta": "Member Login",
        "endpoint": "member_portal_entry",
    },
    {
        "label": "Admin Access",
        "title": "Open the operations workspace.",
        "summary": "Reserved for admin-role accounts managing attendance, inventory, meetings, printing, and exports.",
        "cta": "Admin Login",
        "endpoint": "admin_portal_entry",
    },
)
PUBLIC_CONTACT = {
    "email": "studorg-asme@uiowa.edu",
    "official_url": "https://asme.org.uiowa.edu/",
    "location": "University of Iowa, Iowa City, IA",
}
ALLOWED_GCODE_EXTENSIONS = {"gcode", "gco", "3mf"}
UPLOAD_DIR = Path(app.instance_path) / "gcode_uploads"
PRINT_COMMANDS_ENV_FILE = Path(app.instance_path) / "print_commands.env"
DEFAULT_GCAL_SERVICE_ACCOUNT_FILE = Path(app.instance_path) / "google_service_account.json"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

MEMBER_SCHEMA_READY = False
MEETING_SCHEMA_READY = False


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
            if key:
                os.environ[key] = value
    except Exception:
        pass


load_local_print_command_env()


def default_due_date(days=7):
    return date.today() + timedelta(days=days)


def parse_int(value, default=1):
    try:
        parsed = int(value)
    except Exception:
        return default
    return parsed if parsed > 0 else default


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


def normalize_email(value):
    return (value or "").strip().lower()


def normalize_role(value):
    role = (value or "member").strip().lower()
    return role if role in ROLE_CHOICES else "member"


def safe_redirect_target(raw_target, default_endpoint):
    if raw_target and raw_target.startswith("/") and not raw_target.startswith("//"):
        return raw_target
    return url_for(default_endpoint)


def role_home_endpoint(user):
    return "admin_dashboard" if getattr(user, "role", "") == "admin" else "member_dashboard"


def role_home_redirect(user, raw_target=None):
    return redirect(safe_redirect_target(raw_target, role_home_endpoint(user)))


def has_password_bootstrap():
    return (
        Member.query.filter(Member.password_hash.isnot(None), Member.password_hash != "").count() > 0
    )


def account_setup_needed():
    return not has_password_bootstrap()


def portal_login_config(portal):
    configs = {
        "member": {
            "portal": "member",
            "kicker": "Members Only",
            "title": "Member Portal",
            "heading": "Sign in to the club workspace.",
            "description": (
                "Access member-only pages, club resources, and future portal sections. "
                "For now, this is the clean login entry point for members."
            ),
            "button_label": "Enter Member Portal",
            "support_label": "Need access?",
            "support_copy": "Contact an officer or admin to activate your account.",
        },
        "admin": {
            "portal": "admin",
            "kicker": "Admin Access",
            "title": "Admin Portal",
            "heading": "Sign in to the operations workspace.",
            "description": (
                "This portal is reserved for admin access. Use it to manage operations, and we can shape "
                "the admin dashboard content next once you decide what belongs here."
            ),
            "button_label": "Enter Admin Portal",
            "support_label": "Restricted access",
            "support_copy": "Only admin-role accounts can sign in here.",
        },
    }
    return configs.get(portal, configs["member"])


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

    fallback_url = (
        "https://calendar.google.com/calendar/embed"
        "?src=en.usa%23holiday%40group.v.calendar.google.com"
        "&ctz=America%2FChicago&mode=WEEK&showTabs=0&showPrint=0&showCalendars=0&showTz=0"
    )
    return {"url": fallback_url, "placeholder": True}


def ensure_member_auth_schema_columns():
    global MEMBER_SCHEMA_READY
    if MEMBER_SCHEMA_READY:
        return

    inspector = inspect(db.engine)
    table_names = inspector.get_table_names()
    if "members" not in table_names:
        db.create_all()
        inspector = inspect(db.engine)
        table_names = inspector.get_table_names()
    if "members" not in table_names:
        return

    existing_columns = {col["name"] for col in inspector.get_columns("members")}
    alters = []
    if "password_hash" not in existing_columns:
        alters.append("ALTER TABLE members ADD COLUMN password_hash VARCHAR(256)")
    if "role" not in existing_columns:
        alters.append("ALTER TABLE members ADD COLUMN role VARCHAR(30) NOT NULL DEFAULT 'member'")
    if "is_active" not in existing_columns:
        alters.append("ALTER TABLE members ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT TRUE")

    if alters:
        with db.engine.begin() as conn:
            for statement in alters:
                conn.execute(text(statement))

    MEMBER_SCHEMA_READY = True


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
        alters.append("ALTER TABLE meetings ADD COLUMN cancel_requested_at TIMESTAMP")

    if alters:
        with db.engine.begin() as conn:
            for statement in alters:
                conn.execute(text(statement))

    MEETING_SCHEMA_READY = True


def ensure_database_ready():
    db.create_all()
    ensure_member_auth_schema_columns()
    ensure_meeting_schema_columns()


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
        return None, None, "Google Calendar ID is not configured. Set ASME_GOOGLE_CALENDAR_ID or room-specific IDs."

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

    return {
        "service_account_file": service_account_file,
        "service_account_file_exists": Path(service_account_file).exists(),
        "default_calendar_id": default_calendar_id,
        "robotics_calendar_id": robotics_calendar_id,
        "fluids_calendar_id": fluids_calendar_id,
        "has_any_calendar_id": bool(default_calendar_id or robotics_calendar_id or fluids_calendar_id),
        "smtp_ready": bool(smtp_user and smtp_pass),
        "smtp_user": smtp_user,
        "cancel_notify_to": cancel_notify_to,
        "timezone": get_calendar_timezone(),
    }


def nfc_secret_matches():
    configured = (os.environ.get("ASME_NFC_SECRET") or "").strip()
    provided = (request.headers.get("X-NFC-Secret") or "").strip()
    return bool(configured and provided) and secrets.compare_digest(configured, provided)


def attendance_scan_authorized():
    return nfc_secret_matches() or (current_user.is_authenticated and current_user.role == "admin")


def allowed_gcode(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_GCODE_EXTENSIONS


def append_note(existing, message):
    if not message:
        return existing
    return f"{existing} | {message}" if existing else message


def remove_job_file(job):
    if not job.file_path:
        return False, None
    if not os.path.exists(job.file_path):
        return False, None
    try:
        os.remove(job.file_path)
        return True, None
    except Exception as exc:
        return False, str(exc)


def delete_print_job_with_file(job):
    file_removed, file_error = remove_job_file(job)
    file_name = job.file_name
    printer_type = job.printer_type
    db.session.delete(job)
    db.session.commit()
    dispatch_next_job(printer_type)
    return {"file_name": file_name, "file_removed": file_removed, "file_error": file_error}


def resolve_member(member_tag=None, member_id=None, fallback_member=None, active_only=True):
    if fallback_member is not None:
        return fallback_member if (fallback_member.is_active or not active_only) else None

    tag = (member_tag or "").strip()
    raw_id = (member_id or "").strip()

    if tag:
        member = Member.query.filter_by(nfc_tag=tag).first()
        if member and (member.is_active or not active_only):
            return member

    if raw_id:
        try:
            member = db.session.get(Member, int(raw_id))
        except Exception:
            return None
        if member and (member.is_active or not active_only):
            return member
    return None


def resolve_item(item_tag=None, item_id=None):
    tag = (item_tag or "").strip()
    raw_id = (item_id or "").strip()

    if tag:
        item = Item.query.filter_by(nfc_tag=tag).first()
        if item:
            return item

    if raw_id:
        try:
            return db.session.get(Item, int(raw_id))
        except Exception:
            return None
    return None


def scan_attendance_uid(uid):
    cleaned_uid = (uid or "").strip()
    if not cleaned_uid:
        return False, "Scan failed: UID was empty."

    member = Member.query.filter_by(nfc_tag=cleaned_uid).first()
    if not member or not member.is_active:
        return False, "UID not recognized. Pair this UID to an active member first."

    scan = AttendanceScan(member_id=member.id, scanned_uid=cleaned_uid, attendance_date=date.today())
    db.session.add(scan)
    db.session.commit()

    scans_today = AttendanceScan.query.filter_by(member_id=member.id, attendance_date=date.today()).count()
    if scans_today == 1:
        return True, f"Attendance marked for {member.name}."
    return True, f"{member.name} scanned again. Attendance already marked for today."


def perform_inventory_transaction(member, item, action, qty, due_date_value=None, notes=None):
    if not member:
        return None, "Could not find member. Scan a member UID or select one."
    if not item:
        return None, "Could not find item. Scan an item UID or select one."

    normalized_action = (action or "").strip().lower()
    if normalized_action not in {"checkout", "return"}:
        return None, "Invalid inventory action."

    cleaned_qty = parse_int(qty, default=1)
    due_date_final = due_date_value or default_due_date()

    if normalized_action == "checkout":
        if item.available_qty < cleaned_qty:
            return None, f"Not enough stock. {item.name} has {item.available_qty} available."
        item.available_qty -= cleaned_qty
    else:
        item.available_qty = min(item.total_qty, item.available_qty + cleaned_qty)
        due_date_final = None

    tx = Transaction(
        member_id=member.id,
        item_id=item.id,
        action=normalized_action,
        qty=cleaned_qty,
        due_date=due_date_final,
        notes=(notes or "").strip() or None,
    )
    db.session.add(tx)
    db.session.commit()
    return tx, None


def create_print_job(member, printer_type, file_storage, notes=None, initial_status="pending"):
    if not member:
        return None, "Could not find member for print job. Scan/select member first."

    normalized_printer = (printer_type or "").strip().upper()
    if normalized_printer not in PRINTER_TYPES:
        return None, "Invalid printer queue selected."
    if not file_storage or not file_storage.filename:
        return None, "Please upload a print file."
    if not allowed_gcode(file_storage.filename):
        return None, "Invalid file type. Use .gcode, .gco, or .3mf."

    original_name = secure_filename(file_storage.filename)
    stored_name = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{original_name}"
    file_path = UPLOAD_DIR / stored_name
    file_storage.save(file_path)

    job = PrintJob(
        member_id=member.id,
        printer_type=normalized_printer,
        file_name=original_name,
        file_path=str(file_path),
        notes=(notes or "").strip() or None,
        status=initial_status,
    )
    db.session.add(job)
    db.session.commit()
    return job, None


def launch_print_command(job):
    env_name = f"ASME_{job.printer_type}_PRINT_CMD"
    cmd_template = os.environ.get(env_name)
    if not cmd_template:
        return f"{env_name} is not configured. Add it to {PRINT_COMMANDS_ENV_FILE} and restart the app."

    command = cmd_template.format(file=job.file_path, filename=job.file_name, job_id=job.id)
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    if result.returncode == 0:
        return None

    error_text = (result.stderr or result.stdout or "print command failed").strip()
    return f"{env_name} failed: {error_text[:300]}"


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
    next_job.completed_at = None
    db.session.commit()

    dispatch_error = launch_print_command(next_job)
    if dispatch_error:
        next_job.status = "failed"
        next_job.completed_at = datetime.utcnow()
        next_job.notes = append_note(next_job.notes, dispatch_error)
        db.session.commit()
        return dispatch_next_job(printer_type)

    return next_job


def accept_print_job(job):
    if not job or job.status != "pending":
        return None, "Only pending jobs can be accepted."

    job.status = "queued"
    job.started_at = None
    job.completed_at = None
    db.session.commit()

    started = dispatch_next_job(job.printer_type)
    if started and started.id == job.id:
        return started, f"Accepted job #{job.id}. It auto-started on {job.printer_type}."
    return started, f"Accepted job #{job.id}. It is queued for {job.printer_type}."


def fail_print_job_record(job, remove_file=False, note=None):
    if not job:
        return None

    previous_status = job.status
    file_removed = False
    file_error = None
    if remove_file:
        file_removed, file_error = remove_job_file(job)

    job.status = "failed"
    job.completed_at = datetime.utcnow()
    job.notes = append_note(job.notes, note)
    if remove_file and file_error:
        job.notes = append_note(job.notes, f"File delete failed: {file_error[:200]}")
    db.session.commit()

    if previous_status in {"queued", "printing"}:
        dispatch_next_job(job.printer_type)

    return {"file_removed": file_removed, "file_error": file_error}


def complete_print_job_record(job):
    if not job:
        return
    job.status = "done"
    job.completed_at = datetime.utcnow()
    db.session.commit()
    dispatch_next_job(job.printer_type)


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
        "role": member.role,
        "is_active": member.is_active,
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
    if not job:
        return None
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


def value_from_request(key, default=None):
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        return payload.get(key, default)
    return request.form.get(key, default)


def api_error(message, status=400):
    return jsonify({"ok": False, "error": message}), status


def api_success(message=None, payload=None, status=200):
    return jsonify({"ok": True, "message": message, "payload": payload or {}}), status


def get_today_attendance_unique():
    scans = (
        AttendanceScan.query.filter_by(attendance_date=date.today())
        .order_by(AttendanceScan.scanned_at.desc())
        .all()
    )
    unique = []
    seen = set()
    for scan in scans:
        if scan.member_id in seen:
            continue
        seen.add(scan.member_id)
        unique.append(scan)
    return unique


def find_conflicting_meeting(room, meeting_date_value, start_time_value, end_time_value, ignore_meeting_id=None):
    query = Meeting.query.filter(
        Meeting.room == room,
        Meeting.meeting_date == meeting_date_value,
        Meeting.start_time < end_time_value,
        Meeting.end_time > start_time_value,
        Meeting.cancel_request_token.is_(None),
    )
    if ignore_meeting_id is not None:
        query = query.filter(Meeting.id != ignore_meeting_id)
    return query.order_by(Meeting.start_time.asc(), Meeting.id.asc()).first()


def parse_meeting_request_form():
    team_name = (request.form.get("team_name") or "").strip()
    requester_email = normalize_email(request.form.get("requester_email"))
    room = (request.form.get("room") or "").strip()
    meeting_date_value = parse_due_date(request.form.get("meeting_date"))
    start_time_value = parse_clock_time(request.form.get("start_time"))
    end_time_value = parse_clock_time(request.form.get("end_time"))
    notes = (request.form.get("notes") or "").strip() or None

    if not team_name:
        return None, "Team name is required."
    if room not in MEETING_ROOMS:
        return None, "Please select Robotics Room or Fluids Lab."
    if not meeting_date_value:
        return None, "Meeting date is required."
    if meeting_date_value < date.today():
        return None, "Meeting date cannot be in the past."
    if not start_time_value or not end_time_value:
        return None, "Start and end times are required."
    if end_time_value <= start_time_value:
        return None, "End time must be after start time."

    conflicting = find_conflicting_meeting(room, meeting_date_value, start_time_value, end_time_value)
    if conflicting:
        return None, (
            f"{room} is already booked by {conflicting.team_name} from "
            f"{conflicting.start_time.strftime('%H:%M')} to {conflicting.end_time.strftime('%H:%M')}."
        )

    return {
        "team_name": team_name,
        "requester_email": requester_email or None,
        "room": room,
        "meeting_date": meeting_date_value,
        "start_time": start_time_value,
        "end_time": end_time_value,
        "notes": notes,
    }, None


def get_confirmed_meetings(limit=None):
    ensure_meeting_schema_columns()
    now = datetime.now()
    query = (
        Meeting.query.filter(
            or_(
                Meeting.meeting_date > now.date(),
                and_(Meeting.meeting_date == now.date(), Meeting.end_time >= now.time()),
            )
        )
        .filter(Meeting.cancel_request_token.is_(None))
        .filter(Meeting.google_event_id.isnot(None))
        .order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc())
    )
    if limit is not None:
        query = query.limit(limit)
    return query.all()


def get_pending_meeting_requests():
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
        .filter(Meeting.google_event_id.is_(None))
        .order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc())
        .all()
    )


def get_member_meetings(member):
    ensure_meeting_schema_columns()
    if not member.email:
        return []
    now = datetime.now()
    return (
        Meeting.query.filter(
            Meeting.requester_email == normalize_email(member.email),
            or_(
                Meeting.meeting_date > now.date(),
                and_(Meeting.meeting_date == now.date(), Meeting.end_time >= now.time()),
            ),
            Meeting.cancel_request_token.is_(None),
        )
        .order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc())
        .all()
    )


def build_active_checkout_lots(member_id=None):
    query = Transaction.query.order_by(Transaction.timestamp.asc(), Transaction.id.asc())
    if member_id is not None:
        query = query.filter_by(member_id=member_id)

    lots_by_key = defaultdict(list)
    for tx in query.all():
        key = (tx.member_id, tx.item_id)
        bucket = lots_by_key[key]
        if tx.action == "checkout":
            bucket.append(
                {
                    "member": tx.member,
                    "item": tx.item,
                    "qty": tx.qty,
                    "remaining_qty": tx.qty,
                    "due_date": tx.due_date,
                    "checked_out_at": tx.timestamp,
                    "notes": tx.notes,
                }
            )
            continue

        remaining = tx.qty
        for lot in bucket:
            if lot["remaining_qty"] <= 0:
                continue
            consumed = min(lot["remaining_qty"], remaining)
            lot["remaining_qty"] -= consumed
            remaining -= consumed
            if remaining <= 0:
                break

    rows = []
    today_value = date.today()
    for bucket in lots_by_key.values():
        for lot in bucket:
            if lot["remaining_qty"] <= 0:
                continue
            due_date_value = lot["due_date"]
            rows.append(
                {
                    "member": lot["member"],
                    "item": lot["item"],
                    "qty": lot["remaining_qty"],
                    "due_date": due_date_value,
                    "checked_out_at": lot["checked_out_at"],
                    "notes": lot["notes"],
                    "is_overdue": bool(due_date_value and due_date_value < today_value),
                }
            )

    rows.sort(
        key=lambda row: (
            row["due_date"] is None,
            row["due_date"] or date.max,
            row["checked_out_at"] or datetime.min,
        )
    )
    return rows


def get_queue_state():
    queues = {}
    for printer in PRINTER_TYPES:
        queues[printer] = {
            "pending": (
                PrintJob.query.filter_by(printer_type=printer, status="pending")
                .order_by(PrintJob.submitted_at.asc(), PrintJob.id.asc())
                .all()
            ),
            "active": (
                PrintJob.query.filter_by(printer_type=printer, status="printing")
                .order_by(PrintJob.started_at.asc(), PrintJob.id.asc())
                .first()
            ),
            "queued": (
                PrintJob.query.filter_by(printer_type=printer, status="queued")
                .order_by(PrintJob.submitted_at.asc(), PrintJob.id.asc())
                .all()
            ),
            "recent_finished": (
                PrintJob.query.filter(
                    PrintJob.printer_type == printer,
                    PrintJob.status.in_(["done", "failed"]),
                )
                .order_by(PrintJob.completed_at.desc(), PrintJob.id.desc())
                .limit(8)
                .all()
            ),
        }
    return queues


def get_low_stock_items(limit=None):
    query = Item.query.filter(Item.available_qty <= 2).order_by(Item.available_qty.asc(), Item.name.asc())
    if limit is not None:
        query = query.limit(limit)
    return query.all()


def member_last_active(member):
    timestamps = [member.created_at]

    attendance = (
        AttendanceScan.query.filter_by(member_id=member.id)
        .order_by(AttendanceScan.scanned_at.desc())
        .first()
    )
    if attendance:
        timestamps.append(attendance.scanned_at)

    transaction = (
        Transaction.query.filter_by(member_id=member.id)
        .order_by(Transaction.timestamp.desc())
        .first()
    )
    if transaction:
        timestamps.append(transaction.timestamp)

    print_job = (
        PrintJob.query.filter_by(member_id=member.id)
        .order_by(PrintJob.submitted_at.desc())
        .first()
    )
    if print_job:
        timestamps.append(print_job.submitted_at)

    meeting = (
        Meeting.query.filter_by(requester_email=normalize_email(member.email))
        .order_by(Meeting.created_at.desc())
        .first()
    )
    if meeting:
        timestamps.append(meeting.created_at)

    timestamps = [value for value in timestamps if value is not None]
    return max(timestamps) if timestamps else None


def build_member_rows():
    rows = []
    for member in Member.query.order_by(Member.is_active.desc(), Member.name.asc()).all():
        rows.append(
            {
                "record": member,
                "last_active": member_last_active(member),
                "nfc_paired": bool(member.nfc_tag),
            }
        )
    return rows


def build_admin_metrics():
    active_members = Member.query.filter_by(is_active=True).count()
    attendance_count = len(get_today_attendance_unique())
    low_stock_count = Item.query.filter(Item.available_qty <= 2).count()
    active_checkout_qty = sum(row["qty"] for row in build_active_checkout_lots())
    pending_print_count = PrintJob.query.filter_by(status="pending").count()
    queue_depth = PrintJob.query.filter(PrintJob.status.in_(["queued", "printing"])).count()
    upcoming_meeting_count = len(get_confirmed_meetings())
    return [
        {"label": "Active Members", "value": active_members},
        {"label": "Present Today", "value": attendance_count},
        {"label": "Low Stock", "value": low_stock_count},
        {"label": "Checked Out", "value": active_checkout_qty},
        {"label": "Awaiting Approval", "value": pending_print_count},
        {"label": "Queue + Meetings", "value": queue_depth + upcoming_meeting_count},
    ]


def public_context(page="home"):
    return {
        "public_page": page,
        "public_nav_items": PUBLIC_NAV_ITEMS,
        "featured_message": FEATURED_MESSAGE,
        "home_preview_cards": HOME_PREVIEW_CARDS,
        "about_features": ABOUT_FEATURES,
        "about_rhythm": ABOUT_RHYTHM,
        "project_showcase": PROJECT_SHOWCASE,
        "project_stages": PROJECT_STAGES,
        "join_steps": JOIN_STEPS,
        "join_benefits": JOIN_BENEFITS,
        "login_choices": LOGIN_CHOICES,
        "public_contact": PUBLIC_CONTACT,
        "exec_team_preview": EXEC_TEAM_PLACEHOLDERS[:3],
        "exec_team_entries": EXEC_TEAM_PLACEHOLDERS,
        "public_social_links": PUBLIC_SOCIAL_LINKS,
    }


def render_portal_login(portal, next_target="", forgot_password_notice=False):
    return render_template(
        "auth/login.html",
        **public_context(page="login"),
        portal_config=portal_login_config(portal),
        next_target=next_target,
        setup_available=account_setup_needed(),
        forgot_password_notice=forgot_password_notice,
    )


def member_base_context(active_page, page_title, page_subtitle):
    my_checkouts = build_active_checkout_lots(current_user.id)
    my_open_jobs = (
        PrintJob.query.filter(
            PrintJob.member_id == current_user.id,
            PrintJob.status.in_(["pending", "queued", "printing", "done", "failed"]),
        )
        .order_by(PrintJob.submitted_at.desc(), PrintJob.id.desc())
        .limit(8)
        .all()
    )
    return {
        "active_page": active_page,
        "page_title": page_title,
        "page_subtitle": page_subtitle,
        "my_checkout_count": sum(row["qty"] for row in my_checkouts),
        "my_pending_print_count": len([job for job in my_open_jobs if job.status in {"pending", "queued", "printing"}]),
        "elevated_roles": ELEVATED_ROLES,
    }


def admin_base_context(active_page, page_title, page_subtitle):
    return {
        "active_page": active_page,
        "page_title": page_title,
        "page_subtitle": page_subtitle,
        "admin_metrics": build_admin_metrics(),
    }


def build_admin_bootstrap_payload():
    members = Member.query.order_by(Member.name.asc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    attendance = get_today_attendance_unique()
    transactions = Transaction.query.order_by(Transaction.timestamp.desc()).limit(20).all()
    queues = get_queue_state()
    return {
        "today": str(date.today()),
        "default_due": str(default_due_date()),
        "members": [serialize_member(member) for member in members],
        "items": [serialize_item(item) for item in items],
        "attendance_today": [serialize_attendance_scan(scan) for scan in attendance],
        "attendance_count": len(attendance),
        "recent_transactions": [serialize_transaction(tx) for tx in transactions],
        "pending_print_jobs": [serialize_print_job(job) for job in PrintJob.query.filter_by(status="pending").all()],
        "queues": {
            printer: {
                "pending": [serialize_print_job(job) for job in snapshot["pending"]],
                "active": serialize_print_job(snapshot["active"]),
                "queued": [serialize_print_job(job) for job in snapshot["queued"]],
                "recent_finished": [serialize_print_job(job) for job in snapshot["recent_finished"]],
            }
            for printer, snapshot in queues.items()
        },
    }


def build_member_bootstrap_payload(member):
    checkouts = build_active_checkout_lots(member.id)
    jobs = (
        PrintJob.query.filter_by(member_id=member.id)
        .order_by(PrintJob.submitted_at.desc(), PrintJob.id.desc())
        .limit(12)
        .all()
    )
    return {
        "today": str(date.today()),
        "default_due": str(default_due_date()),
        "member": serialize_member(member),
        "active_checkouts": [
            {
                "item_name": row["item"].name,
                "qty": row["qty"],
                "due_date": iso_or_none(row["due_date"]),
                "checked_out_at": iso_or_none(row["checked_out_at"]),
                "is_overdue": row["is_overdue"],
            }
            for row in checkouts
        ],
        "print_jobs": [serialize_print_job(job) for job in jobs],
    }


def status_badge_class(value):
    normalized = (value or "").strip().lower()
    mapping = {
        "present": "badge-green",
        "in stock": "badge-green",
        "done": "badge-green",
        "approved": "badge-green",
        "pending": "badge-amber",
        "warning": "badge-amber",
        "low stock": "badge-amber",
        "overdue": "badge-red",
        "failed": "badge-red",
        "inactive": "badge-red",
        "queued": "badge-blue",
        "printing": "badge-blue",
        "active": "badge-blue",
        "scheduled": "badge-blue",
        "admin": "badge-gold",
    }
    return mapping.get(normalized, "badge-blue")


def role_badge_class(role):
    if role == "admin":
        return "badge-gold"
    if role in {"team_lead", "project_manager"}:
        return "badge-blue"
    return "badge-green"


def can_access_print_job(job):
    return current_user.is_authenticated and (current_user.role == "admin" or job.member_id == current_user.id)


def handle_member_transaction_request():
    item = resolve_item(request.form.get("item_tag"), request.form.get("item_id"))
    action = request.form.get("action")
    qty = request.form.get("qty")
    due = parse_due_date(request.form.get("due_date"))
    notes = request.form.get("notes")
    return perform_inventory_transaction(current_user, item, action, qty, due, notes)


def handle_admin_transaction_request():
    member = resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    item = resolve_item(request.form.get("item_tag"), request.form.get("item_id"))
    action = request.form.get("action")
    qty = request.form.get("qty")
    due = parse_due_date(request.form.get("due_date"))
    notes = request.form.get("notes")
    return perform_inventory_transaction(member, item, action, qty, due, notes)


@app.context_processor
def inject_template_globals():
    return {
        "current_user": current_user,
        "role_labels": ROLE_LABELS,
        "badge_class": status_badge_class,
        "role_badge_class": role_badge_class,
        "today_value": date.today(),
    }


@app.get("/")
def home():
    return render_template("public/home.html", **public_context(page="home"))


@app.get("/about")
def about_page():
    return render_template("public/about.html", **public_context(page="about"))


@app.get("/executive-team")
@app.get("/exec-team")
def exec_team():
    return render_template("public/exec_team.html", **public_context(page="team"))


@app.get("/projects")
def projects_page():
    return render_template("public/projects.html", **public_context(page="projects"))


@app.get("/join")
def join_page():
    return render_template("public/join.html", **public_context(page="join"))


@app.get("/contact")
@app.get("/socials")
def contact_page():
    return render_template("public/contact.html", **public_context(page="contact"))


def handle_portal_login(portal):
    ensure_database_ready()

    if account_setup_needed():
        return redirect(url_for("setup"))
    next_target = request.args.get("next", "") if request.method == "GET" else request.form.get("next", "")

    if current_user.is_authenticated:
        if portal == "admin" and current_user.role != "admin":
            flash("Admin access required.", "error")
            return redirect(url_for("member_dashboard"))
        return role_home_redirect(current_user, next_target)

    if request.method == "POST":
        email = normalize_email(request.form.get("email"))
        password = request.form.get("password") or ""

        member = Member.query.filter(db.func.lower(Member.email) == email).first() if email else None
        if (
            member is None
            or not member.password_hash
            or not member.is_active
            or not check_password_hash(member.password_hash, password)
        ):
            flash("Invalid email or password.", "error")
        elif portal == "admin" and member.role != "admin":
            flash("Admin access required. Use the member portal for non-admin accounts.", "error")
        else:
            session.clear()
            login_user(member, remember=False)
            session.permanent = False
            flash(f"Welcome back, {member.name}.", "success")
            return role_home_redirect(member, next_target)

    return render_portal_login(portal, next_target=next_target, forgot_password_notice=False)


@app.route("/member/login", methods=["GET", "POST"])
@app.route("/member", methods=["GET", "POST"])
def member_portal_entry():
    return handle_portal_login("member")


@app.route("/admin/login", methods=["GET", "POST"])
@app.route("/admin", methods=["GET", "POST"])
def admin_portal_entry():
    return handle_portal_login("admin")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        return handle_portal_login("member")
    return render_template("public/login.html", **public_context(page="login"))


@app.route("/setup", methods=["GET", "POST"])
def setup():
    ensure_database_ready()
    if has_password_bootstrap():
        flash("Setup is already complete. Please sign in.", "info")
        return redirect(url_for("admin_portal_entry"))

    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        email = normalize_email(request.form.get("email"))
        member_class = (request.form.get("member_class") or "").strip()
        password = request.form.get("password") or ""
        confirm_password = request.form.get("confirm_password") or ""

        if not name or not email or not member_class or not password:
            flash("All fields are required.", "error")
        elif password != confirm_password:
            flash("Passwords do not match.", "error")
        else:
            member = Member.query.filter(db.func.lower(Member.email) == email).first()
            if member is None:
                member = Member(name=name, email=email, member_class=member_class, role="admin", is_active=True)
                db.session.add(member)
            else:
                member.name = name
                member.email = email
                member.member_class = member_class
                member.role = "admin"
                member.is_active = True

            member.password_hash = generate_password_hash(password)
            db.session.commit()
            session.clear()
            login_user(member, remember=False)
            session.permanent = False
            flash("Admin account created. Welcome to ASME at Iowa.", "success")
            return redirect(url_for("admin_dashboard"))

    return render_template("auth/setup.html", **public_context(page="login"))


@app.get("/forgot-password")
def forgot_password():
    return render_portal_login("member", next_target="", forgot_password_notice=True)


@app.get("/logout")
def logout():
    if current_user.is_authenticated:
        logout_user()
        session.clear()
    return redirect(url_for("home"))


@app.get("/member/dashboard")
@member_required
def member_dashboard():
    calendar_embed = google_calendar_embed_context()
    my_checkouts = build_active_checkout_lots(current_user.id)
    my_jobs = (
        PrintJob.query.filter_by(member_id=current_user.id)
        .order_by(PrintJob.submitted_at.desc(), PrintJob.id.desc())
        .limit(8)
        .all()
    )
    requested_meetings = get_member_meetings(current_user)
    confirmed_meetings = get_confirmed_meetings(limit=6)
    context = member_base_context(
        active_page="dashboard",
        page_title="Dashboard",
        page_subtitle="Your equipment, print jobs, and club schedule at a glance.",
    )
    context.update(
        {
            "active_checkouts": my_checkouts,
            "my_print_jobs": my_jobs,
            "requested_meetings": requested_meetings,
            "confirmed_meetings": confirmed_meetings,
            "google_calendar_embed_url": calendar_embed["url"],
            "google_calendar_placeholder": calendar_embed["placeholder"],
        }
    )
    return render_template("member/dashboard.html", **context)


@app.get("/member/checkout")
@member_required
def member_checkout():
    context = member_base_context(
        active_page="checkout",
        page_title="Checkout Equipment",
        page_subtitle="Borrow tools and supplies under your member account.",
    )
    context.update(
        {
            "items": Item.query.order_by(Item.name.asc()).all(),
            "default_due": str(default_due_date()),
        }
    )
    return render_template("member/checkout.html", **context)


@app.get("/member/checkouts")
@member_required
def member_checkouts():
    context = member_base_context(
        active_page="checkouts",
        page_title="My Checkouts",
        page_subtitle="Track what you currently have checked out and what is overdue.",
    )
    context.update({"active_checkouts": build_active_checkout_lots(current_user.id)})
    return render_template("member/checkouts.html", **context)


@app.post("/member/transact")
@member_required
def member_transact():
    tx, error = handle_member_transaction_request()
    if error:
        flash(error, "error")
    else:
        flash(f"{tx.action.title()} saved: {tx.qty} x {tx.item.name}.", "success")
    return redirect(safe_redirect_target(request.form.get("next"), "member_checkout"))


@app.get("/member/print")
@member_required
def member_print():
    context = member_base_context(
        active_page="print",
        page_title="Submit Print Job",
        page_subtitle="Upload a file for review before it enters the print queue.",
    )
    context.update(
        {
            "printer_types": PRINTER_TYPES,
            "my_print_jobs": (
                PrintJob.query.filter_by(member_id=current_user.id)
                .order_by(PrintJob.submitted_at.desc(), PrintJob.id.desc())
                .limit(12)
                .all()
            ),
        }
    )
    return render_template("member/print.html", **context)


@app.post("/member/print/submit")
@member_required
def member_print_submit():
    job, error = create_print_job(
        current_user,
        request.form.get("printer_type"),
        request.files.get("gcode_file"),
        request.form.get("notes"),
        initial_status="pending",
    )
    if error:
        flash(error, "error")
    else:
        flash(f"{job.printer_type} job submitted for admin approval.", "success")
    return redirect(safe_redirect_target(request.form.get("next"), "member_print"))


@app.get("/member/calendar")
@member_required
def member_calendar():
    calendar_embed = google_calendar_embed_context()
    context = member_base_context(
        active_page="calendar",
        page_title="Calendar",
        page_subtitle="View the live ASME calendar and confirmed room bookings.",
    )
    context.update(
        {
            "confirmed_meetings": get_confirmed_meetings(limit=12),
            "google_calendar_embed_url": calendar_embed["url"],
            "google_calendar_placeholder": calendar_embed["placeholder"],
        }
    )
    return render_template("member/calendar.html", **context)


@app.get("/member/meeting/new")
@elevated_required
def member_meeting_new():
    context = member_base_context(
        active_page="meeting",
        page_title="Request Meeting",
        page_subtitle="Submit a room request for admin approval and calendar scheduling.",
    )
    context.update(
        {
            "meeting_rooms": MEETING_ROOMS,
            "calendar_default_date": str(date.today()),
        }
    )
    return render_template("member/meeting_new.html", **context)


@app.post("/member/meeting/submit")
@elevated_required
def member_meeting_submit():
    ensure_meeting_schema_columns()
    payload, error = parse_meeting_request_form()
    if error:
        flash(error, "error")
        return redirect(safe_redirect_target(request.form.get("next"), "member_meeting_new"))

    meeting = Meeting(
        team_name=payload["team_name"],
        requester_email=normalize_email(current_user.email),
        room=payload["room"],
        meeting_date=payload["meeting_date"],
        start_time=payload["start_time"],
        end_time=payload["end_time"],
        notes=payload["notes"],
    )
    db.session.add(meeting)
    db.session.commit()
    flash("Meeting request submitted for admin approval.", "success")
    return redirect(url_for("member_dashboard"))


@app.get("/admin/dashboard")
@admin_required
def admin_dashboard():
    queues = get_queue_state()
    context = admin_base_context(
        active_page="dashboard",
        page_title="Dashboard",
        page_subtitle="Club-wide operations across attendance, inventory, meetings, and printing.",
    )
    context.update(
        {
            "today_attendance": get_today_attendance_unique(),
            "low_stock_items": get_low_stock_items(limit=8),
            "pending_print_jobs": PrintJob.query.filter_by(status="pending").order_by(PrintJob.submitted_at.asc()).all(),
            "pending_meeting_requests": get_pending_meeting_requests(),
            "recent_transactions": Transaction.query.order_by(Transaction.timestamp.desc()).limit(12).all(),
            "queues": queues,
            "upcoming_meetings": get_confirmed_meetings(limit=8),
        }
    )
    return render_template("admin/dashboard.html", **context)


@app.get("/admin/attendance")
@admin_required
def admin_attendance():
    context = admin_base_context(
        active_page="attendance",
        page_title="Attendance",
        page_subtitle="Scan NFC tags and track who is present today.",
    )
    context.update({"today_attendance": get_today_attendance_unique()})
    return render_template("admin/attendance.html", **context)


@app.post("/admin/attendance/scan")
@app.post("/attendance/scan")
def attendance_scan():
    if not attendance_scan_authorized():
        if request.headers.get("X-NFC-Secret"):
            return jsonify({"ok": False, "error": "Unauthorized NFC scan."}), 403
        flash("Admin access or a valid NFC secret is required.", "error")
        return redirect(url_for("admin_portal_entry", next=request.path))

    success, message = scan_attendance_uid(request.form.get("uid") or value_from_request("uid", ""))
    if request.headers.get("X-NFC-Secret") or request.is_json:
        if success:
            return jsonify({"ok": True, "message": message}), 200
        return jsonify({"ok": False, "error": message}), 400

    flash(message, "success" if success else "error")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_attendance"))


@app.get("/admin/members")
@admin_required
def admin_members():
    context = admin_base_context(
        active_page="members",
        page_title="Members",
        page_subtitle="Manage member accounts, roles, passwords, and NFC pairing readiness.",
    )
    context.update({"member_rows": build_member_rows(), "role_choices": ROLE_CHOICES})
    return render_template("admin/members.html", **context)


@app.post("/admin/members/add")
@admin_required
def admin_members_add():
    name = (request.form.get("name") or "").strip()
    email = normalize_email(request.form.get("email"))
    member_class = (request.form.get("member_class") or "").strip()
    role = normalize_role(request.form.get("role"))
    password = request.form.get("password") or ""

    if not name or not email or not member_class or not password:
        flash("Name, email, class/year, role, and temporary password are required.", "error")
        return redirect(url_for("admin_members"))
    if Member.query.filter(db.func.lower(Member.email) == email).first():
        flash("A member with that email already exists.", "error")
        return redirect(url_for("admin_members"))

    member = Member(
        name=name,
        email=email,
        member_class=member_class,
        role=role,
        is_active=True,
        password_hash=generate_password_hash(password),
    )
    db.session.add(member)
    db.session.commit()
    flash(f"Added {member.name}.", "success")
    return redirect(url_for("admin_members"))


@app.post("/admin/members/<int:member_id>/edit")
@admin_required
def admin_member_edit(member_id):
    member = db.session.get(Member, member_id)
    if not member:
        flash("Member not found.", "error")
        return redirect(url_for("admin_members"))

    name = (request.form.get("name") or "").strip()
    email = normalize_email(request.form.get("email"))
    member_class = (request.form.get("member_class") or "").strip()
    role = normalize_role(request.form.get("role"))

    if not name or not email or not member_class:
        flash("Name, email, and class/year are required.", "error")
        return redirect(url_for("admin_members"))

    existing = Member.query.filter(db.func.lower(Member.email) == email, Member.id != member.id).first()
    if existing:
        flash("Another member already uses that email.", "error")
        return redirect(url_for("admin_members"))

    member.name = name
    member.email = email
    member.member_class = member_class
    member.role = role
    db.session.commit()
    flash(f"Updated {member.name}.", "success")
    return redirect(url_for("admin_members"))


@app.post("/admin/members/<int:member_id>/set-password")
@admin_required
def admin_member_set_password(member_id):
    member = db.session.get(Member, member_id)
    if not member:
        flash("Member not found.", "error")
        return redirect(url_for("admin_members"))

    password = request.form.get("password") or ""
    if not password:
        flash("Temporary password is required.", "error")
        return redirect(url_for("admin_members"))

    member.password_hash = generate_password_hash(password)
    member.is_active = True
    db.session.commit()
    flash(f"Password updated for {member.name}.", "success")
    return redirect(url_for("admin_members"))


@app.post("/admin/members/<int:member_id>/deactivate")
@admin_required
def admin_member_deactivate(member_id):
    member = db.session.get(Member, member_id)
    if not member:
        flash("Member not found.", "error")
        return redirect(url_for("admin_members"))
    if member.id == current_user.id:
        flash("You cannot deactivate your own admin account.", "error")
        return redirect(url_for("admin_members"))

    member.is_active = False
    db.session.commit()
    flash(f"Deactivated {member.name}.", "info")
    return redirect(url_for("admin_members"))


@app.get("/admin/inventory")
@admin_required
def admin_inventory():
    context = admin_base_context(
        active_page="inventory",
        page_title="Inventory",
        page_subtitle="Track stock, add items, and process equipment checkouts and returns.",
    )
    context.update(
        {
            "members": Member.query.filter_by(is_active=True).order_by(Member.name.asc()).all(),
            "items": Item.query.order_by(Item.name.asc()).all(),
            "active_checkouts": build_active_checkout_lots(),
            "recent_transactions": Transaction.query.order_by(Transaction.timestamp.desc()).limit(20).all(),
            "default_due": str(default_due_date()),
        }
    )
    return render_template("admin/inventory.html", **context)


@app.post("/admin/inventory/add")
@admin_required
def admin_inventory_add():
    name = (request.form.get("name") or "").strip()
    category = (request.form.get("category") or "").strip() or None
    location = (request.form.get("location") or "").strip() or None
    total_qty = parse_int(request.form.get("total_qty"), default=1)
    available_qty = parse_int(request.form.get("available_qty"), default=total_qty)

    if not name:
        flash("Item name is required.", "error")
        return redirect(url_for("admin_inventory"))

    item = Item(
        name=name,
        category=category,
        location=location,
        total_qty=max(1, total_qty),
        available_qty=min(max(1, total_qty), max(0, available_qty)),
    )
    db.session.add(item)
    db.session.commit()
    flash(f"Added inventory item {item.name}.", "success")
    return redirect(url_for("admin_inventory"))


@app.post("/admin/transact")
@admin_required
def admin_transact():
    tx, error = handle_admin_transaction_request()
    if error:
        flash(error, "error")
    else:
        flash(f"{tx.action.title()} saved: {tx.qty} x {tx.item.name} for {tx.member.name}.", "success")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_inventory"))


@app.get("/admin/prints")
@admin_required
def admin_prints():
    context = admin_base_context(
        active_page="prints",
        page_title="Print Queue",
        page_subtitle="Approve pending jobs, monitor printers, and manage the queue.",
    )
    context.update(
        {
            "members": Member.query.filter_by(is_active=True).order_by(Member.name.asc()).all(),
            "printer_types": PRINTER_TYPES,
            "pending_print_jobs": PrintJob.query.filter_by(status="pending").order_by(PrintJob.submitted_at.asc()).all(),
            "queues": get_queue_state(),
        }
    )
    return render_template("admin/prints.html", **context)


@app.post("/admin/prints/submit")
@admin_required
def admin_print_submit():
    member = resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    job, error = create_print_job(
        member,
        request.form.get("printer_type"),
        request.files.get("gcode_file"),
        request.form.get("notes"),
        initial_status="queued",
    )
    if error:
        flash(error, "error")
    else:
        started = dispatch_next_job(job.printer_type)
        if started and started.id == job.id:
            flash(f"{job.printer_type} job submitted and auto-started.", "success")
        else:
            flash(f"{job.printer_type} job submitted directly to queue.", "success")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_prints"))


@app.post("/admin/prints/job/<int:job_id>/accept")
@admin_required
def admin_print_accept(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
    else:
        _, message = accept_print_job(job)
        flash(message, "success" if job.status in {"queued", "printing"} else "error")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_prints"))


@app.post("/admin/prints/job/<int:job_id>/deny")
@admin_required
def admin_print_deny(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
    else:
        result = fail_print_job_record(job, remove_file=True, note="Denied by admin.")
        if result and result["file_error"]:
            flash(f"Denied job #{job.id}, but file deletion failed: {result['file_error'][:200]}", "error")
        else:
            flash(f"Denied job #{job.id}. File removed from queue storage.", "info")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_prints"))


@app.post("/admin/prints/job/<int:job_id>/complete")
@admin_required
def admin_print_complete(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
    else:
        complete_print_job_record(job)
        flash(f"Marked job #{job.id} done. Next queued job auto-started if available.", "success")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_prints"))


@app.post("/admin/prints/job/<int:job_id>/fail")
@admin_required
def admin_print_fail(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
    else:
        fail_print_job_record(job, remove_file=False, note="Marked failed by admin.")
        flash(f"Marked job #{job.id} failed. Next queued job auto-started if available.", "info")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_prints"))


@app.post("/admin/prints/job/<int:job_id>/delete")
@admin_required
def admin_print_delete(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
    elif job.status == "printing":
        flash("Cannot delete an active printing job. Mark it done or failed first.", "error")
    else:
        result = delete_print_job_with_file(job)
        if result["file_error"]:
            flash(f"Deleted job, but file removal failed: {result['file_error'][:200]}", "error")
        elif result["file_removed"]:
            flash(f"Deleted print job and removed {result['file_name']}.", "info")
        else:
            flash(f"Deleted print job for {result['file_name']}. File was already missing.", "info")
    return redirect(safe_redirect_target(request.form.get("next"), "admin_prints"))


@app.get("/admin/calendar")
@admin_required
def admin_calendar():
    ensure_meeting_schema_columns()
    calendar_embed = google_calendar_embed_context()
    context = admin_base_context(
        active_page="calendar",
        page_title="Meetings",
        page_subtitle="Approve room requests, book meetings directly, and manage the live calendar.",
    )
    context.update(
        {
            "meeting_rooms": MEETING_ROOMS,
            "calendar_default_date": str(date.today()),
            "pending_meeting_requests": get_pending_meeting_requests(),
            "confirmed_meetings": get_confirmed_meetings(limit=20),
            "google_calendar_embed_url": calendar_embed["url"],
            "google_calendar_placeholder": calendar_embed["placeholder"],
            "calendar_automation_status": calendar_automation_status(),
        }
    )
    return render_template("admin/calendar.html", **context)


@app.post("/admin/calendar/book")
@admin_required
def admin_calendar_book():
    ensure_meeting_schema_columns()
    payload, error = parse_meeting_request_form()
    if error:
        flash(error, "error")
        return redirect(url_for("admin_calendar"))

    meeting = Meeting(
        team_name=payload["team_name"],
        requester_email=payload["requester_email"] or normalize_email(current_user.email),
        room=payload["room"],
        meeting_date=payload["meeting_date"],
        start_time=payload["start_time"],
        end_time=payload["end_time"],
        notes=payload["notes"],
    )

    calendar_id, event_id, calendar_error = create_google_calendar_event(meeting)
    if calendar_error:
        flash(f"Meeting was not saved: {calendar_error}", "error")
        return redirect(url_for("admin_calendar"))

    meeting.google_calendar_id = calendar_id
    meeting.google_event_id = event_id
    db.session.add(meeting)
    db.session.commit()
    flash("Meeting booked and synced to Google Calendar.", "success")
    return redirect(url_for("admin_calendar"))


@app.post("/admin/calendar/meeting/<int:meeting_id>/approve")
@admin_required
def admin_calendar_approve(meeting_id):
    ensure_meeting_schema_columns()
    meeting = db.session.get(Meeting, meeting_id)
    if not meeting:
        flash("Meeting request not found.", "error")
        return redirect(url_for("admin_calendar"))
    if meeting.google_event_id:
        flash("Meeting is already approved.", "info")
        return redirect(url_for("admin_calendar"))

    conflicting = find_conflicting_meeting(
        meeting.room,
        meeting.meeting_date,
        meeting.start_time,
        meeting.end_time,
        ignore_meeting_id=meeting.id,
    )
    if conflicting:
        flash(
            f"Cannot approve request because {meeting.room} is already booked by {conflicting.team_name}.",
            "error",
        )
        return redirect(url_for("admin_calendar"))

    calendar_id, event_id, calendar_error = create_google_calendar_event(meeting)
    if calendar_error:
        flash(f"Approval failed: {calendar_error}", "error")
        return redirect(url_for("admin_calendar"))

    meeting.google_calendar_id = calendar_id
    meeting.google_event_id = event_id
    db.session.commit()
    flash(f"Approved meeting request for {meeting.team_name}.", "success")
    return redirect(url_for("admin_calendar"))


@app.post("/admin/calendar/meeting/<int:meeting_id>/deny")
@admin_required
def admin_calendar_deny(meeting_id):
    meeting = db.session.get(Meeting, meeting_id)
    if not meeting:
        flash("Meeting request not found.", "error")
    else:
        db.session.delete(meeting)
        db.session.commit()
        flash("Denied and removed the meeting request.", "info")
    return redirect(url_for("admin_calendar"))


@app.post("/admin/calendar/meeting/<int:meeting_id>/cancel")
@admin_required
def admin_calendar_cancel(meeting_id):
    meeting = db.session.get(Meeting, meeting_id)
    if not meeting:
        flash("Meeting not found.", "error")
        return redirect(url_for("admin_calendar"))

    calendar_error = delete_google_calendar_event(meeting)
    if calendar_error:
        flash(f"Cancellation failed: {calendar_error}", "error")
        return redirect(url_for("admin_calendar"))

    db.session.delete(meeting)
    db.session.commit()
    flash("Meeting canceled and removed from Google Calendar.", "info")
    return redirect(url_for("admin_calendar"))


@app.get("/calendar/cancel/confirm/<token>")
def confirm_meeting_cancel(token):
    ensure_meeting_schema_columns()
    meeting = Meeting.query.filter_by(cancel_request_token=token).first()
    if not meeting:
        return (
            "<h3>Cancellation link is invalid or already used.</h3><p>You can close this tab.</p>",
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
    meeting_date_value = meeting.meeting_date.isoformat()
    db.session.delete(meeting)
    db.session.commit()
    return (
        "<h3>Cancellation confirmed.</h3>"
        f"<p>{team_name} in {room} on {meeting_date_value} was removed from Google Calendar.</p>"
        "<p>You can close this tab.</p>"
    )


@app.get("/calendar/cancel/reject/<token>")
def reject_meeting_cancel(token):
    ensure_meeting_schema_columns()
    meeting = Meeting.query.filter_by(cancel_request_token=token).first()
    if not meeting:
        return (
            "<h3>Rejection link is invalid or already used.</h3><p>You can close this tab.</p>",
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


@app.get("/admin/activity")
@admin_required
def admin_activity():
    query_text = (request.args.get("q") or "").strip()
    transaction_query = (
        Transaction.query.join(Member, Transaction.member_id == Member.id)
        .join(Item, Transaction.item_id == Item.id)
        .order_by(Transaction.timestamp.desc())
    )
    if query_text:
        like = f"%{query_text}%"
        transaction_query = transaction_query.filter(
            or_(
                Member.name.ilike(like),
                Member.email.ilike(like),
                Item.name.ilike(like),
                Transaction.action.ilike(like),
                Transaction.notes.ilike(like),
            )
        )

    context = admin_base_context(
        active_page="activity",
        page_title="Activity",
        page_subtitle="Search the full checkout and return history and export operational data.",
    )
    context.update({"recent_transactions": transaction_query.limit(200).all(), "search_query": query_text})
    return render_template("admin/activity.html", **context)


@app.get("/admin/settings")
@admin_required
def admin_settings():
    h2s_print_cmd = (os.environ.get("ASME_H2S_PRINT_CMD") or "").strip()
    p1s_print_cmd = (os.environ.get("ASME_P1S_PRINT_CMD") or "").strip()
    context = admin_base_context(
        active_page="settings",
        page_title="Settings",
        page_subtitle="Operational configuration, exports, and NFC pairing shortcuts.",
    )
    context.update(
        {
            "calendar_automation_status": calendar_automation_status(),
            "h2s_print_cmd_configured": bool(h2s_print_cmd),
            "p1s_print_cmd_configured": bool(p1s_print_cmd),
            "print_commands_env_file": str(PRINT_COMMANDS_ENV_FILE),
        }
    )
    return render_template("admin/settings.html", **context)


@app.get("/admin/export")
@admin_required
def admin_export():
    ensure_meeting_schema_columns()
    members = Member.query.order_by(Member.id.asc()).all()
    items = Item.query.order_by(Item.id.asc()).all()
    transactions = Transaction.query.order_by(Transaction.timestamp.desc()).all()
    scans = AttendanceScan.query.order_by(AttendanceScan.scanned_at.desc()).all()
    jobs = PrintJob.query.order_by(PrintJob.submitted_at.desc()).all()
    meetings = Meeting.query.order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc()).all()

    members_df = pd.DataFrame(
        [
            {
                "id": m.id,
                "name": m.name,
                "email": m.email,
                "class": m.member_class,
                "role": m.role,
                "is_active": m.is_active,
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


@app.route("/admin/pair/member", methods=["GET", "POST"])
@admin_required
def admin_pair_member():
    if request.method == "POST":
        member_id_raw = (request.form.get("member_id") or "").strip()
        tag = (request.form.get("tag") or "").strip()
        if not member_id_raw or not tag:
            flash("Member and UID are required.", "error")
            return redirect(url_for("admin_pair_member"))

        try:
            member_id = int(member_id_raw)
        except Exception:
            flash("Invalid member selection.", "error")
            return redirect(url_for("admin_pair_member"))

        existing = Member.query.filter_by(nfc_tag=tag).first()
        if existing and existing.id != member_id:
            flash("That UID is already assigned to another member.", "error")
            return redirect(url_for("admin_pair_member"))

        member = db.session.get(Member, member_id)
        if not member:
            flash("Member not found.", "error")
            return redirect(url_for("admin_pair_member"))

        member.nfc_tag = tag
        db.session.commit()
        flash(f"Saved UID for {member.name}.", "success")
        return redirect(url_for("admin_pair_member"))

    context = admin_base_context(
        active_page="settings",
        page_title="Pair Member NFC",
        page_subtitle="Attach physical NFC tags to member records.",
    )
    context.update(
        {
            "members": Member.query.filter_by(is_active=True).order_by(Member.name.asc()).all(),
            "paired": Member.query.filter(Member.nfc_tag.isnot(None)).order_by(Member.name.asc()).all(),
        }
    )
    return render_template("admin/pair_member.html", **context)


@app.route("/admin/pair/item", methods=["GET", "POST"])
@admin_required
def admin_pair_item():
    if request.method == "POST":
        item_id_raw = (request.form.get("item_id") or "").strip()
        tag = (request.form.get("tag") or "").strip()
        if not item_id_raw or not tag:
            flash("Item and UID are required.", "error")
            return redirect(url_for("admin_pair_item"))

        try:
            item_id = int(item_id_raw)
        except Exception:
            flash("Invalid item selection.", "error")
            return redirect(url_for("admin_pair_item"))

        existing = Item.query.filter_by(nfc_tag=tag).first()
        if existing and existing.id != item_id:
            flash("That UID is already assigned to another item.", "error")
            return redirect(url_for("admin_pair_item"))

        item = db.session.get(Item, item_id)
        if not item:
            flash("Item not found.", "error")
            return redirect(url_for("admin_pair_item"))

        item.nfc_tag = tag
        db.session.commit()
        flash(f"Saved UID for {item.name}.", "success")
        return redirect(url_for("admin_pair_item"))

    context = admin_base_context(
        active_page="settings",
        page_title="Pair Item NFC",
        page_subtitle="Attach physical NFC tags to inventory bins and tools.",
    )
    context.update(
        {
            "items": Item.query.order_by(Item.name.asc()).all(),
            "paired": Item.query.filter(Item.nfc_tag.isnot(None)).order_by(Item.name.asc()).all(),
        }
    )
    return render_template("admin/pair_item.html", **context)


@app.get("/dashboard")
def dashboard_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("member_portal_entry", next=request.path))
    return redirect(url_for(role_home_endpoint(current_user)))


@app.get("/attendance")
def attendance_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("admin_portal_entry", next=request.path))
    return redirect(url_for("admin_attendance"))


@app.get("/inventory")
def inventory_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("admin_portal_entry", next=request.path))
    return redirect(url_for("admin_inventory"))


@app.get("/prints")
def prints_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("admin_portal_entry", next=request.path))
    return redirect(url_for("admin_prints"))


@app.get("/calendar")
def calendar_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("member_portal_entry", next=request.path))
    return redirect(url_for("admin_calendar" if current_user.role == "admin" else "member_calendar"))


@app.get("/activity")
def activity_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("admin_portal_entry", next=request.path))
    return redirect(url_for("admin_activity"))


@app.get("/pair/member")
def pair_member_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("admin_portal_entry", next=request.path))
    return redirect(url_for("admin_pair_member"))


@app.get("/pair/item")
def pair_item_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("admin_portal_entry", next=request.path))
    return redirect(url_for("admin_pair_item"))


@app.get("/export")
def export_alias():
    if not current_user.is_authenticated:
        return redirect(url_for("admin_portal_entry", next=request.path))
    return redirect(url_for("admin_export"))


@app.post("/transact")
@member_required
def transact_alias():
    if current_user.role == "admin":
        tx, error = handle_admin_transaction_request()
        default_endpoint = "admin_inventory"
    else:
        tx, error = handle_member_transaction_request()
        default_endpoint = "member_checkout"

    if error:
        flash(error, "error")
    else:
        flash(f"{tx.action.title()} saved: {tx.qty} x {tx.item.name} for {tx.member.name}.", "success")
    return redirect(safe_redirect_target(request.form.get("next"), default_endpoint))


@app.post("/print/submit")
@member_required
def print_submit_alias():
    initial_status = "queued" if current_user.role == "admin" else "pending"
    member = (
        resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
        if current_user.role == "admin"
        else current_user
    )
    job, error = create_print_job(
        member,
        request.form.get("printer_type"),
        request.files.get("gcode_file"),
        request.form.get("notes"),
        initial_status=initial_status,
    )

    default_endpoint = "admin_prints" if current_user.role == "admin" else "member_print"
    if error:
        flash(error, "error")
    else:
        if initial_status == "queued":
            started = dispatch_next_job(job.printer_type)
            if started and started.id == job.id:
                flash(f"{job.printer_type} job submitted and auto-started.", "success")
            else:
                flash(f"{job.printer_type} job submitted directly to queue.", "success")
        else:
            flash(f"{job.printer_type} job submitted for admin approval.", "success")
    return redirect(safe_redirect_target(request.form.get("next"), default_endpoint))


@app.get("/print/job/<int:job_id>/download")
@member_required
def download_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job or not can_access_print_job(job):
        flash("Print file not found for that job.", "error")
        return redirect(url_for(role_home_endpoint(current_user)))
    if not os.path.exists(job.file_path):
        flash("Print file not found for that job.", "error")
        return redirect(url_for(role_home_endpoint(current_user)))
    return send_file(job.file_path, as_attachment=True, download_name=job.file_name)


@app.get("/print/job/<int:job_id>/open")
@member_required
def open_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job or not can_access_print_job(job):
        flash("Print file not found for that job.", "error")
        return redirect(url_for(role_home_endpoint(current_user)))
    if not os.path.exists(job.file_path):
        flash("Print file not found for that job.", "error")
        return redirect(url_for(role_home_endpoint(current_user)))

    mime_type, _ = guess_type(job.file_name)
    return send_file(
        job.file_path,
        as_attachment=False,
        download_name=job.file_name,
        mimetype=mime_type or "application/octet-stream",
    )


@app.post("/print/job/<int:job_id>/complete")
@admin_required
def complete_print_job(job_id):
    return admin_print_complete(job_id)


@app.post("/print/job/<int:job_id>/fail")
@admin_required
def fail_print_job(job_id):
    return admin_print_fail(job_id)


@app.post("/print/job/<int:job_id>/delete")
@admin_required
def delete_print_job(job_id):
    return admin_print_delete(job_id)


@app.get("/api/bootstrap")
def api_bootstrap():
    if not current_user.is_authenticated:
        return api_error("Authentication required.", status=401)
    payload = build_admin_bootstrap_payload() if current_user.role == "admin" else build_member_bootstrap_payload(current_user)
    return api_success(payload=payload)


@app.post("/api/attendance/scan")
def api_attendance_scan():
    if not attendance_scan_authorized():
        return api_error("Unauthorized NFC scan.", status=403)
    success, message = scan_attendance_uid(value_from_request("uid", ""))
    if success:
        return api_success(message=message)
    return api_error(message)


@app.post("/api/inventory/transact")
def api_inventory_transact():
    if not current_user.is_authenticated:
        return api_error("Authentication required.", status=401)

    member = (
        resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
        if current_user.role == "admin"
        else current_user
    )
    item = resolve_item(value_from_request("item_tag"), value_from_request("item_id"))
    tx, error = perform_inventory_transaction(
        member,
        item,
        value_from_request("action", ""),
        value_from_request("qty", 1),
        parse_due_date(value_from_request("due_date")),
        value_from_request("notes", ""),
    )
    if error:
        return api_error(error)
    payload = build_admin_bootstrap_payload() if current_user.role == "admin" else build_member_bootstrap_payload(current_user)
    return api_success(message=f"{tx.action.title()} saved: {tx.qty} x {tx.item.name} for {tx.member.name}.", payload=payload)


@app.post("/api/print/submit")
def api_print_submit():
    if not current_user.is_authenticated:
        return api_error("Authentication required.", status=401)

    member = (
        resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
        if current_user.role == "admin"
        else current_user
    )
    initial_status = "queued" if current_user.role == "admin" else "pending"
    job, error = create_print_job(
        member,
        value_from_request("printer_type", ""),
        request.files.get("gcode_file"),
        value_from_request("notes", ""),
        initial_status=initial_status,
    )
    if error:
        return api_error(error)

    if initial_status == "queued":
        started = dispatch_next_job(job.printer_type)
        message = (
            f"{job.printer_type} job submitted and auto-started."
            if started and started.id == job.id
            else f"{job.printer_type} job submitted to queue."
        )
    else:
        message = f"{job.printer_type} job submitted for admin approval."
    payload = build_admin_bootstrap_payload() if current_user.role == "admin" else build_member_bootstrap_payload(current_user)
    return api_success(message=message, payload=payload, status=201)


@app.post("/api/print/job/<int:job_id>/complete")
def api_complete_print_job(job_id):
    if not current_user.is_authenticated or current_user.role != "admin":
        return api_error("Admin access required.", status=403)
    job = db.session.get(PrintJob, job_id)
    if not job:
        return api_error("Print job not found.", status=404)
    complete_print_job_record(job)
    return api_success(message=f"Marked job #{job.id} done.", payload=build_admin_bootstrap_payload())


@app.post("/api/print/job/<int:job_id>/fail")
def api_fail_print_job(job_id):
    if not current_user.is_authenticated or current_user.role != "admin":
        return api_error("Admin access required.", status=403)
    job = db.session.get(PrintJob, job_id)
    if not job:
        return api_error("Print job not found.", status=404)
    fail_print_job_record(job, remove_file=False, note="Marked failed by admin.")
    return api_success(message=f"Marked job #{job.id} failed.", payload=build_admin_bootstrap_payload())


@app.post("/api/print/job/<int:job_id>/delete")
def api_delete_print_job(job_id):
    if not current_user.is_authenticated or current_user.role != "admin":
        return api_error("Admin access required.", status=403)
    job = db.session.get(PrintJob, job_id)
    if not job:
        return api_error("Print job not found.", status=404)
    if job.status == "printing":
        return api_error("Cannot delete an active printing job. Mark it done or failed first.", status=409)
    delete_print_job_with_file(job)
    return api_success(message="Deleted print job.", payload=build_admin_bootstrap_payload())


with app.app_context():
    ensure_database_ready()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
