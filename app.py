import os
import csv
import io
import json
import re
import secrets
import smtplib
import subprocess
import zipfile
import time as time_module
import platform
import unicodedata
from functools import wraps
from datetime import date, datetime, time, timedelta
from email.message import EmailMessage
from mimetypes import guess_type
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urlparse
from urllib.request import Request, urlopen
from uuid import uuid4
from zoneinfo import ZoneInfo

from flask import (
    Flask,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

# Work around a Windows Python 3.13 + SQLAlchemy import hang inside platform.machine().
# Prefer environment-provided CPU arch instead of WMI queries.
if os.name == "nt":
    _win_arch = (
        os.environ.get("PROCESSOR_ARCHITEW6432")
        or os.environ.get("PROCESSOR_ARCHITECTURE")
        or "AMD64"
    )
    platform.machine = lambda: _win_arch

from sqlalchemy import and_, func, inspect, or_, text
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except Exception:  # pragma: no cover - optional dependency fallback for local dev
    service_account = None
    build = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - optional dependency fallback for local dev
    PdfReader = None

from models import (
    Announcement,
    AttendanceRecord,
    AttendanceScan,
    AuditLog,
    ContactMessage,
    Event,
    Item,
    ItemTag,
    Meeting,
    Member,
    NFCTag,
    PasswordResetToken,
    PrintJob,
    PrintRequest,
    Project,
    Transaction,
    User,
    db,
)

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("ASME_DATABASE_URL", "sqlite:///inventory.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SECRET_KEY"] = os.environ.get("ASME_SECRET_KEY", "asme-dev-secret")
db.init_app(app)

PRINTER_TYPES = ("H2S", "P1S")
PRINT_REQUEST_PRINTERS = ("P1S_1", "P1S_2", "P1S_3", "P1S_4", "H2S")
MEETING_ROOMS = ("Robotics Room", "Fluids Lab")
# Supports standard G-code plus Bambu Studio 3MF containers.
ALLOWED_GCODE_EXTENSIONS = {"gcode", "gco", "3mf"}
ALLOWED_RETURN_PHOTO_EXTENSIONS = {"jpg", "jpeg", "png", "webp"}
TAG_PREFIXES = ("item_id:", "item:")
UPLOAD_DIR = Path(app.instance_path) / "gcode_uploads"
RETURN_PHOTO_DIR = Path(app.instance_path) / "return_photos"
PRINT_COMMANDS_ENV_FILE = Path(app.instance_path) / "print_commands.env"
ROSTER_CREDENTIALS_DIR = Path(app.instance_path) / "roster_credentials"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
RETURN_PHOTO_DIR.mkdir(parents=True, exist_ok=True)
ROSTER_CREDENTIALS_DIR.mkdir(parents=True, exist_ok=True)

KNOWN_NON_MS_MAIL_DOMAINS = {
    "gmail.com",
    "yahoo.com",
    "icloud.com",
    "aol.com",
    "proton.me",
    "protonmail.com",
    "gmx.com",
    "mail.com",
}
POSSIBLY_CONSUMER_MS_DOMAINS = {"outlook.com", "hotmail.com", "live.com", "msn.com"}
OUTLOOK_TOKEN_CACHE = {
    "access_token": None,
    "expires_at": 0,
    "tenant_id": "",
    "client_id": "",
}

FRONT_CLUB_MISSION = (
    "ASME at Iowa is a hands-on engineering organization where members design, build, "
    "test, and iterate real mechanical systems while developing leadership and teamwork."
)

FRONT_CLUB_HIGHLIGHTS = [
    {
        "label": "Design + Fabrication",
        "text": "Members move from CAD to manufacturing and validation in real build cycles.",
    },
    {
        "label": "Technical Leadership",
        "text": "Student leads coordinate subsystems, reviews, and project execution timelines.",
    },
    {
        "label": "Industry Readiness",
        "text": "Project workflows mirror engineering practice: requirements, testing, and documentation.",
    },
]

FRONT_PROJECT_SHOWCASE = [
    {
        "team": "Baja Team",
        "name": "Off-Road Vehicle Program",
        "summary": "End-to-end student-built vehicle development with subsystem integration and track testing.",
        "status": "Active Build Season",
    },
    {
        "team": "Formula Team",
        "name": "Formula Design Initiative",
        "summary": "Performance-focused design loops covering chassis, powertrain, controls, and test data analysis.",
        "status": "Prototype + Validation",
    },
    {
        "team": "Design Team",
        "name": "Crater Crusher Platform",
        "summary": "Mission-driven mechanical system development for robust field operation and reliability.",
        "status": "Iteration + Review",
    },
]

FRONT_ABOUT_ASME_FACTS = [
    "ASME is a not-for-profit membership organization focused on collaboration, knowledge sharing, and career development across engineering disciplines.",
    "ASME was founded in 1880 and now includes more than 100,000 members across 140+ countries.",
    "About 32,000 ASME members are students.",
]

FRONT_UIOWA_CURRENT_PROJECTS = [
    {
        "name": "Design Build Fly",
        "summary": (
            "Teams design, fabricate, and demonstrate an unmanned electric radio-controlled aircraft "
            "to meet a defined mission profile."
        ),
    },
    {
        "name": "Additive Manufacturing Mars Rover (R.O.V.E.R.)",
        "summary": (
            "Students use additive manufacturing and iterative design to build an unmanned vehicle "
            "that gathers and deposits resources in an extraterrestrial-style environment."
        ),
    },
    {
        "name": "Automated Garbage Truck",
        "summary": (
            "Student design teams build and test a waste collection system that navigates a model city, "
            "sorts waste streams, and delivers them to the correct destination."
        ),
    },
]

ROLE_ORDER = {"guest": 0, "member": 1, "team_leader": 2, "admin": 3}
PRINT_REQUEST_STATUSES = ("submitted", "approved", "printing", "completed", "rejected")
PASSWORD_RESET_HOURS = 2
ITEM_TYPES = ("tool", "consumable")
INVENTORY_ITEM_CODE_PREFIX = "ASME-INV-"
TREASURY_TRACKER_PREFIX = "ASME-TRK-"
BULK_INVENTORY_ALLOWED_EXTENSIONS = {"csv", "xlsx", "xls"}
EMAIL_RE = re.compile(r"^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$", re.IGNORECASE)
ROSTER_BOOL_WORDS = {"yes", "no"}

LOGIN_RATE_WINDOW_SECONDS = int((os.environ.get("ASME_LOGIN_RATE_WINDOW_SECONDS") or "900").strip())
LOGIN_RATE_MAX_ATTEMPTS = int((os.environ.get("ASME_LOGIN_RATE_MAX_ATTEMPTS") or "8").strip())
LOGIN_RATE_LIMITS = {}
PRINT_REQUEST_MAX_UPLOAD_BYTES = int((os.environ.get("ASME_PRINT_MAX_UPLOAD_MB") or "50").strip()) * 1024 * 1024
CHECKIN_SOON_WINDOW_MINUTES = int((os.environ.get("ASME_CHECKIN_SOON_WINDOW_MINUTES") or "20").strip())
CALENDAR_SCHEDULING_DAYS_DEFAULT = int((os.environ.get("CALENDAR_SCHEDULING_DAYS") or "14").strip())
CALENDAR_WORK_HOURS_START = (os.environ.get("CALENDAR_WORK_HOURS_START") or "08:00").strip()
CALENDAR_WORK_HOURS_END = (os.environ.get("CALENDAR_WORK_HOURS_END") or "22:00").strip()
GOOGLE_CALENDAR_TIMEZONE = (os.environ.get("GOOGLE_CALENDAR_TIMEZONE") or "America/Chicago").strip()
GOOGLE_CAL_SERVICE_CACHE = {"service": None, "source": ""}


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


def parse_env_flag(name, default=False):
    raw = os.environ.get(name)
    if raw is None:
        return bool(default)
    return str(raw).strip().lower() in {"1", "true", "yes", "on"}


def legacy_ops_enabled():
    return parse_env_flag("ASME_ENABLE_LEGACY_OPS", default=False)


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


def normalize_outlook_embed_url(embed_url):
    raw_url = (embed_url or "").strip()
    if not raw_url:
        return ""

    try:
        parsed = urlparse(raw_url)
        host = (parsed.netloc or "").lower()
        path_parts = [part for part in (parsed.path or "").split("/") if part]
        supported_hosts = {"outlook.live.com", "outlook.office.com", "outlook.office365.com"}
        if (
            host in supported_hosts
            and len(path_parts) >= 6
            and path_parts[0] == "owa"
            and path_parts[1] == "calendar"
            and path_parts[-1].lower() == "index.html"
        ):
            owner_id = path_parts[2]
            publish_id = path_parts[3]
            cid = path_parts[4]
            scheme = parsed.scheme or "https"
            return f"{scheme}://{host}/calendar/0/published/{owner_id}/{publish_id}/{cid}/calendar.html/"
    except Exception:
        # Fall back to the original URL if parsing fails.
        pass

    return raw_url


def outlook_calendar_embed_context():
    raw_embed_url = (os.environ.get("ASME_OUTLOOK_CALENDAR_EMBED_URL") or "").strip()
    embed_url = normalize_outlook_embed_url(raw_embed_url)
    if embed_url:
        return {"url": embed_url, "open_url": embed_url, "placeholder": False}
    return {"url": "", "open_url": "", "placeholder": True}


INVENTORY_SCHEMA_READY = False
MEETING_SCHEMA_READY = False
PORTAL_SCHEMA_READY = False


def clean_tag_value(raw):
    value = (raw or "").strip()
    if not value:
        return ""
    return " ".join(value.split())


def normalize_text_key(raw):
    normalized = unicodedata.normalize("NFKD", str(raw or ""))
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "", ascii_only.lower())


def split_name_parts(full_name):
    cleaned = " ".join((full_name or "").strip().split())
    if not cleaned:
        return "", ""
    tokens = cleaned.split(" ")
    first = tokens[0]
    last = " ".join(tokens[1:]) if len(tokens) > 1 else tokens[0]
    return first, last


def username_base_from_name(first_name, last_name):
    first_key = normalize_text_key(first_name)
    last_key = normalize_text_key(last_name)
    base = f"{first_key[:1]}{last_key}"
    if not base:
        base = first_key or last_key or "member"
    return base[:50]


def password_from_name(first_name, last_name):
    first_key = normalize_text_key(first_name)
    last_key = normalize_text_key(last_name)
    prefix = f"{first_key[:2]}{last_key[:1]}"
    if len(prefix) < 3:
        prefix = (prefix + "asx")[:3]
    return f"{prefix}{secrets.randbelow(100000):05d}"


def make_unique_username(base_username, reserved=None, exclude_user_id=None):
    reserved = reserved if reserved is not None else set()
    candidate_base = normalize_text_key(base_username) or "member"
    candidate = candidate_base[:50]
    suffix = 2
    while True:
        duplicate = (
            User.query.filter(
                func.lower(func.coalesce(User.username, "")) == candidate.lower(),
                User.id != (exclude_user_id or 0),
            )
            .order_by(User.id.asc())
            .first()
        )
        if not duplicate and candidate.lower() not in reserved:
            reserved.add(candidate.lower())
            return candidate
        tail = str(suffix)
        candidate = f"{candidate_base[: max(1, 50 - len(tail))]}{tail}"
        suffix += 1


def make_unique_import_email(base_local, reserved_emails=None):
    reserved_emails = reserved_emails if reserved_emails is not None else set()
    local = normalize_text_key(base_local) or "member"
    domain = "asme.local"
    candidate = f"{local}@{domain}"
    suffix = 2
    while True:
        duplicate_user = User.query.filter(func.lower(User.email) == candidate.lower()).first()
        duplicate_member = Member.query.filter(func.lower(Member.email) == candidate.lower()).first()
        if not duplicate_user and not duplicate_member and candidate.lower() not in reserved_emails:
            reserved_emails.add(candidate.lower())
            return candidate
        candidate = f"{local}{suffix}@{domain}"
        suffix += 1


def find_user_by_login_identifier(identifier):
    cleaned = (identifier or "").strip().lower()
    if not cleaned:
        return None
    by_email = User.query.filter(func.lower(User.email) == cleaned).first()
    if by_email:
        return by_email
    return User.query.filter(func.lower(func.coalesce(User.username, "")) == cleaned).first()


def parse_roster_pdf_entries(pdf_bytes):
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed. Add pypdf to requirements and redeploy.")
    reader = PdfReader(io.BytesIO(pdf_bytes))
    rows = []
    seen = set()
    for page in reader.pages:
        text_blob = page.extract_text() or ""
        for raw_line in text_blob.splitlines():
            line = " ".join(raw_line.split()).strip()
            if not line:
                continue
            lowered = line.lower()
            if lowered.startswith("first name last name"):
                continue
            if "engage general teams project teams" in lowered:
                continue
            tokens = line.split(" ")
            email = ""
            for idx in range(len(tokens) - 1, -1, -1):
                maybe_email = tokens[idx].strip(";,")
                if EMAIL_RE.match(maybe_email):
                    email = maybe_email.lower()
                    tokens.pop(idx)
                    break
            while tokens and tokens[-1].lower() in ROSTER_BOOL_WORDS:
                tokens.pop()
            if len(tokens) < 2:
                continue
            first_name = tokens[0]
            last_name = " ".join(tokens[1:])
            full_name = f"{first_name} {last_name}".strip()
            dedupe_key = (normalize_text_key(full_name), email)
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            rows.append(
                {
                    "name": full_name,
                    "first_name": first_name,
                    "last_name": last_name,
                    "email": email,
                }
            )
    return rows


def import_roster_entries(
    entries,
    default_role="member",
    member_class="Member",
    reset_existing_passwords=True,
):
    ensure_portal_schema()
    role_value = normalize_role(default_role or "member")
    if role_value == "admin":
        role_value = "member"
    member_class = (member_class or "Member").strip()[:80] or "Member"

    reserved_usernames = set()
    for row in User.query.order_by(User.id.asc()).all():
        if row.username:
            reserved_usernames.add(row.username.strip().lower())
        elif row.email and "@" in row.email:
            reserved_usernames.add(normalize_text_key(row.email.split("@", 1)[0]))
    reserved_emails = set()
    for row in User.query.order_by(User.id.asc()).all():
        if row.email:
            reserved_emails.add(row.email.strip().lower())
    for row in Member.query.order_by(Member.id.asc()).all():
        if row.email:
            reserved_emails.add(row.email.strip().lower())

    created_members = 0
    created_users = 0
    updated_users = 0
    reset_passwords = 0
    credentials = []

    for entry in entries:
        first_name = entry.get("first_name") or split_name_parts(entry.get("name") or "")[0]
        last_name = entry.get("last_name") or split_name_parts(entry.get("name") or "")[1]
        full_name = " ".join((entry.get("name") or f"{first_name} {last_name}").split()).strip()
        if not full_name:
            continue
        base_username = username_base_from_name(first_name, last_name)

        roster_email = (entry.get("email") or "").strip().lower()
        email = roster_email
        if email:
            reserved_emails.add(email)

        member = None
        if email:
            member = Member.query.filter(func.lower(Member.email) == email).first()
        if not member:
            member = Member.query.filter(func.lower(Member.name) == full_name.lower()).first()
        if not member:
            if not email:
                email = make_unique_import_email(base_username, reserved_emails=reserved_emails)
            member = Member(
                name=full_name[:120],
                email=email[:160],
                member_class=member_class,
            )
            db.session.add(member)
            db.session.flush()
            created_members += 1
        else:
            member.name = full_name[:120]
            member.member_class = member_class
            if not member.email:
                if not email:
                    email = make_unique_import_email(base_username, reserved_emails=reserved_emails)
                member.email = email[:160]
            email = (member.email or email or "").strip().lower()
            if email:
                reserved_emails.add(email)

        if not email:
            email = make_unique_import_email(base_username, reserved_emails=reserved_emails)

        user = User.query.filter(func.lower(User.email) == email).first()
        if not user:
            user = User.query.filter(User.member_id == member.id).first()

        generated_password = password_from_name(first_name, last_name)
        if user:
            username = make_unique_username(base_username, reserved=reserved_usernames, exclude_user_id=user.id)
            user.name = full_name[:160]
            user.email = email[:160]
            user.member_id = member.id
            user.username = username
            user.is_active = True
            if normalize_role(user.role) not in {"admin", "team_leader"}:
                user.role = role_value
            if reset_existing_passwords:
                user.password_hash = generate_password_hash(generated_password)
                reset_passwords += 1
            updated_users += 1
        else:
            username = make_unique_username(base_username, reserved=reserved_usernames)
            user = User(
                name=full_name[:160],
                email=email[:160],
                username=username,
                password_hash=generate_password_hash(generated_password),
                role=role_value,
                is_active=True,
                member_id=member.id,
            )
            db.session.add(user)
            created_users += 1
            reset_passwords += 1

        credentials.append(
            {
                "name": full_name,
                "member_id": member.id,
                "email": email,
                "username": user.username or "",
                "password": generated_password,
            }
        )

    return {
        "created_members": created_members,
        "created_users": created_users,
        "updated_users": updated_users,
        "reset_passwords": reset_passwords,
        "credentials": credentials,
    }


def save_roster_credentials_csv(rows):
    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    filename = f"roster_credentials_{timestamp}.csv"
    output_path = ROSTER_CREDENTIALS_DIR / filename
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["name", "member_id", "email", "username", "password"])
        for row in rows:
            writer.writerow(
                [
                    row.get("name") or "",
                    row.get("member_id") or "",
                    row.get("email") or "",
                    row.get("username") or "",
                    row.get("password") or "",
                ]
            )
    return filename


def find_user_by_nfc_uid(tag_uid, exclude_user_id=None):
    cleaned = clean_tag_value(tag_uid)
    if not cleaned:
        return None
    query = User.query.filter(func.lower(User.nfc_uid) == cleaned.lower())
    if exclude_user_id:
        query = query.filter(User.id != exclude_user_id)
    return query.order_by(User.id.asc()).first()


def resolve_user_from_tag_uid(tag_uid):
    cleaned = clean_tag_value(tag_uid)
    if not cleaned:
        return None
    direct_user = User.query.filter(
        func.lower(User.nfc_uid) == cleaned.lower(),
        User.is_active.is_(True),
    ).first()
    if direct_user:
        return direct_user
    mapped = NFCTag.query.filter(
        func.lower(NFCTag.tag_uid) == cleaned.lower(),
        NFCTag.active.is_(True),
    ).first()
    if mapped and mapped.user and mapped.user.is_active:
        return mapped.user
    return None


def parse_item_id_from_tag(tag):
    raw = clean_tag_value(tag).lower()
    for prefix in TAG_PREFIXES:
        if raw.startswith(prefix):
            try:
                return int(raw.split(":", 1)[1].strip())
            except Exception:
                return None
    return None


def ensure_inventory_schema_columns():
    global INVENTORY_SCHEMA_READY
    if INVENTORY_SCHEMA_READY:
        return

    inspector = inspect(db.engine)
    table_names = inspector.get_table_names()

    if "items" in table_names:
        item_columns = {col["name"] for col in inspector.get_columns("items")}
        alters = []
        if "description" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN description VARCHAR(300)")
        if "category" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN category VARCHAR(80)")
        if "location" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN location VARCHAR(120)")
        if "item_condition" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN item_condition VARCHAR(120)")
        if "notes" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN notes VARCHAR(500)")
        if "photo_url" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN photo_url VARCHAR(500)")
        if "item_type" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN item_type VARCHAR(20)")
        if "is_consumable" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN is_consumable BOOLEAN")
        if "active" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN active BOOLEAN")
        if "min_stock_threshold" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN min_stock_threshold INTEGER")
        if "nfc_tag" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN nfc_tag VARCHAR(120)")
        if "item_code" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN item_code VARCHAR(40)")
        if "treasury_tracker_id" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN treasury_tracker_id VARCHAR(80)")

        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
                conn.execute(
                    text(
                        "UPDATE items SET item_type = COALESCE(NULLIF(item_type, ''), 'tool'), "
                        "is_consumable = COALESCE(is_consumable, CASE WHEN lower(COALESCE(item_type, 'tool')) = 'consumable' THEN TRUE ELSE FALSE END), "
                        "active = COALESCE(active, TRUE), "
                        "min_stock_threshold = COALESCE(min_stock_threshold, 0)"
                        )
                    )

        items_for_tracking = Item.query.order_by(Item.id.asc()).all()
        used_item_codes = set()
        used_tracker_ids = set()
        tracking_changed = False

        for row in items_for_tracking:
            item_code = normalize_inventory_item_code(row.item_code) or generate_inventory_item_code(row.id)
            base_code = item_code
            code_suffix = 2
            while item_code.lower() in used_item_codes:
                item_code = normalize_inventory_item_code(f"{base_code}-{code_suffix}")
                code_suffix += 1
            if row.item_code != item_code:
                row.item_code = item_code
                tracking_changed = True
            used_item_codes.add(item_code.lower())

        for row in items_for_tracking:
            tracker_id = normalize_treasury_tracker_id(row.treasury_tracker_id) or generate_treasury_tracker_id(
                row.item_code,
                row.id,
            )
            base_tracker = tracker_id
            tracker_suffix = 2
            while tracker_id.lower() in used_tracker_ids:
                tracker_id = normalize_treasury_tracker_id(f"{base_tracker}-{tracker_suffix}")
                tracker_suffix += 1
            if row.treasury_tracker_id != tracker_id:
                row.treasury_tracker_id = tracker_id
                tracking_changed = True
            used_tracker_ids.add(tracker_id.lower())

        if tracking_changed:
            db.session.flush()

        db.session.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS ix_items_item_code_unique ON items (item_code)"))
        db.session.execute(
            text(
                "CREATE UNIQUE INDEX IF NOT EXISTS ix_items_treasury_tracker_id_unique "
                "ON items (treasury_tracker_id)"
            )
        )
        db.session.commit()

    if "transactions" in table_names:
        tx_columns = {col["name"] for col in inspector.get_columns("transactions")}
        alters = []
        if "status" not in tx_columns:
            alters.append("ALTER TABLE transactions ADD COLUMN status VARCHAR(20)")
        if "checkout_time" not in tx_columns:
            alters.append("ALTER TABLE transactions ADD COLUMN checkout_time TIMESTAMP")
        if "return_time" not in tx_columns:
            alters.append("ALTER TABLE transactions ADD COLUMN return_time TIMESTAMP")
        if "checkout_notes" not in tx_columns:
            alters.append("ALTER TABLE transactions ADD COLUMN checkout_notes VARCHAR(300)")
        if "return_condition" not in tx_columns:
            alters.append("ALTER TABLE transactions ADD COLUMN return_condition VARCHAR(120)")
        if "return_notes" not in tx_columns:
            alters.append("ALTER TABLE transactions ADD COLUMN return_notes VARCHAR(300)")
        if "return_photo_path" not in tx_columns:
            alters.append("ALTER TABLE transactions ADD COLUMN return_photo_path VARCHAR(500)")

        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))

        with db.engine.begin() as conn:
            conn.execute(
                text(
                    "UPDATE transactions "
                    "SET checkout_time = COALESCE(checkout_time, timestamp), "
                    "checkout_notes = COALESCE(checkout_notes, notes) "
                    "WHERE action = 'checkout'"
                )
            )
            conn.execute(
                text(
                    "UPDATE transactions "
                    "SET return_time = COALESCE(return_time, timestamp), "
                    "return_notes = COALESCE(return_notes, notes), "
                    "status = COALESCE(status, 'RETURNED') "
                    "WHERE action = 'return'"
                )
            )

    ItemTag.__table__.create(bind=db.engine, checkfirst=True)
    db.session.commit()

    # Keep legacy Item.nfc_tag mappings working while enabling multiple tags per item.
    # Use raw SQL here so startup upgrades do not fail when ORM-mapped item columns were
    # added in code but are not yet present in an older SQLite database.
    with db.engine.begin() as conn:
        legacy_items = conn.execute(
            text("SELECT id, nfc_tag FROM items WHERE nfc_tag IS NOT NULL AND TRIM(nfc_tag) <> ''")
        ).fetchall()
    for item_id, raw_tag in legacy_items:
        tag_value = clean_tag_value(raw_tag)
        if not tag_value:
            continue
        existing = ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag_value.lower()).first()
        if existing:
            continue
        db.session.add(ItemTag(item_id=item_id, tag_value=tag_value, source="legacy_item_tag"))
    db.session.commit()

    INVENTORY_SCHEMA_READY = True


def find_item_by_tag(tag):
    cleaned_tag = clean_tag_value(tag)
    if not cleaned_tag:
        return None, "empty"

    tag_item_id = parse_item_id_from_tag(cleaned_tag)
    if tag_item_id:
        item = db.session.get(Item, tag_item_id)
        if item:
            return item, "payload_item_id"

    mapped = ItemTag.query.filter(func.lower(ItemTag.tag_value) == cleaned_tag.lower()).first()
    if mapped and mapped.item:
        return mapped.item, "item_tags"

    legacy_item = Item.query.filter(func.lower(Item.nfc_tag) == cleaned_tag.lower()).first()
    if legacy_item:
        return legacy_item, "legacy_item_nfc_tag"

    return None, "not_found"


def get_open_checkout(member_id, item_id):
    return (
        Transaction.query.filter_by(member_id=member_id, item_id=item_id, status="OUT")
        .order_by(Transaction.checkout_time.desc(), Transaction.id.desc())
        .first()
    )


def get_active_member():
    member_id = session.get("active_member_id")
    if not member_id:
        return None
    try:
        return db.session.get(Member, int(member_id))
    except Exception:
        return None


def get_admin_email_set():
    raw = os.environ.get("ASME_ADMIN_EMAILS", "")
    return {email.strip().lower() for email in raw.split(",") if email.strip()}


def is_admin_member(member):
    if not member:
        return False

    admin_emails = get_admin_email_set()
    if admin_emails and member.email and member.email.strip().lower() in admin_emails:
        return True

    role = (member.member_class or "").strip().lower()
    role_markers = ("admin", "officer", "lead", "president", "chair")
    return any(marker in role for marker in role_markers)


def require_active_member_json():
    member = get_active_member()
    if not member:
        user = current_auth_user()
        if user:
            linked_member = current_user_member(user)
            if linked_member:
                session["active_member_id"] = linked_member.id
                member = linked_member
    if member:
        return member, None
    return None, (
        jsonify(
            {
                "ok": False,
                "error": "Login required. Select your member profile first.",
            }
        ),
        401,
    )


def allowed_return_photo(filename):
    if "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_RETURN_PHOTO_EXTENSIONS


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
    if "outlook_event_id" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN outlook_event_id VARCHAR(180)")
    if "outlook_calendar_id" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN outlook_calendar_id VARCHAR(240)")
    if "cancel_request_token" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN cancel_request_token VARCHAR(120)")
    if "cancel_requested_at" not in existing_columns:
        alters.append("ALTER TABLE meetings ADD COLUMN cancel_requested_at DATETIME")

    if alters:
        with db.engine.begin() as conn:
            for statement in alters:
                conn.execute(text(statement))

    MEETING_SCHEMA_READY = True


def normalize_role(role):
    cleaned = (role or "").strip().lower()
    if cleaned in ROLE_ORDER:
        return cleaned
    return "member"


def role_allows(user_role, required_role):
    return ROLE_ORDER.get(normalize_role(user_role), 0) >= ROLE_ORDER.get(normalize_role(required_role), 0)


def current_auth_user():
    user_id = session.get("auth_user_id")
    if not user_id:
        return None
    try:
        user = db.session.get(User, int(user_id))
    except Exception:
        return None
    if not user or not user.is_active:
        return None
    return user


def current_user_member(user=None):
    user = user or current_auth_user()
    if not user:
        return None
    if user.member:
        return user.member
    if user.email:
        linked = Member.query.filter(func.lower(Member.email) == user.email.strip().lower()).first()
        if linked:
            user.member_id = linked.id
            db.session.commit()
            return linked
    return None


def sign_in_user(user):
    session["auth_user_id"] = user.id
    session["auth_user_role"] = normalize_role(user.role)
    member = current_user_member(user)
    if member:
        session["active_member_id"] = member.id
    else:
        session.pop("active_member_id", None)
    user.last_login_at = datetime.utcnow()
    db.session.commit()


def sign_out_user():
    session.pop("auth_user_id", None)
    session.pop("auth_user_role", None)
    session.pop("active_member_id", None)


def require_login(view_func):
    @wraps(view_func)
    def wrapped(*args, **kwargs):
        user = current_auth_user()
        if not user:
            flash("Please log in to continue.", "error")
            return redirect(url_for("login_page", next=request.path))
        return view_func(*args, **kwargs)

    return wrapped


def require_role(min_role):
    def decorator(view_func):
        @wraps(view_func)
        def wrapped(*args, **kwargs):
            user = current_auth_user()
            if not user:
                flash("Please log in to continue.", "error")
                return redirect(url_for("login_page", next=request.path))
            if not role_allows(user.role, min_role):
                flash("You do not have permission to view that page.", "error")
                return redirect(url_for("portal_router"))
            return view_func(*args, **kwargs)

        return wrapped

    return decorator


def slugify(text_value):
    raw = (text_value or "").strip().lower()
    if not raw:
        return ""
    cleaned = []
    prev_dash = False
    for char in raw:
        if char.isalnum():
            cleaned.append(char)
            prev_dash = False
        elif not prev_dash:
            cleaned.append("-")
            prev_dash = True
    slug = "".join(cleaned).strip("-")
    return slug[:150]


def parse_datetime_local(raw_value):
    value = (raw_value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            continue
    return None


def parse_positive_int(raw_value, default=1):
    try:
        parsed = int(str(raw_value).strip())
        return parsed if parsed > 0 else default
    except Exception:
        return default


def parse_non_negative_int(raw_value, default=0):
    try:
        parsed = int(str(raw_value).strip())
        return parsed if parsed >= 0 else default
    except Exception:
        return default


def parse_bool_flag(raw_value, default=False):
    cleaned = (str(raw_value or "")).strip().lower()
    if not cleaned:
        return default
    if cleaned in {"1", "true", "yes", "on", "y"}:
        return True
    if cleaned in {"0", "false", "no", "off", "n"}:
        return False
    return default


def normalize_inventory_item_code(raw_value):
    cleaned = clean_tag_value(raw_value).upper()
    if not cleaned:
        return ""
    cleaned = re.sub(r"[^A-Z0-9\-]+", "-", cleaned)
    cleaned = re.sub(r"-{2,}", "-", cleaned).strip("-")
    return cleaned[:40]


def normalize_treasury_tracker_id(raw_value):
    cleaned = clean_tag_value(raw_value).upper()
    if not cleaned:
        return ""
    cleaned = re.sub(r"[^A-Z0-9\-]+", "-", cleaned)
    cleaned = re.sub(r"-{2,}", "-", cleaned).strip("-")
    return cleaned[:80]


def generate_inventory_item_code(item_id):
    try:
        numeric_id = int(item_id)
    except Exception:
        numeric_id = 0
    numeric_id = max(numeric_id, 0)
    return f"{INVENTORY_ITEM_CODE_PREFIX}{numeric_id:05d}"


def generate_treasury_tracker_id(item_code, item_id=None):
    normalized_code = normalize_inventory_item_code(item_code) or generate_inventory_item_code(item_id or 0)
    suffix = normalized_code
    if suffix.startswith(INVENTORY_ITEM_CODE_PREFIX):
        suffix = suffix[len(INVENTORY_ITEM_CODE_PREFIX) :]
    return normalize_treasury_tracker_id(f"{TREASURY_TRACKER_PREFIX}{suffix}")


def ensure_item_tracking_ids(item):
    if not item:
        return
    if not item.id:
        db.session.flush()

    item_code = normalize_inventory_item_code(item.item_code) or generate_inventory_item_code(item.id)
    base_code = item_code
    suffix = 2
    while (
        Item.query.filter(
            func.lower(func.coalesce(Item.item_code, "")) == item_code.lower(),
            Item.id != item.id,
        )
        .order_by(Item.id.asc())
        .first()
    ):
        item_code = normalize_inventory_item_code(f"{base_code}-{suffix}")
        suffix += 1
    item.item_code = item_code

    tracker_id = normalize_treasury_tracker_id(item.treasury_tracker_id) or generate_treasury_tracker_id(
        item.item_code, item.id
    )
    base_tracker = tracker_id
    suffix = 2
    while (
        Item.query.filter(
            func.lower(func.coalesce(Item.treasury_tracker_id, "")) == tracker_id.lower(),
            Item.id != item.id,
        )
        .order_by(Item.id.asc())
        .first()
    ):
        tracker_id = normalize_treasury_tracker_id(f"{base_tracker}-{suffix}")
        suffix += 1
    item.treasury_tracker_id = tracker_id


def get_item_primary_tag_map():
    mapping = {}
    rows = ItemTag.query.order_by(ItemTag.id.asc()).all()
    for row in rows:
        cleaned = clean_tag_value(row.tag_value)
        if not cleaned or row.item_id in mapping:
            continue
        mapping[row.item_id] = cleaned
    return mapping


def get_item_primary_nfc(item, tag_map=None):
    if not item:
        return ""
    direct_tag = clean_tag_value(item.nfc_tag)
    if direct_tag:
        return direct_tag
    tag_map = tag_map or {}
    return clean_tag_value(tag_map.get(item.id))


def validate_item_tag_uid(tag_uid, current_item_id=None):
    cleaned = clean_tag_value(tag_uid)
    if not cleaned:
        return ""

    member_conflict = Member.query.filter(func.lower(Member.nfc_tag) == cleaned.lower()).first()
    if member_conflict:
        return "That NFC UID is already assigned to a member profile."

    user_conflict = User.query.filter(func.lower(User.nfc_uid) == cleaned.lower()).first()
    if user_conflict:
        return "That NFC UID is already assigned to a user account."

    nfc_mapping_conflict = NFCTag.query.filter(
        func.lower(NFCTag.tag_uid) == cleaned.lower(),
        NFCTag.active.is_(True),
    ).first()
    if nfc_mapping_conflict:
        return "That NFC UID is already assigned to a member login tag."

    item_conflict = Item.query.filter(
        func.lower(func.coalesce(Item.nfc_tag, "")) == cleaned.lower(),
        Item.id != (current_item_id or 0),
    ).first()
    if item_conflict:
        return "That NFC UID is already assigned to another inventory item."

    item_tag_conflict = ItemTag.query.filter(
        func.lower(ItemTag.tag_value) == cleaned.lower(),
        ItemTag.item_id != (current_item_id or 0),
    ).first()
    if item_tag_conflict:
        return "That NFC UID is already assigned in the item tag map."

    return ""


def set_item_primary_nfc(item, tag_uid, source="inventory_admin"):
    cleaned = clean_tag_value(tag_uid)
    item.nfc_tag = cleaned or None
    if not cleaned:
        return

    existing = ItemTag.query.filter(func.lower(ItemTag.tag_value) == cleaned.lower()).first()
    if existing:
        existing.item_id = item.id
        if not existing.source:
            existing.source = source[:40]
        return
    db.session.add(ItemTag(item_id=item.id, tag_value=cleaned, source=source[:40]))


def csv_stream_from_rows(headers, rows):
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(headers)
    for row in rows:
        writer.writerow(row)
    return buffer.getvalue()


def bulk_inventory_row_value(row, *aliases):
    normalized = {}
    for key, value in (row or {}).items():
        normalized[normalize_text_key(key)] = value
    for alias in aliases:
        candidate = normalized.get(normalize_text_key(alias))
        if candidate is None:
            continue
        text_value = str(candidate).strip()
        if text_value:
            return text_value
    return ""


def parse_bulk_inventory_rows(upload_file):
    if not upload_file or not upload_file.filename:
        raise ValueError("Upload a CSV or Excel file first.")
    filename = secure_filename(upload_file.filename)
    if "." not in filename:
        raise ValueError("Unsupported file type. Use CSV or Excel.")
    extension = filename.rsplit(".", 1)[1].lower()
    if extension not in BULK_INVENTORY_ALLOWED_EXTENSIONS:
        raise ValueError("Unsupported file type. Use CSV, XLSX, or XLS.")

    if extension == "csv":
        raw_bytes = upload_file.read()
        if not raw_bytes:
            raise ValueError("The uploaded file is empty.")
        try:
            text_blob = raw_bytes.decode("utf-8-sig")
        except Exception:
            text_blob = raw_bytes.decode("latin-1", errors="ignore")
        reader = csv.DictReader(io.StringIO(text_blob))
        if not reader.fieldnames:
            raise ValueError("CSV file must include a header row.")
        return [dict(row) for row in reader]

    try:
        import pandas as pd
    except Exception as exc:
        raise ValueError("Excel import requires pandas and openpyxl in requirements.") from exc

    upload_file.stream.seek(0)
    frame = pd.read_excel(upload_file)
    if frame.empty:
        raise ValueError("The uploaded file is empty.")
    frame = frame.fillna("")
    return frame.to_dict(orient="records")


def normalize_portal_printer_type(raw_value):
    cleaned = (raw_value or "").strip().upper().replace("-", "_")
    if cleaned in PRINT_REQUEST_PRINTERS:
        return cleaned
    return ""


def parse_json_list(raw_value):
    text_value = (raw_value or "").strip()
    if not text_value:
        return []
    try:
        parsed = json.loads(text_value)
        if isinstance(parsed, list):
            return [str(item).strip() for item in parsed if str(item).strip()]
    except Exception:
        pass
    lines = [line.strip() for line in text_value.splitlines() if line.strip()]
    return lines


def redirect_to_next(default_endpoint):
    next_url = (request.form.get("next") or request.args.get("next") or "").strip()
    if next_url.startswith("/"):
        return redirect(next_url)
    return redirect(url_for(default_endpoint))


def request_client_ip():
    forwarded = (request.headers.get("X-Forwarded-For") or "").strip()
    if forwarded:
        return forwarded.split(",")[0].strip()[:120]
    return (request.remote_addr or "unknown")[:120]


def login_rate_key(email):
    return f"{request_client_ip()}|{(email or '').strip().lower()}"


def clear_expired_login_limits():
    now = time_module.time()
    expired = [key for key, info in LOGIN_RATE_LIMITS.items() if now > info.get("reset_at", 0)]
    for key in expired:
        LOGIN_RATE_LIMITS.pop(key, None)


def record_login_failure(email):
    clear_expired_login_limits()
    now = time_module.time()
    key = login_rate_key(email)
    row = LOGIN_RATE_LIMITS.get(key)
    if not row or now > row.get("reset_at", 0):
        LOGIN_RATE_LIMITS[key] = {"count": 1, "reset_at": now + LOGIN_RATE_WINDOW_SECONDS}
        return
    row["count"] = int(row.get("count", 0)) + 1


def clear_login_failures(email):
    LOGIN_RATE_LIMITS.pop(login_rate_key(email), None)


def is_login_rate_limited(email):
    clear_expired_login_limits()
    row = LOGIN_RATE_LIMITS.get(login_rate_key(email))
    if not row:
        return False, 0
    if int(row.get("count", 0)) < LOGIN_RATE_MAX_ATTEMPTS:
        return False, 0
    retry_after = int(max(1, row.get("reset_at", 0) - time_module.time()))
    return True, retry_after


def google_calendar_ids_by_room():
    robotics = (os.environ.get("GOOGLE_CALENDAR_ID_ROBOTICS") or "").strip()
    fluids = (os.environ.get("GOOGLE_CALENDAR_ID_FLUIDS") or "").strip()
    return {
        "Robotics Room": robotics,
        "Fluids Lab": fluids,
    }


def calendar_tzinfo():
    try:
        return ZoneInfo(GOOGLE_CALENDAR_TIMEZONE)
    except Exception:
        return ZoneInfo("UTC")


def parse_work_hours():
    start = parse_clock_time(CALENDAR_WORK_HOURS_START) or time(hour=8, minute=0)
    end = parse_clock_time(CALENDAR_WORK_HOURS_END) or time(hour=22, minute=0)
    if end <= start:
        end = time(hour=22, minute=0)
    return start, end


def google_calendar_config():
    provider = calendar_provider_name()
    room_ids = google_calendar_ids_by_room()
    errors = []
    if provider != "google":
        errors.append("ASME_CALENDAR_PROVIDER is not set to google.")
    if not room_ids.get("Robotics Room"):
        errors.append("GOOGLE_CALENDAR_ID_ROBOTICS is not configured.")
    if not room_ids.get("Fluids Lab"):
        errors.append("GOOGLE_CALENDAR_ID_FLUIDS is not configured.")
    service_raw = (os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") or "").strip()
    if not service_raw:
        errors.append("GOOGLE_SERVICE_ACCOUNT_JSON is not configured.")
    return {
        "provider": provider,
        "room_ids": room_ids,
        "service_raw": service_raw,
        "enabled": len(errors) == 0,
        "errors": errors,
    }


def load_google_service_account_info():
    config = google_calendar_config()
    raw = config["service_raw"]
    if not raw:
        return None
    maybe_path = Path(raw)
    if maybe_path.exists():
        try:
            return json.loads(maybe_path.read_text(encoding="utf-8"))
        except Exception:
            return None
    try:
        return json.loads(raw)
    except Exception:
        return None


def google_calendar_service():
    if service_account is None or build is None:
        return None
    config = google_calendar_config()
    source_key = config["service_raw"]
    if not source_key:
        return None
    cached = GOOGLE_CAL_SERVICE_CACHE.get("service")
    if cached and GOOGLE_CAL_SERVICE_CACHE.get("source") == source_key:
        return cached
    info = load_google_service_account_info()
    if not info:
        return None
    try:
        creds = service_account.Credentials.from_service_account_info(
            info,
            scopes=["https://www.googleapis.com/auth/calendar"],
        )
        service = build("calendar", "v3", credentials=creds, cache_discovery=False)
    except Exception:
        return None
    GOOGLE_CAL_SERVICE_CACHE["service"] = service
    GOOGLE_CAL_SERVICE_CACHE["source"] = source_key
    return service


def build_google_embed_url_from_room_ids():
    room_ids = google_calendar_ids_by_room()
    ids = [calendar_id for calendar_id in room_ids.values() if calendar_id]
    if not ids:
        return ""
    query = [("ctz", GOOGLE_CALENDAR_TIMEZONE), ("mode", "WEEK"), ("showTitle", "0"), ("showPrint", "0")]
    for calendar_id in ids:
        query.append(("src", calendar_id))
    return "https://calendar.google.com/calendar/embed?" + urlencode(query)


def overlap_exists(start_a, end_a, start_b, end_b):
    return start_a < end_b and end_a > start_b


def fetch_google_busy_windows(time_min, time_max):
    service = google_calendar_service()
    config = google_calendar_config()
    if not service or not config["enabled"]:
        return None, "Google Calendar integration is not configured."
    room_ids = config["room_ids"]
    body = {
        "timeMin": time_min.isoformat(),
        "timeMax": time_max.isoformat(),
        "timeZone": GOOGLE_CALENDAR_TIMEZONE,
        "items": [{"id": room_ids["Robotics Room"]}, {"id": room_ids["Fluids Lab"]}],
    }
    try:
        response = service.freebusy().query(body=body).execute()
    except Exception as exc:
        return None, f"Google freebusy request failed: {exc}"

    tzinfo = calendar_tzinfo()
    busy_by_room = {room: [] for room in MEETING_ROOMS}
    for room in MEETING_ROOMS:
        calendar_id = room_ids.get(room)
        payload = ((response.get("calendars") or {}).get(calendar_id) or {}) if calendar_id else {}
        entries = payload.get("busy") or []
        for entry in entries:
            try:
                start_dt = datetime.fromisoformat((entry.get("start") or "").replace("Z", "+00:00")).astimezone(tzinfo)
                end_dt = datetime.fromisoformat((entry.get("end") or "").replace("Z", "+00:00")).astimezone(tzinfo)
            except Exception:
                continue
            busy_by_room[room].append((start_dt, end_dt))
    return busy_by_room, None


def compute_available_slots(duration_minutes, day_count, preferred_room=""):
    day_count = max(1, min(day_count, 30))
    duration_minutes = max(30, min(duration_minutes, 180))
    duration_delta = timedelta(minutes=duration_minutes)
    tzinfo = calendar_tzinfo()
    now = datetime.now(tzinfo)
    work_start, work_end = parse_work_hours()
    today = now.date()
    end_day = today + timedelta(days=day_count - 1)
    range_start = datetime.combine(today, work_start, tzinfo=tzinfo)
    range_end = datetime.combine(end_day, work_end, tzinfo=tzinfo)

    busy_by_room, busy_error = fetch_google_busy_windows(range_start, range_end)
    if busy_error:
        return [], busy_error
    rooms = [preferred_room] if preferred_room in MEETING_ROOMS else list(MEETING_ROOMS)
    slots = []
    for day_offset in range(day_count):
        day_value = today + timedelta(days=day_offset)
        for room in rooms:
            day_start = datetime.combine(day_value, work_start, tzinfo=tzinfo)
            day_end = datetime.combine(day_value, work_end, tzinfo=tzinfo)
            cursor = day_start
            while cursor + duration_delta <= day_end:
                slot_start = cursor
                slot_end = cursor + duration_delta
                cursor += timedelta(minutes=30)
                if slot_start < now + timedelta(minutes=2):
                    continue
                room_busy = busy_by_room.get(room, [])
                if any(overlap_exists(slot_start, slot_end, busy_start, busy_end) for busy_start, busy_end in room_busy):
                    continue
                local_start = slot_start.replace(tzinfo=None)
                local_end = slot_end.replace(tzinfo=None)
                if find_event_room_conflict(room, local_start, local_end):
                    continue
                slots.append(
                    {
                        "room": room,
                        "start": slot_start,
                        "end": slot_end,
                        "token": f"{room}|{slot_start.strftime('%Y-%m-%dT%H:%M')}|{duration_minutes}",
                    }
                )
    return slots, None


def create_google_meeting_event(room, team_name, organizer_name, start_local, end_local, notes=""):
    config = google_calendar_config()
    service = google_calendar_service()
    if not config["enabled"] or not service:
        return None, None, "Google Calendar integration is not configured."
    calendar_id = config["room_ids"].get(room)
    if not calendar_id:
        return None, None, f"No Google calendar ID configured for {room}."
    tz_name = GOOGLE_CALENDAR_TIMEZONE
    summary = f"ASME - {team_name} Meeting"
    details = [f"Organizer: {organizer_name}"]
    if notes:
        details.append(notes)
    body = {
        "summary": summary[:220],
        "description": "\n".join(details)[:4000],
        "location": room,
        "start": {"dateTime": start_local.isoformat(), "timeZone": tz_name},
        "end": {"dateTime": end_local.isoformat(), "timeZone": tz_name},
    }
    try:
        payload = service.events().insert(calendarId=calendar_id, body=body).execute()
    except Exception as exc:
        return None, None, f"Google event create failed: {exc}"
    return (payload.get("id") or None), (payload.get("htmlLink") or None), None


def portal_calendar_embed_url():
    google_embed = (os.environ.get("ASME_GOOGLE_CALENDAR_EMBED_URL") or "").strip()
    if google_embed:
        return google_embed
    generated_google_embed = build_google_embed_url_from_room_ids()
    if generated_google_embed:
        return generated_google_embed
    outlook_embed = (os.environ.get("ASME_OUTLOOK_CALENDAR_EMBED_URL") or "").strip()
    if outlook_embed:
        return normalize_outlook_embed_url(outlook_embed)
    return ""


def calendar_provider_name():
    provider = (os.environ.get("ASME_CALENDAR_PROVIDER") or "google").strip().lower()
    if provider not in {"google", "outlook"}:
        provider = "google"
    return provider


def normalize_meeting_room(raw_location):
    location = (raw_location or "").strip()
    if not location:
        return ""
    lookup = {room.lower(): room for room in MEETING_ROOMS}
    return lookup.get(location.lower(), "")


def find_event_room_conflict(room, start_time, end_time, exclude_event_id=None):
    if not room:
        return None
    query = Event.query.filter(
        func.lower(func.coalesce(Event.location, "")) == room.lower(),
        Event.status.in_(["requested", "scheduled"]),
        Event.start_time < end_time,
        Event.end_time > start_time,
    )
    if exclude_event_id:
        query = query.filter(Event.id != exclude_event_id)
    return query.order_by(Event.start_time.asc(), Event.id.asc()).first()


def current_checkin_candidate_events():
    now = datetime.now()
    soon = now + timedelta(minutes=CHECKIN_SOON_WINDOW_MINUTES)
    return (
        Event.query.filter(
            Event.status == "scheduled",
            Event.end_time >= now,
            Event.start_time <= soon,
        )
        .order_by(Event.start_time.asc(), Event.id.asc())
        .all()
    )


def check_user_into_event(user, event, method="shared_tag", tag_uid=None):
    existing = (
        AttendanceRecord.query.filter(
            AttendanceRecord.event_id == event.id,
            AttendanceRecord.user_id == user.id,
        )
        .order_by(AttendanceRecord.id.desc())
        .first()
    )
    if existing:
        return existing, False
    member = current_user_member(user)
    row = AttendanceRecord(
        event_id=event.id,
        user_id=user.id,
        member_id=member.id if member else None,
        tag_uid=clean_tag_value(tag_uid) or None,
        checkin_method=(method or "shared_tag")[:40],
        checkin_time=datetime.utcnow(),
    )
    db.session.add(row)
    db.session.commit()
    return row, True


def add_audit_log(action, details=""):
    user = current_auth_user()
    log = AuditLog(
        admin_user_id=user.id if user else None,
        action=(action or "").strip()[:160] or "action",
        details=(details or "").strip()[:4000] or None,
        ip_address=(request.remote_addr or "")[:120] or None,
    )
    db.session.add(log)


def ensure_portal_schema():
    global PORTAL_SCHEMA_READY
    if PORTAL_SCHEMA_READY:
        return

    db.create_all()
    inspector = inspect(db.engine)
    table_names = inspector.get_table_names()

    if "users" in table_names:
        user_columns = {col["name"] for col in inspector.get_columns("users")}
        alters = []
        if "username" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN username VARCHAR(80)")
        if "role" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN role VARCHAR(30)")
        if "is_active" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN is_active BOOLEAN")
        if "nfc_uid" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN nfc_uid VARCHAR(160)")
        if "major" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN major VARCHAR(120)")
        if "graduation_year" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN graduation_year INTEGER")
        if "exec_title" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN exec_title VARCHAR(160)")
        if "exec_message" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN exec_message VARCHAR(500)")
        if "headshot_url" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN headshot_url VARCHAR(500)")
        if "member_id" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN member_id INTEGER")
        if "last_login_at" not in user_columns:
            alters.append("ALTER TABLE users ADD COLUMN last_login_at DATETIME")
        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
        with db.engine.begin() as conn:
            try:
                conn.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_nfc_uid_unique ON users(nfc_uid)"))
                conn.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_unique ON users(username)"))
            except Exception:
                # Keep startup resilient on legacy DBs; uniqueness is also enforced in admin routes.
                pass
        with db.engine.begin() as conn:
            conn.execute(text("UPDATE users SET role = COALESCE(role, 'member')"))
            conn.execute(text("UPDATE users SET is_active = COALESCE(is_active, TRUE)"))

        users_missing_username = User.query.filter(
            or_(User.username.is_(None), func.trim(User.username) == "")
        ).order_by(User.id.asc()).all()
        reserved_usernames = {
            (row.username or "").strip().lower()
            for row in User.query.filter(User.username.isnot(None)).all()
            if (row.username or "").strip()
        }
        for row in users_missing_username:
            base = ""
            if row.email and "@" in row.email:
                base = row.email.split("@", 1)[0]
            elif row.name:
                first_name, last_name = split_name_parts(row.name)
                base = username_base_from_name(first_name, last_name)
            row.username = make_unique_username(base or f"user{row.id}", reserved=reserved_usernames, exclude_user_id=row.id)

    if "items" in table_names:
        item_columns = {col["name"] for col in inspector.get_columns("items")}
        alters = []
        if "item_condition" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN item_condition VARCHAR(120)")
        if "notes" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN notes VARCHAR(500)")
        if "photo_url" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN photo_url VARCHAR(500)")
        if "item_type" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN item_type VARCHAR(20)")
        if "is_consumable" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN is_consumable BOOLEAN")
        if "active" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN active BOOLEAN")
        if "min_stock_threshold" not in item_columns:
            alters.append("ALTER TABLE items ADD COLUMN min_stock_threshold INTEGER")
        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
                conn.execute(
                    text(
                        "UPDATE items SET item_type = COALESCE(NULLIF(item_type, ''), 'tool'), "
                        "is_consumable = COALESCE(is_consumable, CASE WHEN lower(COALESCE(item_type, 'tool')) = 'consumable' THEN TRUE ELSE FALSE END), "
                        "active = COALESCE(active, TRUE), "
                        "min_stock_threshold = COALESCE(min_stock_threshold, 0)"
                    )
                )

    if "transactions" in table_names:
        tx_columns = {col["name"] for col in inspector.get_columns("transactions")}
        if "user_id" not in tx_columns:
            with db.engine.begin() as conn:
                conn.execute(text("ALTER TABLE transactions ADD COLUMN user_id INTEGER"))

    # Backfill transaction.user_id where possible via member email match.
    pending_user_links = Transaction.query.filter(Transaction.user_id.is_(None)).all()
    for tx in pending_user_links:
        if not tx.member or not tx.member.email:
            continue
        matched_user = User.query.filter(func.lower(User.email) == tx.member.email.strip().lower()).first()
        if matched_user:
            tx.user_id = matched_user.id

    if "events" in table_names:
        event_columns = {col["name"] for col in inspector.get_columns("events")}
        alters = []
        if "status" not in event_columns:
            alters.append("ALTER TABLE events ADD COLUMN status VARCHAR(40)")
        if "requested_by_user_id" not in event_columns:
            alters.append("ALTER TABLE events ADD COLUMN requested_by_user_id INTEGER")
        if "google_event_id" not in event_columns:
            alters.append("ALTER TABLE events ADD COLUMN google_event_id VARCHAR(220)")
        if "google_calendar_id" not in event_columns:
            alters.append("ALTER TABLE events ADD COLUMN google_calendar_id VARCHAR(260)")
        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
                conn.execute(text("UPDATE events SET status = COALESCE(status, 'scheduled')"))

    if "print_requests" in table_names:
        pr_columns = {col["name"] for col in inspector.get_columns("print_requests")}
        alters = []
        if "filament" not in pr_columns:
            alters.append("ALTER TABLE print_requests ADD COLUMN filament VARCHAR(160)")
        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
                conn.execute(
                    text(
                        "UPDATE print_requests SET filament = COALESCE(NULLIF(filament, ''), "
                        "TRIM(COALESCE(material, '') || CASE WHEN COALESCE(color, '') <> '' THEN ' / ' || color ELSE '' END))"
                    )
                )

    if "projects" in table_names:
        project_columns = {col["name"] for col in inspector.get_columns("projects")}
        alters = []
        if "project_type" not in project_columns:
            alters.append("ALTER TABLE projects ADD COLUMN project_type VARCHAR(80)")
        if "timeline" not in project_columns:
            alters.append("ALTER TABLE projects ADD COLUMN timeline TEXT")
        if "gallery_json" not in project_columns:
            alters.append("ALTER TABLE projects ADD COLUMN gallery_json TEXT")
        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
                conn.execute(text("UPDATE projects SET project_type = COALESCE(project_type, title)"))

    if "contact_messages" in table_names:
        contact_columns = {col["name"] for col in inspector.get_columns("contact_messages")}
        alters = []
        if "kind" not in contact_columns:
            alters.append("ALTER TABLE contact_messages ADD COLUMN kind VARCHAR(40)")
        if "user_id" not in contact_columns:
            alters.append("ALTER TABLE contact_messages ADD COLUMN user_id INTEGER")
        if "target" not in contact_columns:
            alters.append("ALTER TABLE contact_messages ADD COLUMN target VARCHAR(80)")
        if "admin_reply" not in contact_columns:
            alters.append("ALTER TABLE contact_messages ADD COLUMN admin_reply TEXT")
        if "updated_at" not in contact_columns:
            alters.append("ALTER TABLE contact_messages ADD COLUMN updated_at DATETIME")
        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
                conn.execute(text("UPDATE contact_messages SET kind = COALESCE(kind, 'contact')"))
                conn.execute(text("UPDATE contact_messages SET updated_at = COALESCE(updated_at, created_at)"))

    if "announcements" in table_names:
        announcement_columns = {col["name"] for col in inspector.get_columns("announcements")}
        alters = []
        if "show_on_public" not in announcement_columns:
            alters.append("ALTER TABLE announcements ADD COLUMN show_on_public BOOLEAN")
        if "show_on_member" not in announcement_columns:
            alters.append("ALTER TABLE announcements ADD COLUMN show_on_member BOOLEAN")
        if alters:
            with db.engine.begin() as conn:
                for statement in alters:
                    conn.execute(text(statement))
                conn.execute(text("UPDATE announcements SET show_on_public = COALESCE(show_on_public, TRUE)"))
                conn.execute(text("UPDATE announcements SET show_on_member = COALESCE(show_on_member, TRUE)"))

    # Ensure all portal tables exist even for upgraded existing DBs.
    for model in (
        User,
        NFCTag,
        Project,
        ContactMessage,
        PrintRequest,
        Event,
        AttendanceRecord,
        Announcement,
        AuditLog,
        PasswordResetToken,
    ):
        model.__table__.create(bind=db.engine, checkfirst=True)

    seed_portal_data()
    PORTAL_SCHEMA_READY = True


def seed_portal_data():
    default_password = (os.environ.get("ASME_DEFAULT_USER_PASSWORD") or "ChangeMe123!").strip()
    default_admin_email = (os.environ.get("ASME_DEFAULT_ADMIN_EMAIL") or "admin@uiowa.edu").strip().lower()
    default_admin_password = (os.environ.get("ASME_DEFAULT_ADMIN_PASSWORD") or "ChangeMe123!").strip()

    if Project.query.count() == 0:
        defaults = [
            {
                "slug": "rover",
                "title": "Rover",
                "project_type": "Rover",
                "summary": "Mobility-focused platform with drivetrain, controls, and testing milestones.",
                "description": (
                    "The Rover team develops a rugged platform for terrain handling, control stability, "
                    "and subsystem validation through iterative build cycles."
                ),
                "status": "Active",
                "timeline": "Concept -> CAD -> Fabrication -> Integration -> Field Test",
                "gallery_json": json.dumps([]),
                "lead_name": "Rover Lead",
                "image_url": "",
                "external_link": "",
            },
            {
                "slug": "arm",
                "title": "Robotic Arm",
                "project_type": "Arm",
                "summary": "Manipulator design integrating structure, actuators, and controls workflows.",
                "description": (
                    "The Arm project focuses on payload handling, repeatability, and manufacturing-ready "
                    "component design for reliable operation."
                ),
                "status": "Prototype",
                "timeline": "Kinematics Study -> Linkage Design -> Controls Tuning -> Validation",
                "gallery_json": json.dumps([]),
                "lead_name": "Controls Lead",
                "image_url": "",
                "external_link": "",
            },
            {
                "slug": "manufacturing",
                "title": "Manufacturing",
                "project_type": "Manufacturing",
                "summary": "CAD-to-fabrication process ownership and quality-first part production.",
                "description": (
                    "The Manufacturing track supports all project teams with machining plans, print strategy, "
                    "and documentation for production consistency."
                ),
                "status": "In Progress",
                "timeline": "Manufacturing Plans -> Material Prep -> Production -> QA",
                "gallery_json": json.dumps([]),
                "lead_name": "Manufacturing Lead",
                "image_url": "",
                "external_link": "",
            },
        ]
        for payload in defaults:
            db.session.add(Project(**payload))

    if Announcement.query.count() == 0:
        db.session.add(
            Announcement(
                title="Welcome to ASME @ UIowa",
                body="Spring build season is active. Check project boards and weekly meeting updates.",
                is_published=True,
                show_on_public=True,
                show_on_member=True,
                published_at=datetime.utcnow(),
            )
        )

    if User.query.count() == 0:
        members = Member.query.order_by(Member.id.asc()).all()
        for member in members:
            role = "member"
            role_hint = (member.member_class or "").lower()
            if "lead" in role_hint:
                role = "team_leader"
            if is_admin_member(member):
                role = "admin"
            db.session.add(
                User(
                    name=member.name,
                    email=member.email,
                    username=(member.email.split("@", 1)[0] if member.email and "@" in member.email else None),
                    password_hash=generate_password_hash(default_password),
                    role=role,
                    is_active=True,
                    member_id=member.id,
                )
            )

        if not members:
            db.session.add(
                User(
                    name="ASME Admin",
                    email=default_admin_email,
                    username=(default_admin_email.split("@", 1)[0] if "@" in default_admin_email else "admin"),
                    password_hash=generate_password_hash(default_admin_password),
                    role="admin",
                    is_active=True,
                )
            )
    else:
        existing_admin = User.query.filter(func.lower(User.email) == default_admin_email).first()
        if not existing_admin:
            db.session.add(
                User(
                    name="ASME Admin",
                    email=default_admin_email,
                    username=(default_admin_email.split("@", 1)[0] if "@" in default_admin_email else "admin"),
                    password_hash=generate_password_hash(default_admin_password),
                    role="admin",
                    is_active=True,
                )
            )

    db.session.commit()


def get_shared_admin_credentials():
    email = (
        os.environ.get("ASME_DEFAULT_ADMIN_EMAIL")
        or os.environ.get("ASME_SHARED_ADMIN_EMAIL")
        or os.environ.get("ASME_ADMIN_EMAIL")
        or ""
    ).strip().lower()
    password = (
        os.environ.get("ASME_DEFAULT_ADMIN_PASSWORD")
        or os.environ.get("ASME_SHARED_ADMIN_PASSWORD")
        or os.environ.get("ASME_ADMIN_PASSWORD")
        or ""
    ).strip()
    return email, password


def ensure_shared_admin_user(email, password):
    if not email or not password:
        return None
    user = User.query.filter(func.lower(User.email) == email).first()
    if not user:
        user = User(
            name="ASME Admin",
            email=email,
            username=(email.split("@", 1)[0] if "@" in email else "admin"),
            password_hash=generate_password_hash(password),
            role="admin",
            is_active=True,
        )
        db.session.add(user)
    else:
        user.role = "admin"
        user.is_active = True
        if not check_password_hash(user.password_hash, password):
            user.password_hash = generate_password_hash(password)
        if not (user.username or "").strip():
            user.username = make_unique_username(email.split("@", 1)[0] if "@" in email else f"admin{user.id}")
    db.session.commit()
    return user


def get_calendar_timezone():
    return (os.environ.get("ASME_OUTLOOK_CALENDAR_TZ") or "Central Standard Time").strip()


def get_outlook_calendar_id_for_room(room):
    robotics_calendar_id = (os.environ.get("ASME_OUTLOOK_CALENDAR_ROBOTICS_ID") or "").strip()
    fluids_calendar_id = (os.environ.get("ASME_OUTLOOK_CALENDAR_FLUIDS_ID") or "").strip()
    default_calendar_id = (os.environ.get("ASME_OUTLOOK_CALENDAR_ID") or "").strip()

    if room == "Robotics Room" and robotics_calendar_id:
        return robotics_calendar_id
    if room == "Fluids Lab" and fluids_calendar_id:
        return fluids_calendar_id
    return default_calendar_id


def get_outlook_oauth_config():
    tenant_id = (os.environ.get("ASME_OUTLOOK_TENANT_ID") or "").strip()
    client_id = (os.environ.get("ASME_OUTLOOK_CLIENT_ID") or "").strip()
    client_secret = (os.environ.get("ASME_OUTLOOK_CLIENT_SECRET") or "").strip()
    return tenant_id, client_id, client_secret


def validate_outlook_sync_config():
    errors = []
    warnings = []

    mailbox_user = (os.environ.get("ASME_OUTLOOK_CALENDAR_USER") or "").strip()
    tenant_id, client_id, client_secret = get_outlook_oauth_config()

    if not mailbox_user:
        errors.append("ASME_OUTLOOK_CALENDAR_USER is missing.")
    if not tenant_id:
        errors.append("ASME_OUTLOOK_TENANT_ID is missing.")
    elif tenant_id.lower() in {"common", "organizations", "consumers"}:
        errors.append("ASME_OUTLOOK_TENANT_ID must be your tenant ID/domain, not common/organizations/consumers.")
    if not client_id:
        errors.append("ASME_OUTLOOK_CLIENT_ID is missing.")
    if not client_secret:
        errors.append("ASME_OUTLOOK_CLIENT_SECRET is missing.")

    if mailbox_user and "@" in mailbox_user:
        domain = mailbox_user.split("@", 1)[1].strip().lower()
        if domain in KNOWN_NON_MS_MAIL_DOMAINS:
            errors.append(
                "ASME_OUTLOOK_CALENDAR_USER must be a Microsoft mailbox (Outlook/Microsoft 365), "
                "not a consumer mailbox such as Gmail/Yahoo."
            )
        elif domain in POSSIBLY_CONSUMER_MS_DOMAINS:
            warnings.append(
                "Consumer Outlook domains may not support app-only Graph calendar access. "
                "Microsoft 365 tenant mailboxes are recommended."
            )
    elif mailbox_user:
        warnings.append("ASME_OUTLOOK_CALENDAR_USER is not in email format.")

    return errors, warnings


def has_any_outlook_sync_inputs():
    mailbox_user = (os.environ.get("ASME_OUTLOOK_CALENDAR_USER") or "").strip()
    tenant_id, client_id, client_secret = get_outlook_oauth_config()
    return bool(mailbox_user or tenant_id or client_id or client_secret)


def is_outlook_sync_configured():
    errors, _warnings = validate_outlook_sync_config()
    return len(errors) == 0


def _http_json_request(method, url, headers=None, form_data=None, json_data=None, timeout=30, retries=0, retry_backoff=0.8):
    body = None
    request_headers = dict(headers or {})

    if form_data is not None:
        body = urlencode(form_data).encode("utf-8")
        request_headers["Content-Type"] = "application/x-www-form-urlencoded"
    elif json_data is not None:
        body = json.dumps(json_data).encode("utf-8")
        request_headers["Content-Type"] = "application/json"

    req = Request(url=url, data=body, method=method)
    for key, value in request_headers.items():
        req.add_header(key, value)

    for attempt in range(retries + 1):
        try:
            with urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8") if resp else ""
                payload = json.loads(raw) if raw else {}
                return resp.status, payload, None
        except HTTPError as exc:
            error_body = exc.read().decode("utf-8", errors="ignore")
            try:
                payload = json.loads(error_body) if error_body else {}
            except Exception:
                payload = {"raw": error_body}

            if exc.code in (429, 500, 502, 503, 504) and attempt < retries:
                time_module.sleep(retry_backoff * (2**attempt))
                continue
            return exc.code, payload, error_body
        except URLError as exc:
            if attempt < retries:
                time_module.sleep(retry_backoff * (2**attempt))
                continue
            return None, {}, str(exc)
        except Exception as exc:
            if attempt < retries:
                time_module.sleep(retry_backoff * (2**attempt))
                continue
            return None, {}, str(exc)

    return None, {}, "HTTP request failed"


def get_outlook_access_token():
    now = time_module.time()
    tenant_id, client_id, client_secret = get_outlook_oauth_config()
    if not tenant_id or not client_id or not client_secret:
        return None, (
            "Outlook OAuth is not configured. Set ASME_OUTLOOK_TENANT_ID, "
            "ASME_OUTLOOK_CLIENT_ID, and ASME_OUTLOOK_CLIENT_SECRET."
        )

    cached_token = OUTLOOK_TOKEN_CACHE.get("access_token")
    cached_expiry = OUTLOOK_TOKEN_CACHE.get("expires_at") or 0
    cached_tenant = OUTLOOK_TOKEN_CACHE.get("tenant_id") or ""
    cached_client = OUTLOOK_TOKEN_CACHE.get("client_id") or ""
    if (
        cached_token
        and cached_expiry > (now + 60)
        and cached_tenant == tenant_id
        and cached_client == client_id
    ):
        return cached_token, None

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    status, payload, error = _http_json_request(
        method="POST",
        url=token_url,
        form_data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        retries=2,
    )
    if status != 200:
        error_text = payload.get("error_description") or payload.get("error") or error or "token request failed"
        return None, f"Outlook token request failed: {str(error_text)[:250]}"

    access_token = payload.get("access_token")
    if not access_token:
        return None, "Outlook token response missing access_token."

    expires_in = int(payload.get("expires_in") or 0)
    OUTLOOK_TOKEN_CACHE["access_token"] = access_token
    OUTLOOK_TOKEN_CACHE["tenant_id"] = tenant_id
    OUTLOOK_TOKEN_CACHE["client_id"] = client_id
    OUTLOOK_TOKEN_CACHE["expires_at"] = now + max(60, expires_in - 120) if expires_in else (now + 900)
    return access_token, None


def create_outlook_calendar_event(meeting):
    validation_errors, _warnings = validate_outlook_sync_config()
    if validation_errors:
        return None, None, " ".join(validation_errors)

    mailbox_user = (os.environ.get("ASME_OUTLOOK_CALENDAR_USER") or "").strip()

    access_token, token_error = get_outlook_access_token()
    if token_error:
        return None, None, token_error

    timezone = get_calendar_timezone()
    start_dt = datetime.combine(meeting.meeting_date, meeting.start_time)
    end_dt = datetime.combine(meeting.meeting_date, meeting.end_time)
    calendar_id = get_outlook_calendar_id_for_room(meeting.room)

    description_lines = []
    if meeting.requester_email:
        description_lines.append(f"Requested by: {meeting.requester_email}")
    if meeting.notes:
        description_lines.append(f"Notes: {meeting.notes}")

    event_body = {
        "subject": f"{meeting.team_name} - {meeting.room}",
        "body": {
            "contentType": "Text",
            "content": "\n".join(description_lines).strip() or "Created from ASME website scheduler.",
        },
        "start": {"dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": timezone},
        "end": {"dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": timezone},
        "location": {"displayName": meeting.room},
    }

    encoded_user = quote(mailbox_user, safe="")
    if calendar_id:
        encoded_calendar = quote(calendar_id, safe="")
        endpoint = f"https://graph.microsoft.com/v1.0/users/{encoded_user}/calendars/{encoded_calendar}/events"
    else:
        endpoint = f"https://graph.microsoft.com/v1.0/users/{encoded_user}/events"

    status, payload, error = _http_json_request(
        method="POST",
        url=endpoint,
        headers={"Authorization": f"Bearer {access_token}"},
        json_data=event_body,
        retries=2,
    )
    if status not in (200, 201):
        msg = payload.get("error", {}).get("message") if isinstance(payload.get("error"), dict) else None
        error_text = msg or error or "event create failed"
        return None, None, f"Outlook event create failed: {str(error_text)[:250]}"

    event_id = payload.get("id")
    if not event_id:
        return None, None, "Outlook event created but no event ID returned."
    return calendar_id, event_id, None


def delete_outlook_calendar_event(meeting):
    if not meeting.outlook_event_id:
        return None

    validation_errors, _warnings = validate_outlook_sync_config()
    if validation_errors:
        return " ".join(validation_errors)

    mailbox_user = (os.environ.get("ASME_OUTLOOK_CALENDAR_USER") or "").strip()

    access_token, token_error = get_outlook_access_token()
    if token_error:
        return token_error

    encoded_user = quote(mailbox_user, safe="")
    encoded_event = quote(meeting.outlook_event_id, safe="")
    endpoint = f"https://graph.microsoft.com/v1.0/users/{encoded_user}/events/{encoded_event}"

    status, payload, error = _http_json_request(
        method="DELETE",
        url=endpoint,
        headers={"Authorization": f"Bearer {access_token}"},
        retries=2,
    )
    if status in (200, 202, 204, 404):
        return None

    msg = payload.get("error", {}).get("message") if isinstance(payload.get("error"), dict) else None
    error_text = msg or error or "event delete failed"
    return f"Outlook event delete failed: {str(error_text)[:250]}"


def send_meeting_cancel_confirmation_email(meeting, confirm_url, reject_url):
    smtp_host = (os.environ.get("ASME_SMTP_HOST") or "smtp.office365.com").strip()
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
    default_calendar_id = (os.environ.get("ASME_OUTLOOK_CALENDAR_ID") or "").strip()
    robotics_calendar_id = (os.environ.get("ASME_OUTLOOK_CALENDAR_ROBOTICS_ID") or "").strip()
    fluids_calendar_id = (os.environ.get("ASME_OUTLOOK_CALENDAR_FLUIDS_ID") or "").strip()
    mailbox_user = (os.environ.get("ASME_OUTLOOK_CALENDAR_USER") or "").strip()
    tenant_id, client_id, client_secret = get_outlook_oauth_config()
    smtp_user = (os.environ.get("ASME_SMTP_USER") or "").strip()
    smtp_pass = (os.environ.get("ASME_SMTP_PASS") or "").strip()
    cancel_notify_to = (os.environ.get("ASME_CANCEL_NOTIFY_TO") or "").strip()

    has_any_calendar_id = bool(default_calendar_id or robotics_calendar_id or fluids_calendar_id)
    return {
        "tenant_id_set": bool(tenant_id),
        "client_id_set": bool(client_id),
        "client_secret_set": bool(client_secret),
        "calendar_user": mailbox_user,
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
    member_tag = clean_tag_value(member_tag)
    member_id = str(member_id or "").strip()

    if member_tag:
        member = Member.query.filter_by(nfc_tag=member_tag).first()
        if member:
            return member
        user = resolve_user_from_tag_uid(member_tag)
        if user:
            linked = current_user_member(user)
            if linked:
                return linked

    if member_id:
        try:
            return db.session.get(Member, int(member_id))
        except Exception:
            return None
    return None


def resolve_item(item_tag, item_id):
    item_tag = clean_tag_value(item_tag)
    item_id = str(item_id or "").strip()

    if item_tag:
        item, _resolved_via = find_item_by_tag(item_tag)
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


def serialize_user(user):
    if not user:
        return None
    return {
        "id": user.id,
        "name": user.name,
        "email": user.email,
        "username": user.username or "",
        "role": normalize_role(user.role),
        "is_active": bool(user.is_active),
        "member_id": user.member_id,
        "created_at": iso_or_none(user.created_at),
        "last_login_at": iso_or_none(user.last_login_at),
    }


def serialize_item(item):
    return {
        "id": item.id,
        "name": item.name,
        "description": item.description,
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
        "status": tx.status,
        "qty": tx.qty,
        "checkout_time": iso_or_none(tx.checkout_time),
        "return_time": iso_or_none(tx.return_time),
        "due_date": iso_or_none(tx.due_date),
        "notes": tx.notes,
        "checkout_notes": tx.checkout_notes,
        "return_condition": tx.return_condition,
        "return_notes": tx.return_notes,
        "return_photo_path": tx.return_photo_path,
    }


def serialize_open_checkout(tx):
    checkout_at = tx.checkout_time or tx.timestamp
    return {
        "transaction_id": tx.id,
        "item_id": tx.item_id,
        "item_name": tx.item.name if tx.item else "",
        "item_category": tx.item.category if tx.item else None,
        "item_location": tx.item.location if tx.item else None,
        "qty": tx.qty,
        "checkout_time": iso_or_none(checkout_at),
        "checkout_notes": tx.checkout_notes or tx.notes,
        "status": tx.status,
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
    ensure_inventory_schema_columns()
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
    ensure_inventory_schema_columns()
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
    ensure_inventory_schema_columns()
    h2s_print_cmd = (os.environ.get("ASME_H2S_PRINT_CMD") or "").strip()
    p1s_print_cmd = (os.environ.get("ASME_P1S_PRINT_CMD") or "").strip()
    active_member = get_active_member()

    context = dashboard_context(transaction_limit=transaction_limit)
    context.update(
        {
            "active_page": active_page,
            "page_title": page_title,
            "page_subtitle": page_subtitle,
            "active_member": active_member,
            "active_member_is_admin": is_admin_member(active_member),
            "h2s_print_cmd_configured": bool(h2s_print_cmd),
            "p1s_print_cmd_configured": bool(p1s_print_cmd),
            "h2s_print_cmd_value": h2s_print_cmd,
            "p1s_print_cmd_value": p1s_print_cmd,
            "print_commands_env_file": str(PRINT_COMMANDS_ENV_FILE),
        }
    )
    return render_template(template_name, **context)


def frontend_portal_context():
    ensure_inventory_schema_columns()
    ensure_portal_schema()
    active_member = get_active_member()
    members = Member.query.order_by(Member.name.asc()).all()
    item_count = Item.query.count()
    available_total = db.session.query(func.coalesce(func.sum(Item.available_qty), 0)).scalar() or 0
    my_open_count = 0
    if active_member:
        my_open_count = (
            Transaction.query.filter_by(member_id=active_member.id, status="OUT")
            .count()
        )
    return {
        "active_member": active_member,
        "active_member_is_admin": is_admin_member(active_member),
        "members": members,
        "item_count": item_count,
        "available_total": int(available_total),
        "my_open_count": my_open_count,
        "club_mission": FRONT_CLUB_MISSION,
        "club_highlights": FRONT_CLUB_HIGHLIGHTS,
        "project_showcase": FRONT_PROJECT_SHOWCASE,
    }


@app.before_request
def ensure_runtime_schema():
    # Keeps local SQLite upgrades seamless without introducing migration tooling for this project.
    if request.path.startswith("/static/"):
        return
    ensure_inventory_schema_columns()
    ensure_meeting_schema_columns()
    ensure_portal_schema()


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
        "scan": "scan_page",
        "my_items": "my_items_page",
        "admin_nfc": "admin_nfc_page",
    }
    endpoint = page_endpoints.get(page, "dashboard_page")
    return redirect(url_for(endpoint))


def public_site_context(page_title):
    user = current_auth_user()
    projects = Project.query.order_by(Project.created_at.desc(), Project.id.desc()).all()
    announcements = (
        Announcement.query.filter_by(is_published=True, show_on_public=True)
        .order_by(func.coalesce(Announcement.published_at, Announcement.created_at).desc(), Announcement.id.desc())
        .limit(5)
        .all()
    )
    project_filters = sorted({(project.project_type or "General").strip() for project in projects if project})
    if not project_filters:
        project_filters = ["General"]
    executives = (
        User.query.filter(User.role.in_(["team_leader", "admin"]), User.is_active.is_(True))
        .order_by(User.role.desc(), User.name.asc())
        .all()
    )
    executive_cards = []
    for exec_user in executives:
        title = "Executive Member"
        if normalize_role(exec_user.role) == "admin":
            title = "Administrator"
        elif normalize_role(exec_user.role) == "team_leader":
            title = "Team Leader"
        exec_title = (exec_user.exec_title or "").strip() or title
        exec_message = (
            (exec_user.exec_message or "").strip()
            or "Focused on safe builds, strong documentation, and reliable execution."
        )
        headshot_url = (exec_user.headshot_url or "").strip() or url_for("static", filename="asme_logo.png")
        executive_cards.append(
            {
                "name": exec_user.name,
                "title": exec_title,
                "message": exec_message,
                "headshot": headshot_url,
            }
        )
    if not executive_cards:
        executive_cards = [
            {
                "name": "ASME Executive Team",
                "title": "Leadership",
                "message": "Add executive member accounts to populate this section.",
                "headshot": url_for("static", filename="asme_logo.png"),
            }
        ]

    return {
        "page_title": page_title,
        "current_user": user,
        "projects": projects,
        "project_filters": project_filters,
        "announcements": announcements,
        "executive_cards": executive_cards,
    }


def member_dashboard_context():
    user = current_auth_user()
    member = current_user_member(user)
    open_checkouts = []
    if member:
        open_checkouts = (
            Transaction.query.filter_by(member_id=member.id, status="OUT")
            .order_by(Transaction.checkout_time.desc(), Transaction.id.desc())
            .all()
        )
    items = Item.query.filter(Item.active.is_(True)).order_by(Item.name.asc(), Item.id.asc()).all()
    print_requests = (
        PrintRequest.query.filter_by(user_id=user.id)
        .order_by(PrintRequest.created_at.desc(), PrintRequest.id.desc())
        .all()
    )
    upcoming_events = (
        Event.query.filter(
            Event.status != "cancelled",
            Event.end_time >= datetime.now() - timedelta(hours=1),
        )
        .order_by(Event.start_time.asc(), Event.id.asc())
        .limit(25)
        .all()
    )
    announcements = (
        Announcement.query.filter_by(is_published=True, show_on_member=True)
        .order_by(func.coalesce(Announcement.published_at, Announcement.created_at).desc(), Announcement.id.desc())
        .limit(8)
        .all()
    )
    my_help_messages = (
        ContactMessage.query.filter_by(user_id=user.id, kind="help")
        .order_by(ContactMessage.created_at.desc(), ContactMessage.id.desc())
        .limit(40)
        .all()
    )
    return {
        "current_user": user,
        "member_profile": member,
        "items": items,
        "open_checkouts": open_checkouts,
        "print_requests": print_requests,
        "portal_print_printers": PRINT_REQUEST_PRINTERS,
        "upcoming_events": upcoming_events,
        "calendar_embed_url": portal_calendar_embed_url(),
        "announcements": announcements,
        "my_help_messages": my_help_messages,
    }


def admin_dashboard_context():
    user = current_auth_user()
    members = Member.query.order_by(Member.name.asc()).all()
    users = User.query.order_by(User.created_at.desc(), User.id.desc()).all()
    tags = NFCTag.query.order_by(NFCTag.assigned_at.desc(), NFCTag.id.desc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    item_primary_tags = get_item_primary_tag_map()
    transactions = (
        Transaction.query.order_by(Transaction.timestamp.desc(), Transaction.id.desc())
        .limit(120)
        .all()
    )
    print_requests = (
        PrintRequest.query.order_by(PrintRequest.created_at.desc(), PrintRequest.id.desc()).limit(120).all()
    )
    events = Event.query.order_by(Event.start_time.desc(), Event.id.desc()).limit(120).all()
    attendance_records = (
        AttendanceRecord.query.order_by(AttendanceRecord.checkin_time.desc(), AttendanceRecord.id.desc())
        .limit(200)
        .all()
    )
    projects = Project.query.order_by(Project.created_at.desc(), Project.id.desc()).all()
    contact_messages = (
        ContactMessage.query.order_by(ContactMessage.created_at.desc(), ContactMessage.id.desc()).limit(60).all()
    )
    audit_logs = AuditLog.query.order_by(AuditLog.created_at.desc(), AuditLog.id.desc()).limit(150).all()
    now = datetime.now()
    active_members_count = User.query.filter(User.is_active.is_(True), User.role.in_(["member", "team_leader", "admin"])).count()
    checked_out_now_count = (
        db.session.query(func.coalesce(func.sum(Transaction.qty), 0))
        .filter(Transaction.status == "OUT")
        .scalar()
        or 0
    )
    today_start = datetime.combine(date.today(), time.min)
    today_end = today_start + timedelta(days=1)
    attendance_today_count = (
        AttendanceRecord.query.filter(
            AttendanceRecord.checkin_time >= today_start,
            AttendanceRecord.checkin_time < today_end,
        ).count()
    )
    overdue_items_count = Transaction.query.filter(
        Transaction.status == "OUT",
        Transaction.due_date.isnot(None),
        Transaction.due_date < now.date(),
    ).count()
    low_stock_items_count = (
        Item.query.filter(Item.active.is_(True), Item.available_qty <= Item.min_stock_threshold)
        .order_by(Item.available_qty.asc(), Item.id.asc())
        .count()
    )
    upcoming_meetings_count = Event.query.filter(Event.start_time >= now).count()
    pending_print_count = PrintRequest.query.filter(PrintRequest.status.in_(["submitted", "approved", "printing"])).count()
    calendar_provider = calendar_provider_name()
    google_embed_set = bool((os.environ.get("ASME_GOOGLE_CALENDAR_EMBED_URL") or "").strip())
    outlook_embed_set = bool((os.environ.get("ASME_OUTLOOK_CALENDAR_EMBED_URL") or "").strip())
    return {
        "current_user": user,
        "members": members,
        "users": users,
        "tags": tags,
        "items": items,
        "item_primary_tags": item_primary_tags,
        "transactions": transactions,
        "print_requests": print_requests,
        "events": events,
        "attendance_records": attendance_records,
        "projects": projects,
        "contact_messages": contact_messages,
        "audit_logs": audit_logs,
        "roles": ["member", "team_leader", "admin"],
        "print_request_statuses": PRINT_REQUEST_STATUSES,
        "portal_print_printers": PRINT_REQUEST_PRINTERS,
        "item_types": ITEM_TYPES,
        "active_members_count": active_members_count,
        "checked_out_now_count": int(checked_out_now_count),
        "attendance_today_count": attendance_today_count,
        "overdue_items_count": overdue_items_count,
        "low_stock_items_count": low_stock_items_count,
        "upcoming_meetings_count": upcoming_meetings_count,
        "pending_print_count": pending_print_count,
        "calendar_provider": calendar_provider,
        "google_embed_set": google_embed_set,
        "outlook_embed_set": outlook_embed_set,
    }


@app.get("/")
def public_home():
    return render_template("site/home.html", **public_site_context("Home"))


@app.get("/who-we-are")
def public_who_we_are():
    context = public_site_context("Who We Are")
    context["mission"] = FRONT_CLUB_MISSION
    context["highlights"] = FRONT_CLUB_HIGHLIGHTS
    context["asme_facts"] = FRONT_ABOUT_ASME_FACTS
    context["uiowa_projects"] = FRONT_UIOWA_CURRENT_PROJECTS
    return render_template("site/who_we_are.html", **context)


@app.get("/about")
def public_about_alias():
    return redirect(url_for("public_who_we_are"))


@app.get("/executive-team")
def public_executive_team():
    context = public_site_context("Executive Team")
    return render_template("site/executive_team.html", **context)


@app.get("/exec")
def public_exec_alias():
    return redirect(url_for("public_executive_team"))


@app.get("/projects")
def public_projects():
    return render_template("site/projects.html", **public_site_context("Projects"))


@app.get("/projects/<slug>")
def public_project_detail(slug):
    project = Project.query.filter_by(slug=(slug or "").strip().lower()).first()
    if not project:
        flash("Project not found.", "error")
        return redirect(url_for("public_projects"))
    context = public_site_context(project.title)
    context["project"] = project
    context["project_gallery"] = parse_json_list(project.gallery_json)
    context["project_timeline"] = parse_json_list(project.timeline)
    return render_template("site/project_detail.html", **context)


@app.route("/contact", methods=["GET", "POST"])
def public_contact():
    context = public_site_context("Socials + Contact")
    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        email = (request.form.get("email") or "").strip()
        subject = (request.form.get("subject") or "").strip() or None
        message = (request.form.get("message") or "").strip()
        if not name or not email or not message:
            flash("Name, email, and message are required.", "error")
            return render_template("site/contact.html", **context)
        db.session.add(
            ContactMessage(
                name=name[:160],
                email=email[:160],
                kind="contact",
                subject=(subject or "")[:220] or None,
                message=message[:5000],
                status="new",
            )
        )
        db.session.commit()
        flash("Message sent. Our admin team will follow up.", "success")
        return redirect(url_for("public_contact"))
    return render_template("site/contact.html", **context)


@app.get("/socials")
def public_socials():
    return redirect(url_for("public_contact"))


@app.route("/join", methods=["GET", "POST"])
def public_join():
    context = public_site_context("Join / Get Involved")
    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        email = (request.form.get("email") or "").strip()
        interest = (request.form.get("interest") or "").strip()
        message = (request.form.get("message") or "").strip()
        if not name or not email or not interest:
            flash("Name, email, and interest area are required.", "error")
            return render_template("site/join.html", **context)
        db.session.add(
            ContactMessage(
                name=name[:160],
                email=email[:160],
                kind="join",
                target="membership",
                subject=f"Join Interest: {interest[:120]}",
                message=(message or f"Interested in: {interest}")[:5000],
                status="new",
            )
        )
        db.session.commit()
        flash("Interest form submitted. We will contact you with onboarding details.", "success")
        return redirect(url_for("public_join"))
    return render_template("site/join.html", **context)


@app.route("/sponsors", methods=["GET", "POST"])
def public_sponsors():
    context = public_site_context("Sponsors / Partners")
    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        email = (request.form.get("email") or "").strip()
        message = (request.form.get("message") or "").strip()
        if not name or not email or not message:
            flash("Name, email, and message are required.", "error")
            return render_template("site/sponsors.html", **context)
        db.session.add(
            ContactMessage(
                name=name[:160],
                email=email[:160],
                kind="sponsor",
                target="sponsorship",
                subject="Sponsorship Inquiry",
                message=message[:5000],
                status="new",
            )
        )
        db.session.commit()
        flash("Sponsorship inquiry sent. Thank you for supporting ASME at Iowa.", "success")
        return redirect(url_for("public_sponsors"))
    return render_template("site/sponsors.html", **context)


@app.route("/signup", methods=["GET", "POST"])
def signup_page():
    context = public_site_context("Sign Up")
    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        email = (request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""
        role = normalize_role(request.form.get("role") or "member")
        if role == "admin":
            role = "member"

        if len(name) < 2:
            flash("Please enter your full name.", "error")
            return render_template("site/signup.html", **context)
        if "@" not in email or len(email) < 5:
            flash("Please enter a valid email.", "error")
            return render_template("site/signup.html", **context)
        if len(password) < 8:
            flash("Password must be at least 8 characters.", "error")
            return render_template("site/signup.html", **context)
        if User.query.filter(func.lower(User.email) == email).first():
            flash("An account already exists for that email.", "error")
            return render_template("site/signup.html", **context)

        linked_member = Member.query.filter(func.lower(Member.email) == email).first()
        username = make_unique_username(email.split("@", 1)[0] if "@" in email else name)
        user = User(
            name=name[:160],
            email=email,
            username=username,
            password_hash=generate_password_hash(password),
            role=role,
            is_active=True,
            member_id=linked_member.id if linked_member else None,
        )
        db.session.add(user)
        db.session.commit()
        flash("Account created. You can now log in.", "success")
        return redirect(url_for("login_page"))
    return render_template("site/signup.html", **context)


@app.get("/kiosk")
def kiosk_entry():
    user = current_auth_user()
    if user:
        return redirect(url_for("portal_member_inventory"))
    context = public_site_context("Kiosk Login")
    context["inventory_next"] = url_for("portal_member_inventory")
    context["admin_next"] = url_for("portal_admin_dashboard")
    return render_template("site/kiosk.html", **context)


@app.get("/checkin")
@require_login
def checkin_entry():
    user = current_auth_user()
    candidates = current_checkin_candidate_events()
    if not candidates:
        context = public_site_context("Meeting Check-in")
        context["active_event"] = None
        return render_template("site/checkin.html", **context)
    if len(candidates) == 1:
        event = candidates[0]
        _row, created = check_user_into_event(user, event, method="shared_tag")
        status = "new" if created else "existing"
        return redirect(url_for("checkin_success", event_id=event.id, status=status))
    session["checkin_candidate_ids"] = [row.id for row in candidates]
    return redirect(url_for("checkin_select"))


@app.route("/checkin/select", methods=["GET", "POST"])
@require_login
def checkin_select():
    candidate_ids = session.get("checkin_candidate_ids") or []
    if not candidate_ids:
        flash("No active meeting choices. Try scanning again.", "info")
        return redirect(url_for("checkin_entry"))
    candidates = (
        Event.query.filter(Event.id.in_(candidate_ids))
        .order_by(Event.start_time.asc(), Event.id.asc())
        .all()
    )
    if not candidates:
        session.pop("checkin_candidate_ids", None)
        flash("No active meeting choices. Try scanning again.", "info")
        return redirect(url_for("checkin_entry"))
    if request.method == "POST":
        event_id = parse_positive_int(request.form.get("event_id"), default=0)
        event = next((row for row in candidates if row.id == event_id), None)
        if not event:
            flash("Select a valid meeting.", "error")
            return redirect(url_for("checkin_select"))
        _row, created = check_user_into_event(current_auth_user(), event, method="shared_tag")
        session.pop("checkin_candidate_ids", None)
        status = "new" if created else "existing"
        return redirect(url_for("checkin_success", event_id=event.id, status=status))
    context = public_site_context("Select Meeting")
    context["candidate_events"] = candidates
    return render_template("site/checkin_select.html", **context)


@app.get("/checkin/success")
@require_login
def checkin_success():
    event_id = parse_positive_int(request.args.get("event_id"), default=0)
    status = (request.args.get("status") or "new").strip().lower()
    event = db.session.get(Event, event_id)
    if not event:
        flash("Meeting not found.", "error")
        return redirect(url_for("checkin_entry"))
    context = public_site_context("Check-in Success")
    context["event"] = event
    context["status"] = status
    return render_template("site/checkin_success.html", **context)


@app.route("/login", methods=["GET", "POST"])
def login_page():
    context = public_site_context("Login")
    next_url = (request.args.get("next") or request.form.get("next") or "").strip()
    if request.method == "POST":
        identifier = (
            request.form.get("identifier")
            or request.form.get("email")
            or ""
        ).strip().lower()
        password = request.form.get("password") or ""
        blocked, retry_after = is_login_rate_limited(identifier)
        if blocked:
            flash(f"Too many login attempts. Try again in about {retry_after} seconds.", "error")
            context["next_url"] = next_url
            return render_template("site/login.html", **context)
        user = find_user_by_login_identifier(identifier)
        if not user or not user.is_active or not check_password_hash(user.password_hash, password):
            record_login_failure(identifier)
            flash("Invalid email or password.", "error")
            context["next_url"] = next_url
            return render_template("site/login.html", **context)
        clear_login_failures(identifier)
        sign_in_user(user)
        if next_url.startswith("/"):
            return redirect(next_url)
        return redirect(url_for("portal_router"))
    context["next_url"] = next_url
    return render_template("site/login.html", **context)


@app.route("/admin-login", methods=["GET", "POST"])
def admin_login_page():
    context = public_site_context("Admin Login")
    next_url = (request.args.get("next") or request.form.get("next") or "").strip()
    shared_admin_email, shared_admin_password = get_shared_admin_credentials()
    context["shared_admin_email"] = shared_admin_email
    current = current_auth_user()
    if current and role_allows(current.role, "admin"):
        return redirect(url_for("portal_admin_dashboard"))

    if request.method == "POST":
        identifier = (
            request.form.get("identifier")
            or request.form.get("email")
            or ""
        ).strip().lower()
        if not identifier and shared_admin_email:
            identifier = shared_admin_email
        password = request.form.get("password") or ""
        blocked, retry_after = is_login_rate_limited(identifier)
        if blocked:
            flash(f"Too many login attempts. Try again in about {retry_after} seconds.", "error")
            context["next_url"] = next_url
            return render_template("site/admin_login.html", **context)

        # Allow one shared admin login from environment variables.
        if (
            shared_admin_email
            and shared_admin_password
            and identifier == shared_admin_email
            and password == shared_admin_password
        ):
            user = ensure_shared_admin_user(shared_admin_email, shared_admin_password)
            if user:
                clear_login_failures(identifier)
                sign_in_user(user)
                if next_url.startswith("/"):
                    return redirect(next_url)
                return redirect(url_for("portal_admin_dashboard"))

        user = find_user_by_login_identifier(identifier)
        if not user or not user.is_active or not check_password_hash(user.password_hash, password):
            record_login_failure(identifier)
            flash("Invalid admin credentials.", "error")
            context["next_url"] = next_url
            return render_template("site/admin_login.html", **context)
        if not role_allows(user.role, "admin"):
            record_login_failure(identifier)
            flash("This account is not an admin account.", "error")
            context["next_url"] = next_url
            return render_template("site/admin_login.html", **context)
        clear_login_failures(identifier)
        sign_in_user(user)
        if next_url.startswith("/"):
            return redirect(next_url)
        return redirect(url_for("portal_admin_dashboard"))

    context["next_url"] = next_url
    return render_template("site/admin_login.html", **context)


@app.get("/logout")
@app.post("/logout")
def logout_page():
    sign_out_user()
    flash("You have been logged out.", "info")
    return redirect(url_for("public_home"))


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password_page():
    context = public_site_context("Forgot Password")
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        user = User.query.filter(func.lower(User.email) == email).first()
        if user and user.is_active:
            token = secrets.token_urlsafe(32)
            expires_at = datetime.utcnow() + timedelta(hours=PASSWORD_RESET_HOURS)
            db.session.add(PasswordResetToken(user_id=user.id, token=token, expires_at=expires_at))
            db.session.commit()
            reset_link = url_for("reset_password_page", token=token, _external=True)
            flash(f"Reset link generated: {reset_link}", "info")
        else:
            flash("If this email exists, a reset link has been generated.", "info")
        return redirect(url_for("forgot_password_page"))
    return render_template("site/forgot_password.html", **context)


@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password_page(token):
    context = public_site_context("Reset Password")
    reset_row = PasswordResetToken.query.filter_by(token=(token or "").strip()).first()
    if (
        not reset_row
        or reset_row.used_at is not None
        or reset_row.expires_at is None
        or reset_row.expires_at < datetime.utcnow()
    ):
        flash("Reset link is invalid or expired.", "error")
        return redirect(url_for("forgot_password_page"))

    if request.method == "POST":
        password = request.form.get("password") or ""
        confirm = request.form.get("confirm_password") or ""
        if len(password) < 8:
            flash("Password must be at least 8 characters.", "error")
            return render_template("site/reset_password.html", **context)
        if password != confirm:
            flash("Passwords do not match.", "error")
            return render_template("site/reset_password.html", **context)
        user = db.session.get(User, reset_row.user_id)
        if not user:
            flash("User no longer exists.", "error")
            return redirect(url_for("forgot_password_page"))
        user.password_hash = generate_password_hash(password)
        reset_row.used_at = datetime.utcnow()
        db.session.commit()
        flash("Password updated. You can now log in.", "success")
        return redirect(url_for("login_page"))

    return render_template("site/reset_password.html", **context)


@app.get("/portal")
@require_login
def portal_router():
    user = current_auth_user()
    if role_allows(user.role, "admin"):
        return redirect(url_for("portal_admin_dashboard"))
    return redirect(url_for("portal_member_dashboard"))


@app.get("/portal/member")
@require_role("member")
def portal_member_dashboard():
    context = member_dashboard_context()
    context["page_title"] = "Member Dashboard"
    context["active_page"] = "member_dashboard"
    context["tile_counts"] = {
        "inventory": len(context.get("items", [])),
        "my_items": len(context.get("open_checkouts", [])),
        "prints": len(context.get("print_requests", [])),
        "upcoming_events": len(context.get("upcoming_events", [])),
        "announcements": len(context.get("announcements", [])),
    }
    return render_template("portal/member_dashboard_home.html", **context)


@app.get("/portal/member/inventory")
@require_role("member")
def portal_member_inventory():
    context = member_dashboard_context()
    context["page_title"] = "Inventory"
    context["active_page"] = "member_inventory"
    query = (request.args.get("q") or "").strip().lower()
    if query:
        filtered_items = []
        for item in context["items"]:
            haystack = " ".join(
                [
                    item.name or "",
                    item.category or "",
                    item.location or "",
                    item.description or "",
                    item.notes or "",
                ]
            ).lower()
            if query in haystack:
                filtered_items.append(item)
        context["items"] = filtered_items
    context["query"] = query
    return render_template("portal/member_inventory.html", **context)


@app.get("/portal/member/items/<int:item_id>")
@app.get("/portal/member/inventory/<int:item_id>")
@require_role("member")
def portal_member_item_detail(item_id):
    context = member_dashboard_context()
    context["page_title"] = "Item Detail"
    context["active_page"] = "member_inventory"
    item = db.session.get(Item, item_id)
    if not item:
        flash("Item not found.", "error")
        return redirect(url_for("portal_member_inventory"))
    member = context.get("member_profile")
    open_tx = get_open_checkout(member.id, item.id) if member else None
    context["item"] = item
    context["open_tx"] = open_tx
    return render_template("portal/member_item_detail.html", **context)


@app.get("/portal/member/checkouts")
@app.get("/portal/member/my-items")
@require_role("member")
def portal_member_checkouts():
    context = member_dashboard_context()
    context["page_title"] = "My Items"
    context["active_page"] = "member_my_items"
    return render_template("portal/member_checkouts.html", **context)


@app.get("/portal/member/prints")
@require_role("member")
def portal_member_prints():
    context = member_dashboard_context()
    context["page_title"] = "3D Printing"
    context["active_page"] = "member_prints"
    return render_template("portal/member_prints.html", **context)


@app.get("/portal/member/prints/<int:request_id>")
@require_role("member")
def portal_member_print_detail(request_id):
    context = member_dashboard_context()
    context["page_title"] = "Print Request Detail"
    context["active_page"] = "member_prints"
    row = db.session.get(PrintRequest, request_id)
    user = context["current_user"]
    if not row or row.user_id != user.id:
        flash("Print request not found.", "error")
        return redirect(url_for("portal_member_prints"))
    context["print_request"] = row
    return render_template("portal/member_print_detail.html", **context)


@app.get("/portal/member/calendar")
@require_role("member")
def portal_member_calendar():
    context = member_dashboard_context()
    context["page_title"] = "Calendar"
    context["active_page"] = "member_calendar"
    context["team_mode"] = role_allows(context["current_user"].role, "team_leader")
    context["meeting_rooms"] = MEETING_ROOMS
    context["calendar_provider"] = calendar_provider_name()
    return render_template("portal/member_calendar.html", **context)


@app.route("/portal/member/schedule", methods=["GET", "POST"])
@require_role("team_leader")
def portal_member_schedule():
    context = member_dashboard_context()
    context["page_title"] = "Schedule Meeting"
    context["active_page"] = "member_schedule"
    context["meeting_rooms"] = MEETING_ROOMS
    context["calendar_provider"] = calendar_provider_name()
    config = google_calendar_config()
    context["google_schedule_enabled"] = config["enabled"]
    context["google_schedule_errors"] = config["errors"]
    context["duration_options"] = [30, 60, 90]
    context["range_options"] = [7, 14]
    form_values = {
        "team_name": "",
        "duration": 60,
        "days": min(max(CALENDAR_SCHEDULING_DAYS_DEFAULT, 7), 14),
        "location": "",
        "notes": "",
    }
    slots = []

    if request.method == "POST":
        form_values["team_name"] = (request.form.get("team_name") or "").strip()
        form_values["duration"] = parse_positive_int(request.form.get("duration"), default=60)
        if form_values["duration"] not in {30, 60, 90}:
            form_values["duration"] = 60
        form_values["days"] = parse_positive_int(request.form.get("days"), default=14)
        if form_values["days"] not in {7, 14}:
            form_values["days"] = 14
        form_values["location"] = normalize_meeting_room(request.form.get("location"))
        form_values["notes"] = (request.form.get("notes") or "").strip()
        action = (request.form.get("action") or "preview").strip().lower()
        if not form_values["team_name"]:
            flash("Team name is required.", "error")
            context["form_values"] = form_values
            context["slots"] = slots
            return render_template("portal/member_schedule.html", **context)

        if action == "book":
            if not config["enabled"]:
                flash("Admin needs to configure Google Calendar integration before scheduling.", "error")
                return render_template("portal/member_schedule.html", **context, form_values=form_values, slots=slots)
            team_name = form_values["team_name"] or "Team"
            slot_token = (request.form.get("slot_token") or "").strip()
            try:
                room, start_token, duration_token = slot_token.split("|", 2)
                room = normalize_meeting_room(room)
                start_dt = datetime.strptime(start_token, "%Y-%m-%dT%H:%M")
                duration_minutes = int(duration_token)
            except Exception:
                flash("Select a valid available slot.", "error")
                slots, _slot_error = compute_available_slots(
                    form_values["duration"], form_values["days"], form_values["location"]
                )
                return render_template("portal/member_schedule.html", **context, form_values=form_values, slots=slots)
            end_dt = start_dt + timedelta(minutes=duration_minutes)
            if find_event_room_conflict(room, start_dt, end_dt):
                flash("Selected slot is no longer available. Please refresh availability.", "error")
                slots, _slot_error = compute_available_slots(
                    form_values["duration"], form_values["days"], form_values["location"]
                )
                return render_template("portal/member_schedule.html", **context, form_values=form_values, slots=slots)
            event_id, event_link, create_error = create_google_meeting_event(
                room=room,
                team_name=team_name,
                organizer_name=context["current_user"].name,
                start_local=start_dt,
                end_local=end_dt,
                notes=form_values["notes"],
            )
            if create_error:
                flash(create_error, "error")
                slots, _slot_error = compute_available_slots(
                    form_values["duration"], form_values["days"], form_values["location"]
                )
                return render_template("portal/member_schedule.html", **context, form_values=form_values, slots=slots)
            event = Event(
                title=f"ASME - {team_name} Meeting"[:220],
                description=form_values["notes"][:4000] if form_values["notes"] else None,
                location=room,
                status="scheduled",
                start_time=start_dt,
                end_time=end_dt,
                requested_by_user_id=context["current_user"].id,
                created_by_user_id=context["current_user"].id,
                google_event_id=event_id,
                google_calendar_id=config["room_ids"].get(room),
                calendar_event_link=event_link,
            )
            db.session.add(event)
            add_audit_log("schedule_google_meeting", f"room={room} start={start_dt.isoformat()} by={context['current_user'].email}")
            db.session.commit()
            flash("Meeting scheduled successfully.", "success")
            return redirect(url_for("portal_member_calendar"))

        slots, slot_error = compute_available_slots(
            form_values["duration"],
            form_values["days"],
            form_values["location"],
        )
        if slot_error:
            flash(slot_error, "error")

    context["form_values"] = form_values
    context["slots"] = slots
    return render_template("portal/member_schedule.html", **context)


@app.get("/portal/leader/schedule")
@require_role("team_leader")
def portal_leader_schedule_alias():
    return redirect(url_for("portal_member_schedule"))


@app.route("/portal/member/help", methods=["GET", "POST"])
@require_role("member")
def portal_member_help():
    context = member_dashboard_context()
    context["page_title"] = "Help / Ask Leads"
    context["active_page"] = "member_help"
    user = context["current_user"]
    if request.method == "POST":
        target = (request.form.get("target") or "").strip() or "team_leads"
        subject = (request.form.get("subject") or "").strip() or "Help request"
        message = (request.form.get("message") or "").strip()
        if not message:
            flash("Message is required.", "error")
            return render_template("portal/member_help.html", **context)
        db.session.add(
            ContactMessage(
                name=user.name[:160],
                email=user.email[:160],
                kind="help",
                user_id=user.id,
                target=target[:80],
                subject=subject[:220],
                message=message[:5000],
                status="new",
            )
        )
        db.session.commit()
        flash("Help request sent.", "success")
        return redirect(url_for("portal_member_help"))
    return render_template("portal/member_help.html", **context)


@app.route("/portal/member/profile", methods=["GET", "POST"])
@require_role("member")
def portal_member_profile():
    context = member_dashboard_context()
    context["page_title"] = "Profile / Account Settings"
    context["active_page"] = "member_profile"
    user = context["current_user"]
    tag_row = NFCTag.query.filter_by(user_id=user.id, active=True).order_by(NFCTag.assigned_at.desc(), NFCTag.id.desc()).first()
    context["active_tag"] = tag_row
    if request.method == "POST":
        action = (request.form.get("action") or "").strip().lower()
        if action == "profile":
            name = (request.form.get("name") or "").strip()
            email = (request.form.get("email") or "").strip().lower()
            if not name or "@" not in email:
                flash("Provide a valid name and email.", "error")
                return render_template("portal/member_profile.html", **context)
            if email != user.email and User.query.filter(func.lower(User.email) == email, User.id != user.id).first():
                flash("Email already in use by another account.", "error")
                return render_template("portal/member_profile.html", **context)
            user.name = name[:160]
            user.email = email[:160]
            if user.member:
                user.member.name = user.name
                user.member.email = user.email
            db.session.commit()
            flash("Profile updated.", "success")
            return redirect(url_for("portal_member_profile"))

        if action == "password":
            current_password = request.form.get("current_password") or ""
            new_password = request.form.get("new_password") or ""
            confirm_password = request.form.get("confirm_password") or ""
            if not check_password_hash(user.password_hash, current_password):
                flash("Current password is incorrect.", "error")
                return render_template("portal/member_profile.html", **context)
            if len(new_password) < 8:
                flash("New password must be at least 8 characters.", "error")
                return render_template("portal/member_profile.html", **context)
            if new_password != confirm_password:
                flash("New password and confirm password do not match.", "error")
                return render_template("portal/member_profile.html", **context)
            user.password_hash = generate_password_hash(new_password)
            db.session.commit()
            flash("Password updated.", "success")
            return redirect(url_for("portal_member_profile"))
    return render_template("portal/member_profile.html", **context)


@app.get("/portal/team")
@require_role("team_leader")
def portal_team_dashboard():
    return redirect(url_for("portal_member_schedule"))


@app.get("/portal/admin")
@require_role("admin")
def portal_admin_dashboard():
    context = admin_dashboard_context()
    context["page_title"] = "Admin Home"
    context["active_page"] = "admin_dashboard"
    return render_template("portal/admin_dashboard_home.html", **context)


@app.get("/portal/admin/members")
@require_role("admin")
def portal_admin_members_page():
    context = admin_dashboard_context()
    context["page_title"] = "Members / Roles"
    context["active_page"] = "admin_members"
    context["latest_roster_credentials_file"] = (session.get("latest_roster_credentials_file") or "").strip()
    return render_template("portal/admin_members.html", **context)


@app.get("/portal/admin/nfc")
@require_role("admin")
def portal_admin_nfc_page():
    return redirect(url_for("portal_admin_dashboard"))


@app.get("/portal/admin/attendance")
@require_role("admin")
def portal_admin_attendance_page():
    context = admin_dashboard_context()
    context["page_title"] = "Attendance"
    context["active_page"] = "admin_attendance"
    active_tab = (request.args.get("tab") or "live").strip().lower()
    if active_tab not in {"live", "history"}:
        active_tab = "live"
    checkin_event_id = parse_positive_int(
        request.args.get("checkin_event_id") or request.args.get("event_id"),
        default=0,
    )
    checkin_event = db.session.get(Event, checkin_event_id) if checkin_event_id else None
    if not checkin_event:
        checkin_event = (
            Event.query.filter(Event.end_time >= datetime.now(), Event.status != "cancelled")
            .order_by(Event.start_time.asc(), Event.id.asc())
            .first()
        )
    if not checkin_event and context["events"]:
        checkin_event = context["events"][0]
    history_event_id = parse_positive_int(request.args.get("history_event_id"), default=0)
    history_date_raw = (request.args.get("history_date") or "").strip()
    history_date = None
    if history_date_raw:
        try:
            history_date = datetime.strptime(history_date_raw, "%Y-%m-%d").date()
        except Exception:
            history_date = None
    history_query = AttendanceRecord.query
    if history_event_id:
        history_query = history_query.filter(AttendanceRecord.event_id == history_event_id)
    if history_date:
        history_start = datetime.combine(history_date, time.min)
        history_end = history_start + timedelta(days=1)
        history_query = history_query.filter(
            AttendanceRecord.checkin_time >= history_start,
            AttendanceRecord.checkin_time < history_end,
        )
    context["active_tab"] = active_tab
    context["checkin_event"] = checkin_event
    context["history_records"] = (
        history_query.order_by(AttendanceRecord.checkin_time.desc(), AttendanceRecord.id.desc())
        .limit(350)
        .all()
    )
    context["history_event_id"] = history_event_id
    context["history_date"] = history_date_raw
    return render_template("portal/admin_attendance.html", **context)


@app.get("/portal/admin/inventory")
@require_role("admin")
def portal_admin_inventory_page():
    context = admin_dashboard_context()
    context["page_title"] = "Inventory Admin"
    context["active_page"] = "admin_inventory"
    context["low_stock_items"] = (
        Item.query.filter(Item.active.is_(True), Item.available_qty <= Item.min_stock_threshold)
        .order_by(Item.available_qty.asc(), Item.name.asc(), Item.id.asc())
        .all()
    )
    context["overdue_transactions"] = (
        Transaction.query.filter(
            Transaction.status == "OUT",
            Transaction.due_date.isnot(None),
            Transaction.due_date < date.today(),
        )
        .order_by(Transaction.due_date.asc(), Transaction.id.asc())
        .all()
    )
    return render_template("portal/admin_inventory.html", **context)


@app.get("/portal/admin/prints")
@require_role("admin")
def portal_admin_prints_page():
    context = admin_dashboard_context()
    context["page_title"] = "Prints Admin"
    context["active_page"] = "admin_prints"
    grouped = {printer: [] for printer in PRINT_REQUEST_PRINTERS}
    for row in context.get("print_requests", []):
        grouped.setdefault(row.printer_type, []).append(row)
    context["print_requests_by_printer"] = grouped
    return render_template("portal/admin_prints.html", **context)


@app.route("/portal/admin/announcements", methods=["GET", "POST"])
@require_role("admin")
def portal_admin_announcements_page():
    return redirect(url_for("portal_admin_dashboard"))


@app.route("/portal/admin/content", methods=["GET", "POST"])
@require_role("admin")
def portal_admin_content_page():
    return redirect(url_for("portal_admin_dashboard"))


@app.get("/portal/admin/calendar")
@require_role("admin")
def portal_admin_calendar_page():
    return redirect(url_for("portal_admin_dashboard"))


@app.get("/portal/admin/exports")
@require_role("admin")
def portal_admin_exports_page():
    return redirect(url_for("portal_admin_dashboard"))


@app.get("/portal/admin/exports/download.zip")
@require_role("admin")
def portal_admin_exports_zip():
    return redirect(url_for("portal_admin_dashboard"))


@app.get("/portal/admin/settings")
@require_role("admin")
def portal_admin_settings_page():
    context = admin_dashboard_context()
    context["page_title"] = "System Status"
    context["active_page"] = "admin_settings"
    google_config = google_calendar_config()
    calendar_provider = calendar_provider_name()
    embed_configured = bool((portal_calendar_embed_url() or "").strip())
    context["system_status"] = {
        "google_calendar_configured": calendar_provider == "google" and (google_config["enabled"] or embed_configured),
        "printer_automation_configured": bool((os.environ.get("ASME_H2S_PRINT_CMD") or "").strip())
        and bool((os.environ.get("ASME_P1S_PRINT_CMD") or "").strip()),
        "legacy_ops_enabled": legacy_ops_enabled(),
    }
    context["calendar_provider"] = calendar_provider
    context["calendar_errors"] = google_config["errors"] if calendar_provider == "google" else []
    return render_template("portal/admin_settings.html", **context)


@app.post("/portal/admin/announcements/<int:announcement_id>/update")
@require_role("admin")
def portal_admin_announcement_update(announcement_id):
    row = db.session.get(Announcement, announcement_id)
    if not row:
        flash("Announcement not found.", "error")
        return redirect_to_next("portal_admin_announcements_page")
    row.title = (request.form.get("title") or row.title).strip()[:220]
    row.body = (request.form.get("body") or row.body).strip()[:10000]
    row.is_published = (request.form.get("is_published") or "1").strip() == "1"
    row.show_on_public = (request.form.get("show_on_public") or "1").strip() == "1"
    row.show_on_member = (request.form.get("show_on_member") or "1").strip() == "1"
    row.published_at = datetime.utcnow() if row.is_published else None
    add_audit_log("update_announcement", f"announcement_id={row.id}")
    db.session.commit()
    flash("Announcement updated.", "success")
    return redirect_to_next("portal_admin_announcements_page")


@app.post("/portal/admin/announcements/<int:announcement_id>/delete")
@require_role("admin")
def portal_admin_announcement_delete(announcement_id):
    row = db.session.get(Announcement, announcement_id)
    if not row:
        flash("Announcement not found.", "error")
        return redirect_to_next("portal_admin_announcements_page")
    add_audit_log("delete_announcement", f"announcement_id={row.id}")
    db.session.delete(row)
    db.session.commit()
    flash("Announcement deleted.", "success")
    return redirect_to_next("portal_admin_announcements_page")


@app.post("/portal/inventory/checkout")
@require_role("member")
def portal_checkout():
    user = current_auth_user()
    member = current_user_member(user)
    if not member:
        flash("Your user account is not linked to a member profile.", "error")
        return redirect_to_next("portal_member_inventory")
    item_id = parse_positive_int(request.form.get("item_id"), default=0)
    qty = parse_positive_int(request.form.get("qty"), default=1)
    notes = (request.form.get("notes") or "").strip() or None
    due_date = parse_due_date(request.form.get("due_date"))
    item = db.session.get(Item, item_id)
    if not item:
        flash("Item not found.", "error")
        return redirect_to_next("portal_member_inventory")
    existing_open = get_open_checkout(member.id, item.id)
    if existing_open:
        flash("You already have this item checked out.", "error")
        return redirect_to_next("portal_member_checkouts")

    updated = (
        Item.query.filter(Item.id == item.id, Item.available_qty >= qty)
        .update({Item.available_qty: Item.available_qty - qty}, synchronize_session=False)
    )
    if updated != 1:
        db.session.rollback()
        db.session.refresh(item)
        flash(f"{item.name} only has {item.available_qty} available.", "error")
        return redirect_to_next("portal_member_inventory")

    now = datetime.utcnow()
    db.session.add(
        Transaction(
            member_id=member.id,
            user_id=user.id,
            item_id=item.id,
            action="checkout",
            qty=qty,
            status="OUT",
            timestamp=now,
            checkout_time=now,
            due_date=due_date,
            notes=notes,
            checkout_notes=notes,
        )
    )
    db.session.commit()
    flash(f"Checked out {qty} x {item.name}.", "success")
    return redirect_to_next("portal_member_checkouts")


@app.post("/portal/inventory/return")
@require_role("member")
def portal_return():
    user = current_auth_user()
    member = current_user_member(user)
    if not member:
        flash("Your user account is not linked to a member profile.", "error")
        return redirect_to_next("portal_member_checkouts")
    item_id = parse_positive_int(request.form.get("item_id"), default=0)
    qty = parse_positive_int(request.form.get("qty"), default=1)
    condition = (request.form.get("condition") or "").strip()
    notes = (request.form.get("notes") or "").strip() or None

    item = db.session.get(Item, item_id)
    if not item:
        flash("Item not found.", "error")
        return redirect_to_next("portal_member_checkouts")
    open_tx = get_open_checkout(member.id, item.id)
    if not open_tx:
        flash("No open checkout found for this item.", "error")
        return redirect_to_next("portal_member_checkouts")
    if qty != open_tx.qty:
        flash(f"Return quantity must match checked out quantity ({open_tx.qty}).", "error")
        return redirect_to_next("portal_member_checkouts")

    item.available_qty = min(item.total_qty, item.available_qty + qty)
    open_tx.status = "RETURNED"
    open_tx.return_time = datetime.utcnow()
    open_tx.return_condition = condition or "good"
    open_tx.return_notes = notes
    db.session.commit()
    flash(f"Returned {qty} x {item.name}.", "success")
    return redirect_to_next("portal_member_checkouts")


@app.post("/portal/print/request")
@require_role("member")
def portal_print_request():
    user = current_auth_user()
    member = current_user_member(user)
    printer_type = normalize_portal_printer_type(request.form.get("printer_type"))
    filament = (request.form.get("filament") or "").strip()
    if printer_type not in PRINT_REQUEST_PRINTERS:
        flash("Choose a valid printer queue (P1S-1..P1S-4 or H2S).", "error")
        return redirect_to_next("portal_member_prints")
    if not filament:
        flash("Filament type/color is required.", "error")
        return redirect_to_next("portal_member_prints")

    file_link = (request.form.get("file_link") or "").strip() or None
    material = (request.form.get("material") or "").strip() or None
    color = (request.form.get("color") or "").strip() or None
    infill_percent = parse_positive_int(request.form.get("infill_percent"), default=20)
    priority = (request.form.get("priority") or "normal").strip().lower() or "normal"
    deadline = parse_due_date(request.form.get("deadline"))
    notes = (request.form.get("notes") or "").strip() or None

    file_path = None
    file_upload = request.files.get("print_file")
    if file_upload and file_upload.filename:
        safe_name = secure_filename(file_upload.filename)
        if not safe_name or not allowed_gcode(safe_name):
            flash("Upload .gcode, .gco, or .3mf files only.", "error")
            return redirect_to_next("portal_member_prints")
        try:
            file_upload.stream.seek(0, os.SEEK_END)
            file_size = file_upload.stream.tell()
            file_upload.stream.seek(0)
        except Exception:
            file_size = 0
        if file_size > PRINT_REQUEST_MAX_UPLOAD_BYTES:
            max_mb = max(1, PRINT_REQUEST_MAX_UPLOAD_BYTES // (1024 * 1024))
            flash(f"File is too large. Max upload size is {max_mb} MB.", "error")
            return redirect_to_next("portal_member_prints")
        stored_name = f"printreq_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{safe_name}"
        output_path = UPLOAD_DIR / stored_name
        file_upload.save(output_path)
        file_path = str(output_path)

    if not file_path and not file_link:
        flash("Upload a file or provide a link.", "error")
        return redirect_to_next("portal_member_prints")

    db.session.add(
        PrintRequest(
            user_id=user.id,
            member_id=member.id if member else None,
            printer_type=printer_type,
            file_path=file_path,
            file_link=file_link,
            filament=filament[:160],
            material=material,
            color=color,
            infill_percent=min(infill_percent, 100),
            priority=priority[:40],
            deadline=deadline,
            notes=notes,
            status="submitted",
        )
    )
    db.session.commit()
    flash("3D print request submitted.", "success")
    return redirect_to_next("portal_member_prints")


@app.post("/portal/events/request")
@require_role("team_leader")
def portal_event_request():
    flash("Use Schedule Meeting for slot-based availability and room conflict prevention.", "info")
    return redirect(url_for("portal_member_schedule"))


@app.post("/portal/events/<int:event_id>/rsvp")
@require_role("member")
def portal_event_rsvp(event_id):
    user = current_auth_user()
    member = current_user_member(user)
    event = db.session.get(Event, event_id)
    if not event:
        flash("Event not found.", "error")
        return redirect_to_next("portal_member_calendar")
    existing = AttendanceRecord.query.filter_by(
        event_id=event.id,
        user_id=user.id,
        checkin_method="rsvp",
    ).first()
    if existing:
        flash("RSVP already recorded.", "info")
        return redirect_to_next("portal_member_calendar")
    db.session.add(
        AttendanceRecord(
            event_id=event.id,
            user_id=user.id,
            member_id=member.id if member else None,
            checkin_method="rsvp",
            checkin_time=datetime.utcnow(),
        )
    )
    db.session.commit()
    flash("RSVP saved.", "success")
    return redirect_to_next("portal_member_calendar")


@app.post("/portal/admin/members/create")
@require_role("admin")
def portal_admin_create_member():
    name = (request.form.get("name") or "").strip()
    email = (request.form.get("email") or "").strip().lower()
    member_class = (request.form.get("member_class") or "").strip() or "Member"
    create_user = (request.form.get("create_user") or "1").strip() in {"1", "true", "on", "yes"}
    role = normalize_role(request.form.get("role") or "member")
    password = request.form.get("password") or ""
    major = (request.form.get("major") or "").strip()[:120] or None
    grad_year_raw = (request.form.get("graduation_year") or "").strip()
    graduation_year = parse_non_negative_int(grad_year_raw, default=0) if grad_year_raw else 0
    nfc_uid = clean_tag_value(request.form.get("nfc_uid"))
    exec_title = (request.form.get("exec_title") or "").strip()[:160] or None
    exec_message = (request.form.get("exec_message") or "").strip()[:500] or None
    headshot_url = (request.form.get("headshot_url") or "").strip()[:500] or None
    first_name, last_name = split_name_parts(name)
    username_base = username_base_from_name(first_name, last_name)
    if graduation_year and (graduation_year < 1900 or graduation_year > 2100):
        graduation_year = 0

    if not name or "@" not in email:
        flash("Enter a valid member name and email.", "error")
        return redirect_to_next("portal_admin_members_page")
    if Member.query.filter(func.lower(Member.email) == email).first():
        flash("A member with that email already exists.", "error")
        return redirect_to_next("portal_admin_members_page")
    if nfc_uid and find_user_by_nfc_uid(nfc_uid):
        flash("That NFC UID is already assigned to another user.", "error")
        return redirect_to_next("portal_admin_members_page")
    if nfc_uid:
        member_uid_conflict = Member.query.filter(func.lower(Member.nfc_tag) == nfc_uid.lower()).first()
        if member_uid_conflict:
            flash("That NFC UID is already assigned to another member record.", "error")
            return redirect_to_next("portal_admin_members_page")

    member = Member(
        name=name[:120],
        email=email[:160],
        member_class=member_class[:80] or "Member",
        nfc_tag=nfc_uid or None,
    )
    db.session.add(member)
    db.session.flush()

    linked_user = None
    if create_user:
        existing_user = User.query.filter(func.lower(User.email) == email).first()
        if existing_user:
            if existing_user.member_id and existing_user.member_id != member.id:
                db.session.rollback()
                flash("That email already belongs to a different user/member link.", "error")
                return redirect_to_next("portal_admin_members_page")
            existing_user.member_id = member.id
            existing_user.name = name[:160]
            existing_user.is_active = True
            existing_user.role = role
            if not (existing_user.username or "").strip():
                existing_user.username = make_unique_username(
                    username_base or email.split("@", 1)[0],
                    exclude_user_id=existing_user.id,
                )
            existing_user.major = major
            existing_user.graduation_year = graduation_year or None
            existing_user.nfc_uid = nfc_uid or None
            existing_user.exec_title = exec_title
            existing_user.exec_message = exec_message
            existing_user.headshot_url = headshot_url
            if password:
                if len(password) < 8:
                    db.session.rollback()
                    flash("Password must be at least 8 characters.", "error")
                    return redirect_to_next("portal_admin_members_page")
                existing_user.password_hash = generate_password_hash(password)
            linked_user = existing_user
        else:
            if len(password) < 8:
                db.session.rollback()
                flash("Password is required (min 8 chars) when creating a new login account.", "error")
                return redirect_to_next("portal_admin_members_page")
            linked_user = User(
                name=name[:160],
                email=email[:160],
                username=make_unique_username(username_base or email.split("@", 1)[0]),
                password_hash=generate_password_hash(password),
                role=role,
                is_active=True,
                nfc_uid=nfc_uid or None,
                major=major,
                graduation_year=graduation_year or None,
                exec_title=exec_title,
                exec_message=exec_message,
                headshot_url=headshot_url,
                member_id=member.id,
            )
            db.session.add(linked_user)

    add_audit_log(
        "create_member",
        f"member_id={member.id} email={member.email} create_user={bool(linked_user)}",
    )
    db.session.commit()
    if linked_user:
        flash(f"Member created with ID #{member.id} and linked user account.", "success")
    else:
        flash(f"Member created with ID #{member.id}.", "success")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/users/create")
@require_role("admin")
def portal_admin_create_user():
    name = (request.form.get("name") or "").strip()
    email = (request.form.get("email") or "").strip().lower()
    password = request.form.get("password") or ""
    role = normalize_role(request.form.get("role") or "member")
    member_id = parse_positive_int(request.form.get("member_id"), default=0)
    major = (request.form.get("major") or "").strip()[:120] or None
    grad_year_raw = (request.form.get("graduation_year") or "").strip()
    graduation_year = parse_non_negative_int(grad_year_raw, default=0) if grad_year_raw else 0
    nfc_uid = clean_tag_value(request.form.get("nfc_uid"))
    exec_title = (request.form.get("exec_title") or "").strip()[:160] or None
    exec_message = (request.form.get("exec_message") or "").strip()[:500] or None
    headshot_url = (request.form.get("headshot_url") or "").strip()[:500] or None
    first_name, last_name = split_name_parts(name)
    username_base = username_base_from_name(first_name, last_name)
    if graduation_year and (graduation_year < 1900 or graduation_year > 2100):
        graduation_year = 0

    if not name or "@" not in email or len(password) < 8:
        flash("Enter valid name, email, and password (min 8 chars).", "error")
        return redirect_to_next("portal_admin_members_page")
    if User.query.filter(func.lower(User.email) == email).first():
        flash("User with that email already exists.", "error")
        return redirect_to_next("portal_admin_members_page")
    if nfc_uid and find_user_by_nfc_uid(nfc_uid):
        flash("That NFC UID is already assigned to another user.", "error")
        return redirect_to_next("portal_admin_members_page")
    if nfc_uid:
        member_uid_conflict = Member.query.filter(func.lower(Member.nfc_tag) == nfc_uid.lower()).first()
        if member_uid_conflict:
            flash("That NFC UID is already assigned to a member record.", "error")
            return redirect_to_next("portal_admin_members_page")

    member = db.session.get(Member, member_id) if member_id else None
    if not member:
        member = Member.query.filter(func.lower(Member.email) == email).first()
    if member:
        linked_existing = User.query.filter(User.member_id == member.id).first()
        if linked_existing:
            flash(f"Member ID #{member.id} is already linked to {linked_existing.email}.", "error")
            return redirect_to_next("portal_admin_members_page")

    new_user = User(
        name=name[:160],
        email=email[:160],
        username=make_unique_username(username_base or email.split("@", 1)[0]),
        password_hash=generate_password_hash(password),
        role=role,
        is_active=True,
        nfc_uid=nfc_uid or None,
        major=major,
        graduation_year=graduation_year or None,
        exec_title=exec_title,
        exec_message=exec_message,
        headshot_url=headshot_url,
        member_id=member.id if member else None,
    )
    if member and nfc_uid:
        member.nfc_tag = nfc_uid
    db.session.add(new_user)
    add_audit_log("create_user", f"{new_user.email} role={new_user.role}")
    db.session.commit()
    flash("User created.", "success")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/users/<int:user_id>/role")
@require_role("admin")
def portal_admin_update_user_role(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("portal_admin_members_page")
    previous_member = user.member
    previous_nfc_uid = clean_tag_value(user.nfc_uid)
    role = normalize_role(request.form.get("role") or "member")
    active_raw = (request.form.get("is_active") or "1").strip()
    name = (request.form.get("name") or user.name).strip()
    email = (request.form.get("email") or user.email).strip().lower()
    username_raw = (request.form.get("username") or "").strip().lower()
    member_id_raw = (request.form.get("member_id") or "").strip()
    member_id = parse_positive_int(member_id_raw, default=0) if member_id_raw else 0
    member_row = db.session.get(Member, member_id) if member_id else None
    if member_id and not member_row:
        flash("Member ID not found.", "error")
        return redirect_to_next("portal_admin_members_page")
    if member_row:
        member_conflict = User.query.filter(User.member_id == member_row.id, User.id != user.id).first()
        if member_conflict:
            flash("That member ID is already linked to another user.", "error")
            return redirect_to_next("portal_admin_members_page")
    major = (request.form.get("major") or "").strip()[:120] or None
    grad_year_raw = (request.form.get("graduation_year") or "").strip()
    graduation_year = parse_non_negative_int(grad_year_raw, default=0) if grad_year_raw else 0
    nfc_uid = clean_tag_value(request.form.get("nfc_uid"))
    exec_title = (request.form.get("exec_title") or "").strip()[:160] or None
    exec_message = (request.form.get("exec_message") or "").strip()[:500] or None
    headshot_url = (request.form.get("headshot_url") or "").strip()[:500] or None
    if graduation_year and (graduation_year < 1900 or graduation_year > 2100):
        graduation_year = 0
    if "@" not in email:
        flash("Enter a valid email.", "error")
        return redirect_to_next("portal_admin_members_page")
    duplicate = (
        User.query.filter(func.lower(User.email) == email, User.id != user.id)
        .order_by(User.id.asc())
        .first()
    )
    if duplicate:
        flash("Another user already has that email.", "error")
        return redirect_to_next("portal_admin_members_page")
    requested_username = normalize_text_key(username_raw or "")
    if requested_username:
        username_conflict = (
            User.query.filter(
                func.lower(func.coalesce(User.username, "")) == requested_username.lower(),
                User.id != user.id,
            )
            .order_by(User.id.asc())
            .first()
        )
        if username_conflict:
            flash("That username is already in use.", "error")
            return redirect_to_next("portal_admin_members_page")
    uid_conflict = find_user_by_nfc_uid(nfc_uid, exclude_user_id=user.id) if nfc_uid else None
    if uid_conflict:
        flash("That NFC UID is already assigned to another user.", "error")
        return redirect_to_next("portal_admin_members_page")
    if nfc_uid:
        member_uid_conflict = (
            Member.query.filter(
                func.lower(Member.nfc_tag) == nfc_uid.lower(),
                Member.id != (member_row.id if member_row else 0),
            )
            .order_by(Member.id.asc())
            .first()
        )
        if member_uid_conflict:
            flash("That NFC UID is already assigned to another member record.", "error")
            return redirect_to_next("portal_admin_members_page")

    user.name = name[:160] or user.name
    user.email = email[:160]
    if requested_username:
        user.username = requested_username[:80]
    elif not (user.username or "").strip():
        first_name, last_name = split_name_parts(user.name)
        user.username = make_unique_username(
            username_base_from_name(first_name, last_name) or user.email.split("@", 1)[0],
            exclude_user_id=user.id,
        )
    user.role = role
    user.is_active = active_raw == "1"
    user.member_id = member_row.id if member_row else None
    user.nfc_uid = nfc_uid or None
    user.major = major
    user.graduation_year = graduation_year or None
    user.exec_title = exec_title
    user.exec_message = exec_message
    user.headshot_url = headshot_url
    if previous_member and (not member_row or previous_member.id != member_row.id):
        if clean_tag_value(previous_member.nfc_tag).lower() == previous_nfc_uid.lower():
            previous_member.nfc_tag = None
    if member_row:
        member_row.nfc_tag = nfc_uid or None
    add_audit_log(
        "update_user_role",
        (
            f"user_id={user.id} role={role} active={user.is_active} "
            f"email={user.email} member_id={user.member_id or ''} nfc_uid={user.nfc_uid or ''}"
        ),
    )
    db.session.commit()
    flash("User updated.", "success")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/users/<int:user_id>/reset-password")
@require_role("admin")
def portal_admin_reset_user_password(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("portal_admin_members_page")
    provided = (request.form.get("new_password") or "").strip()
    if provided and len(provided) < 8:
        flash("Password must be at least 8 characters.", "error")
        return redirect_to_next("portal_admin_members_page")
    new_password = provided or f"ASME-{uuid4().hex[:10]}"
    user.password_hash = generate_password_hash(new_password)
    add_audit_log("reset_user_password", f"user_id={user.id}")
    db.session.commit()
    flash(f"Password reset for {user.email}. Temporary password: {new_password}", "info")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/users/<int:user_id>/invite-link")
@require_role("admin")
def portal_admin_user_invite_link(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("portal_admin_members_page")
    token = secrets.token_urlsafe(32)
    expires_at = datetime.utcnow() + timedelta(hours=72)
    db.session.add(PasswordResetToken(user_id=user.id, token=token, expires_at=expires_at))
    add_audit_log("create_user_invite_link", f"user_id={user.id}")
    db.session.commit()
    reset_link = url_for("reset_password_page", token=token, _external=True)
    flash(f"Invite/reset link for {user.email}: {reset_link}", "info")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/users/<int:user_id>/delete")
@require_role("admin")
def portal_admin_delete_user(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("portal_admin_members_page")

    admin_user = current_auth_user()
    if admin_user and admin_user.id == user.id:
        flash("You cannot delete your own account while logged in.", "error")
        return redirect_to_next("portal_admin_members_page")

    previous_uid = clean_tag_value(user.nfc_uid)
    related_member = user.member

    # Always deactivate active NFC mappings first to avoid stale tag ownership.
    for row in NFCTag.query.filter_by(user_id=user.id, active=True).all():
        row.active = False
        row.unassigned_at = datetime.utcnow()

    user.nfc_uid = None
    if related_member and previous_uid and clean_tag_value(related_member.nfc_tag).lower() == previous_uid.lower():
        related_member.nfc_tag = None

    references = {
        "transactions": Transaction.query.filter(Transaction.user_id == user.id).count(),
        "print_requests": PrintRequest.query.filter(PrintRequest.user_id == user.id).count(),
        "events_created": Event.query.filter(Event.created_by_user_id == user.id).count(),
        "events_requested": Event.query.filter(Event.requested_by_user_id == user.id).count(),
        "attendance": AttendanceRecord.query.filter(AttendanceRecord.user_id == user.id).count(),
        "contact_messages": ContactMessage.query.filter(ContactMessage.user_id == user.id).count(),
        "audit_logs": AuditLog.query.filter(AuditLog.admin_user_id == user.id).count(),
        "password_tokens": PasswordResetToken.query.filter(PasswordResetToken.user_id == user.id).count(),
        "nfc_history": NFCTag.query.filter(NFCTag.user_id == user.id).count(),
    }
    total_refs = sum(references.values())

    if total_refs == 0:
        db.session.delete(user)
        add_audit_log("delete_user", f"user_id={user_id} hard_delete=1")
        db.session.commit()
        flash("User deleted.", "success")
        return redirect_to_next("portal_admin_members_page")

    user.is_active = False
    add_audit_log("delete_user", f"user_id={user_id} hard_delete=0 refs={total_refs}")
    db.session.commit()
    flash("User has history, so the account was deactivated instead of hard-deleted.", "info")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/members/import-roster")
@require_role("admin")
def portal_admin_import_roster():
    roster_file = request.files.get("roster_file")
    if not roster_file or not roster_file.filename:
        flash("Upload a roster PDF file.", "error")
        return redirect_to_next("portal_admin_members_page")
    file_name = (roster_file.filename or "").strip().lower()
    if not file_name.endswith(".pdf"):
        flash("Roster import currently supports PDF files only.", "error")
        return redirect_to_next("portal_admin_members_page")
    try:
        pdf_bytes = roster_file.read()
        entries = parse_roster_pdf_entries(pdf_bytes)
    except Exception as exc:
        flash(f"Failed to parse roster PDF: {str(exc)[:220]}", "error")
        return redirect_to_next("portal_admin_members_page")
    if not entries:
        flash("No roster members were found in the uploaded file.", "error")
        return redirect_to_next("portal_admin_members_page")

    role = normalize_role(request.form.get("role") or "member")
    if role == "admin":
        role = "member"
    member_class = (request.form.get("member_class") or "Member").strip() or "Member"
    reset_existing = (request.form.get("reset_existing_passwords") or "1").strip() in {"1", "true", "yes", "on"}

    result = import_roster_entries(
        entries,
        default_role=role,
        member_class=member_class,
        reset_existing_passwords=reset_existing,
    )
    credentials = result.get("credentials") or []
    if not credentials:
        flash("Roster processed, but no account credentials were generated.", "info")
        return redirect_to_next("portal_admin_members_page")

    csv_file = save_roster_credentials_csv(credentials)
    session["latest_roster_credentials_file"] = csv_file
    add_audit_log(
        "bulk_roster_import",
        (
            f"entries={len(entries)} created_members={result['created_members']} "
            f"created_users={result['created_users']} updated_users={result['updated_users']} "
            f"reset_passwords={result['reset_passwords']} file={csv_file}"
        ),
    )
    db.session.commit()
    flash(
        (
            f"Roster imported: {len(entries)} rows, {result['created_members']} members created, "
            f"{result['created_users']} users created, {result['updated_users']} users updated. "
            f"Credentials CSV is ready for download."
        ),
        "success",
    )
    return redirect_to_next("portal_admin_members_page")


@app.get("/portal/admin/members/import-credentials/<filename>")
@require_role("admin")
def portal_admin_download_roster_credentials(filename):
    safe_name = Path(filename or "").name
    if not safe_name:
        flash("Invalid credentials file name.", "error")
        return redirect(url_for("portal_admin_members_page"))
    file_path = ROSTER_CREDENTIALS_DIR / safe_name
    if not file_path.exists() or not file_path.is_file():
        flash("Credentials file not found.", "error")
        return redirect(url_for("portal_admin_members_page"))
    return send_file(
        file_path,
        as_attachment=True,
        download_name=safe_name,
        mimetype="text/csv",
    )


@app.post("/portal/admin/nfc/assign")
@require_role("admin")
def portal_admin_assign_nfc():
    user_id = parse_positive_int(request.form.get("user_id"), default=0)
    tag_uid = clean_tag_value(request.form.get("tag_uid"))
    notes = (request.form.get("notes") or "").strip() or None
    if not user_id or not tag_uid:
        flash("User and tag UID are required.", "error")
        return redirect_to_next("portal_admin_members_page")
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("portal_admin_members_page")
    uid_conflict = find_user_by_nfc_uid(tag_uid, exclude_user_id=user.id)
    if uid_conflict:
        flash("That tag UID is already assigned to another user account.", "error")
        return redirect_to_next("portal_admin_members_page")

    item_tag_conflict = ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag_uid.lower()).first()
    if item_tag_conflict:
        flash("That UID is already assigned to an inventory item tag.", "error")
        return redirect_to_next("portal_admin_members_page")
    legacy_item_conflict = Item.query.filter(func.lower(Item.nfc_tag) == tag_uid.lower()).first()
    if legacy_item_conflict:
        flash("That UID is already assigned to an inventory item.", "error")
        return redirect_to_next("portal_admin_members_page")

    existing_active = NFCTag.query.filter(func.lower(NFCTag.tag_uid) == tag_uid.lower(), NFCTag.active.is_(True)).first()
    if existing_active and existing_active.user_id != user.id:
        flash("That tag UID is already assigned.", "error")
        return redirect_to_next("portal_admin_members_page")

    current_active_for_user = NFCTag.query.filter_by(user_id=user.id, active=True).all()
    for row in current_active_for_user:
        row.active = False
        row.unassigned_at = datetime.utcnow()

    if existing_active and existing_active.user_id == user.id:
        existing_active.notes = notes
        existing_active.assigned_at = datetime.utcnow()
    else:
        admin = current_auth_user()
        db.session.add(
            NFCTag(
                tag_uid=tag_uid,
                user_id=user.id,
                active=True,
                assigned_at=datetime.utcnow(),
                assigned_by_user_id=admin.id if admin else None,
                notes=notes,
            )
        )
    user.nfc_uid = tag_uid
    add_audit_log("assign_nfc", f"user_id={user.id} tag_uid={tag_uid}")
    db.session.commit()
    flash("NFC tag assignment updated.", "success")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/nfc/unassign/<int:tag_id>")
@require_role("admin")
def portal_admin_unassign_nfc(tag_id):
    row = db.session.get(NFCTag, tag_id)
    if not row:
        flash("NFC tag assignment not found.", "error")
        return redirect_to_next("portal_admin_members_page")
    row.active = False
    row.unassigned_at = datetime.utcnow()
    if row.user and clean_tag_value(row.user.nfc_uid).lower() == clean_tag_value(row.tag_uid).lower():
        row.user.nfc_uid = None
    add_audit_log("unassign_nfc", f"tag_id={row.id} uid={row.tag_uid}")
    db.session.commit()
    flash("Tag unassigned.", "success")
    return redirect_to_next("portal_admin_members_page")


@app.post("/portal/admin/events/create")
@require_role("admin")
def portal_admin_create_event():
    user = current_auth_user()
    title = (request.form.get("title") or "").strip()
    description = (request.form.get("description") or "").strip() or None
    room = normalize_meeting_room(request.form.get("location"))
    start_time = parse_datetime_local(request.form.get("start_time"))
    end_time = parse_datetime_local(request.form.get("end_time"))
    status = (request.form.get("status") or "scheduled").strip().lower()
    if status not in {"requested", "scheduled", "cancelled"}:
        status = "scheduled"
    if not title or not room or not start_time or not end_time or end_time <= start_time:
        flash("Enter valid title, room, start, and end times.", "error")
        return redirect_to_next("portal_admin_attendance_page")
    if status in {"requested", "scheduled"}:
        conflict = find_event_room_conflict(room, start_time, end_time)
        if conflict:
            flash(
                f"{room} conflicts with '{conflict.title}' "
                f"({conflict.start_time.strftime('%Y-%m-%d %H:%M')} - {conflict.end_time.strftime('%H:%M')}).",
                "error",
            )
            return redirect_to_next("portal_admin_attendance_page")

    db.session.add(
        Event(
            title=title[:220],
            description=description,
            location=room,
            status=status,
            start_time=start_time,
            end_time=end_time,
            created_by_user_id=user.id,
        )
    )
    add_audit_log("create_event", title[:220])
    db.session.commit()
    flash("Event saved.", "success")
    return redirect_to_next("portal_admin_attendance_page")


@app.post("/portal/admin/events/<int:event_id>/status")
@require_role("admin")
def portal_admin_event_status_update(event_id):
    event = db.session.get(Event, event_id)
    if not event:
        flash("Event not found.", "error")
        return redirect_to_next("portal_admin_attendance_page")

    status = (request.form.get("status") or event.status or "scheduled").strip().lower()
    if status not in {"requested", "scheduled", "cancelled"}:
        status = event.status or "scheduled"

    title = (request.form.get("title") or event.title or "").strip()
    start_time = parse_datetime_local(request.form.get("start_time")) or event.start_time
    end_time = parse_datetime_local(request.form.get("end_time")) or event.end_time
    room = normalize_meeting_room(request.form.get("location") or event.location)
    if not title or not room or not start_time or not end_time or end_time <= start_time:
        flash("Enter valid title, room, start, and end times.", "error")
        return redirect_to_next("portal_admin_attendance_page")

    if status in {"requested", "scheduled"}:
        conflict = find_event_room_conflict(room, start_time, end_time, exclude_event_id=event.id)
        if conflict:
            flash(
                f"Cannot update event. {room} conflicts with '{conflict.title}' "
                f"({conflict.start_time.strftime('%Y-%m-%d %H:%M')} - {conflict.end_time.strftime('%H:%M')}).",
                "error",
            )
            return redirect_to_next("portal_admin_attendance_page")

    event.title = title[:220]
    event.location = room
    event.start_time = start_time
    event.end_time = end_time
    event.status = status
    add_audit_log("update_event_status", f"event_id={event.id} status={status}")
    db.session.commit()
    flash("Event updated.", "success")
    return redirect_to_next("portal_admin_attendance_page")


@app.post("/portal/admin/attendance/checkin")
@require_role("admin")
def portal_admin_attendance_checkin():
    event_id = parse_positive_int(request.form.get("event_id"), default=0)
    tag_uid = clean_tag_value(request.form.get("tag_uid"))
    manual_user_id = parse_positive_int(request.form.get("user_id"), default=0)
    requested_method = (request.form.get("checkin_method") or "").strip().lower()
    event = db.session.get(Event, event_id)
    if not event:
        flash("Event not found.", "error")
        return redirect_to_next("portal_admin_attendance_page")

    user = None
    member = None
    checkin_method = "manual"

    if tag_uid:
        user_from_uid = resolve_user_from_tag_uid(tag_uid)
        if not user_from_uid:
            flash("Tag UID not assigned to an active user.", "error")
            return redirect_to_next("portal_admin_attendance_page")
        user = user_from_uid
        member = current_user_member(user)
        checkin_method = requested_method if requested_method in {"kiosk", "nfc"} else "nfc"
    elif manual_user_id:
        user = db.session.get(User, manual_user_id)
        if not user:
            flash("Selected user not found.", "error")
            return redirect_to_next("portal_admin_attendance_page")
        member = current_user_member(user)
    else:
        flash("Provide tag UID or select a user.", "error")
        return redirect_to_next("portal_admin_attendance_page")

    existing_query = AttendanceRecord.query.filter_by(event_id=event.id)
    if user:
        existing_query = existing_query.filter(AttendanceRecord.user_id == user.id)
    elif member:
        existing_query = existing_query.filter(AttendanceRecord.member_id == member.id)
    existing_row = existing_query.order_by(AttendanceRecord.id.desc()).first()
    if existing_row:
        flash("This person is already checked in for the selected event.", "info")
        return redirect_to_next("portal_admin_attendance_page")

    db.session.add(
        AttendanceRecord(
            event_id=event.id,
            user_id=user.id if user else None,
            member_id=member.id if member else None,
            tag_uid=tag_uid or None,
            checkin_method=checkin_method,
            checkin_time=datetime.utcnow(),
        )
    )
    add_audit_log("attendance_checkin", f"event_id={event.id} method={checkin_method}")
    db.session.commit()
    flash(f"You have been marked present for {event.title}.", "success")
    return redirect_to_next("portal_admin_attendance_page")


@app.get("/portal/admin/attendance/export.csv")
@require_role("admin")
def portal_admin_attendance_export():
    rows = AttendanceRecord.query.order_by(AttendanceRecord.checkin_time.desc(), AttendanceRecord.id.desc()).all()
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(["record_id", "event_id", "event_title", "user_id", "user_name", "tag_uid", "method", "checkin_time"])
    for row in rows:
        writer.writerow(
            [
                row.id,
                row.event_id,
                row.event.title if row.event else "",
                row.user_id or "",
                row.user.name if row.user else (row.member.name if row.member else ""),
                row.tag_uid or "",
                row.checkin_method or "",
                row.checkin_time.isoformat() if row.checkin_time else "",
            ]
        )
    content = buffer.getvalue()
    return Response(
        content,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=attendance_records.csv"},
    )


@app.post("/portal/admin/inventory/item/save")
@require_role("admin")
def portal_admin_inventory_item_save():
    item_id = parse_positive_int(request.form.get("item_id"), default=0)
    name = (request.form.get("name") or "").strip()
    if not name:
        flash("Item name is required.", "error")
        return redirect_to_next("portal_admin_inventory_page")
    item = db.session.get(Item, item_id) if item_id else None
    if not item:
        item = Item(name=name[:160], total_qty=0, available_qty=0)
        db.session.add(item)

    nfc_tag = clean_tag_value(request.form.get("nfc_tag"))
    tag_error = validate_item_tag_uid(nfc_tag, current_item_id=item.id if item and item.id else None)
    if tag_error:
        flash(tag_error, "error")
        return redirect_to_next("portal_admin_inventory_page")

    submitted_item_code = normalize_inventory_item_code(request.form.get("item_code"))
    if submitted_item_code:
        existing_item_code = Item.query.filter(
            func.lower(func.coalesce(Item.item_code, "")) == submitted_item_code.lower(),
            Item.id != (item.id or 0),
        ).first()
        if existing_item_code:
            flash("Item ID is already in use. Choose a different Item ID.", "error")
            return redirect_to_next("portal_admin_inventory_page")

    submitted_tracker_id = normalize_treasury_tracker_id(request.form.get("treasury_tracker_id"))
    if submitted_tracker_id:
        existing_tracker_id = Item.query.filter(
            func.lower(func.coalesce(Item.treasury_tracker_id, "")) == submitted_tracker_id.lower(),
            Item.id != (item.id or 0),
        ).first()
        if existing_tracker_id:
            flash("Tracker ID is already in use. Choose a different tracker ID.", "error")
            return redirect_to_next("portal_admin_inventory_page")

    total_qty = parse_non_negative_int(request.form.get("total_qty"), default=max(item.total_qty, 0))
    available_qty = parse_non_negative_int(request.form.get("available_qty"), default=min(item.available_qty, total_qty))
    item_type = (request.form.get("item_type") or "").strip().lower() or "tool"
    if item_type not in ITEM_TYPES:
        item_type = "tool"
    is_consumable = item_type == "consumable"
    active = (request.form.get("active") or "1").strip() in {"1", "true", "yes", "on"}
    min_stock_threshold = parse_non_negative_int(request.form.get("min_stock_threshold"), default=0)
    item.name = name[:160]
    item.description = (request.form.get("description") or "").strip()[:300] or None
    item.category = (request.form.get("category") or "").strip()[:80] or None
    item.location = (request.form.get("location") or "").strip()[:120] or None
    item.item_condition = (request.form.get("item_condition") or "").strip()[:120] or None
    item.notes = (request.form.get("notes") or "").strip()[:500] or None
    item.photo_url = (request.form.get("photo_url") or "").strip()[:500] or None
    item.item_type = item_type
    item.is_consumable = is_consumable
    item.active = active
    item.min_stock_threshold = min_stock_threshold
    item.total_qty = max(total_qty, 0)
    item.available_qty = max(0, min(available_qty, item.total_qty))
    if submitted_item_code:
        item.item_code = submitted_item_code
    if submitted_tracker_id:
        item.treasury_tracker_id = submitted_tracker_id
    ensure_item_tracking_ids(item)
    set_item_primary_nfc(item, nfc_tag, source="inventory_admin")

    add_audit_log(
        "save_inventory_item",
        (
            f"item_id={item.id or 'new'} name={item.name} "
            f"item_code={item.item_code} tracker_id={item.treasury_tracker_id}"
        ),
    )
    db.session.commit()
    flash("Inventory item saved.", "success")
    return redirect_to_next("portal_admin_inventory_page")


@app.post("/portal/admin/inventory/adjust")
@require_role("admin")
def portal_admin_inventory_adjust():
    item_id = parse_positive_int(request.form.get("item_id"), default=0)
    total_qty = parse_non_negative_int(request.form.get("total_qty"), default=0)
    available_qty = parse_non_negative_int(request.form.get("available_qty"), default=0)
    notes = (request.form.get("notes") or "").strip() or "Admin quantity correction"
    item = db.session.get(Item, item_id)
    if not item:
        flash("Item not found.", "error")
        return redirect_to_next("portal_admin_inventory_page")

    item.total_qty = max(0, total_qty)
    item.available_qty = max(0, min(available_qty, item.total_qty))

    admin_user = current_auth_user()
    admin_member = current_user_member(admin_user)
    if admin_member:
        db.session.add(
            Transaction(
                member_id=admin_member.id,
                user_id=admin_user.id,
                item_id=item.id,
                action="return",
                qty=0,
                status="RETURNED",
                timestamp=datetime.utcnow(),
                return_time=datetime.utcnow(),
                return_condition="admin-correction",
                return_notes=notes[:300],
                notes=notes[:300],
            )
        )
    add_audit_log("adjust_inventory", f"item_id={item.id} total={item.total_qty} available={item.available_qty}")
    db.session.commit()
    flash("Inventory counts updated.", "success")
    return redirect_to_next("portal_admin_inventory_page")


@app.post("/portal/admin/inventory/import")
@require_role("admin")
def portal_admin_inventory_bulk_import():
    upload_file = request.files.get("inventory_file")
    try:
        rows = parse_bulk_inventory_rows(upload_file)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect_to_next("portal_admin_inventory_page")

    created_count = 0
    updated_count = 0
    skipped_count = 0
    errors = []

    for index, row in enumerate(rows, start=2):
        if not any(str(value).strip() for value in (row or {}).values()):
            continue

        name = bulk_inventory_row_value(row, "name", "item", "item_name")
        if not name:
            skipped_count += 1
            errors.append(f"Row {index}: missing item name.")
            continue

        item_code_input = normalize_inventory_item_code(
            bulk_inventory_row_value(row, "item_code", "item id", "inventory_id", "asset_id")
        )
        tracker_input = normalize_treasury_tracker_id(
            bulk_inventory_row_value(
                row,
                "treasury_tracker_id",
                "tracker_id",
                "tracker",
                "treasury_id",
            )
        )
        nfc_tag_input = clean_tag_value(
            bulk_inventory_row_value(row, "nfc_tag", "nfc", "nfc_uid", "tag_uid", "tag")
        )

        item = None
        if item_code_input:
            item = Item.query.filter(
                func.lower(func.coalesce(Item.item_code, "")) == item_code_input.lower()
            ).first()
        if not item and nfc_tag_input:
            item = Item.query.filter(func.lower(func.coalesce(Item.nfc_tag, "")) == nfc_tag_input.lower()).first()
            if not item:
                mapped_tag = ItemTag.query.filter(func.lower(ItemTag.tag_value) == nfc_tag_input.lower()).first()
                if mapped_tag:
                    item = mapped_tag.item
        if not item:
            item = (
                Item.query.filter(func.lower(Item.name) == name.lower())
                .order_by(Item.id.asc())
                .first()
            )

        is_new = item is None
        if is_new:
            item = Item(name=name[:160], total_qty=0, available_qty=0)

        if item_code_input:
            item_code_conflict = Item.query.filter(
                func.lower(func.coalesce(Item.item_code, "")) == item_code_input.lower(),
                Item.id != (item.id or 0),
            ).first()
            if item_code_conflict:
                skipped_count += 1
                errors.append(f"Row {index}: item code '{item_code_input}' is already used.")
                continue

        if tracker_input:
            tracker_conflict = Item.query.filter(
                func.lower(func.coalesce(Item.treasury_tracker_id, "")) == tracker_input.lower(),
                Item.id != (item.id or 0),
            ).first()
            if tracker_conflict:
                skipped_count += 1
                errors.append(f"Row {index}: tracker ID '{tracker_input}' is already used.")
                continue

        tag_error = validate_item_tag_uid(nfc_tag_input, current_item_id=item.id if item and item.id else None)
        if tag_error:
            skipped_count += 1
            errors.append(f"Row {index}: {tag_error}")
            continue

        total_qty = parse_non_negative_int(
            bulk_inventory_row_value(row, "total_qty", "total", "quantity_total", "qty_total"),
            default=max(item.total_qty, 0),
        )
        available_raw = bulk_inventory_row_value(row, "available_qty", "available", "quantity_available", "qty_available")
        if available_raw:
            available_qty = parse_non_negative_int(available_raw, default=total_qty)
        else:
            available_qty = min(max(item.available_qty, 0), total_qty)

        item_type_raw = (bulk_inventory_row_value(row, "item_type", "type", "inventory_type") or "").strip().lower()
        if item_type_raw not in ITEM_TYPES:
            item_type_raw = item.item_type if item.item_type in ITEM_TYPES else "tool"
        active_raw = bulk_inventory_row_value(row, "active", "is_active", "enabled")
        min_threshold = parse_non_negative_int(
            bulk_inventory_row_value(row, "min_stock_threshold", "min_stock", "reorder_threshold"),
            default=max(item.min_stock_threshold or 0, 0),
        )

        item.name = name[:160]
        item.category = bulk_inventory_row_value(row, "category")[:80] or None
        item.location = bulk_inventory_row_value(row, "location", "bin", "shelf")[:120] or None
        item.description = bulk_inventory_row_value(row, "description")[:300] or None
        item.item_condition = bulk_inventory_row_value(row, "item_condition", "condition")[:120] or None
        item.notes = bulk_inventory_row_value(row, "notes")[:500] or None
        item.photo_url = bulk_inventory_row_value(row, "photo_url", "photo", "image_url")[:500] or None
        item.item_type = item_type_raw
        item.is_consumable = item_type_raw == "consumable"
        item.active = parse_bool_flag(active_raw, default=True if is_new else bool(item.active))
        item.min_stock_threshold = min_threshold
        item.total_qty = max(total_qty, 0)
        item.available_qty = max(0, min(available_qty, item.total_qty))
        if item_code_input:
            item.item_code = item_code_input
        if tracker_input:
            item.treasury_tracker_id = tracker_input

        if is_new:
            db.session.add(item)
        ensure_item_tracking_ids(item)
        set_item_primary_nfc(item, nfc_tag_input, source="bulk_inventory_import")

        if is_new:
            created_count += 1
        else:
            updated_count += 1

    if not created_count and not updated_count:
        db.session.rollback()
        flash("Bulk import made no changes. Fix file issues and try again.", "error")
    else:
        add_audit_log(
            "bulk_inventory_import",
            f"rows={len(rows)} created={created_count} updated={updated_count} skipped={skipped_count}",
        )
        db.session.commit()
        flash(
            (
                f"Bulk import complete: {created_count} created, {updated_count} updated, "
                f"{skipped_count} skipped."
            ),
            "success",
        )

    if errors:
        preview = "; ".join(errors[:4])
        if len(errors) > 4:
            preview += f"; +{len(errors) - 4} more"
        flash(preview, "info")

    return redirect_to_next("portal_admin_inventory_page")


@app.get("/portal/admin/inventory/items-tags.pdf")
@require_role("admin")
def portal_admin_inventory_item_nfc_pdf():
    items = Item.query.order_by(Item.name.asc(), Item.id.asc()).all()
    tag_map = get_item_primary_tag_map()
    now = datetime.now()

    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
    except Exception:
        flash("PDF export requires reportlab. Add reportlab to requirements and redeploy.", "error")
        return redirect(url_for("portal_admin_inventory_page", tab="items"))

    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    page_width, page_height = letter
    margin_left = 36
    y = page_height - 44

    def draw_header():
        nonlocal y
        pdf.setFont("Helvetica-Bold", 13)
        pdf.drawString(margin_left, y, "ASME Inventory NFC / Tracker List")
        pdf.setFont("Helvetica", 9)
        pdf.drawString(margin_left, y - 13, f"Generated: {now.strftime('%Y-%m-%d %H:%M')}")
        y -= 30
        pdf.setFont("Helvetica-Bold", 8.5)
        pdf.drawString(margin_left, y, "Item Code")
        pdf.drawString(margin_left + 76, y, "Item Name")
        pdf.drawString(margin_left + 292, y, "NFC ID")
        pdf.drawString(margin_left + 430, y, "Tracker ID")
        y -= 8
        pdf.line(margin_left, y, page_width - margin_left, y)
        y -= 12

    draw_header()
    pdf.setFont("Helvetica", 8.5)
    for item in items:
        if y < 36:
            pdf.showPage()
            y = page_height - 44
            draw_header()
            pdf.setFont("Helvetica", 8.5)
        pdf.drawString(margin_left, y, (item.item_code or "")[:18])
        pdf.drawString(margin_left + 76, y, (item.name or "")[:42])
        pdf.drawString(margin_left + 292, y, (get_item_primary_nfc(item, tag_map) or "-")[:25])
        pdf.drawString(margin_left + 430, y, (item.treasury_tracker_id or "")[:26])
        y -= 13

    pdf.save()
    buffer.seek(0)
    download_name = f"inventory_item_nfc_{date.today().isoformat()}.pdf"
    return send_file(buffer, as_attachment=True, download_name=download_name, mimetype="application/pdf")


@app.get("/portal/admin/inventory/treasury-report.csv")
@require_role("admin")
def portal_admin_inventory_treasury_report():
    items = Item.query.order_by(Item.name.asc(), Item.id.asc()).all()
    tag_map = get_item_primary_tag_map()
    rows = []
    for item in items:
        checked_out_qty = max((item.total_qty or 0) - (item.available_qty or 0), 0)
        rows.append(
            [
                item.item_code or "",
                item.treasury_tracker_id or "",
                item.id,
                item.name or "",
                get_item_primary_nfc(item, tag_map) or "",
                item.category or "",
                item.location or "",
                item.total_qty or 0,
                item.available_qty or 0,
                checked_out_qty,
                item.min_stock_threshold or 0,
                "yes" if item.active else "no",
            ]
        )
    content = csv_stream_from_rows(
        [
            "item_code",
            "treasury_tracker_id",
            "db_item_id",
            "item_name",
            "nfc_id",
            "category",
            "location",
            "total_qty",
            "available_qty",
            "checked_out_qty",
            "min_stock_threshold",
            "active",
        ],
        rows,
    )
    return Response(
        content,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=inventory_treasury_report.csv"},
    )


@app.post("/portal/admin/print/<int:request_id>/status")
@require_role("admin")
def portal_admin_print_status(request_id):
    row = db.session.get(PrintRequest, request_id)
    if not row:
        flash("Print request not found.", "error")
        return redirect_to_next("portal_admin_prints_page")
    status = (request.form.get("status") or "").strip().lower()
    if status not in PRINT_REQUEST_STATUSES:
        flash("Invalid status value.", "error")
        return redirect_to_next("portal_admin_prints_page")
    printer_type = normalize_portal_printer_type(request.form.get("printer_type")) or row.printer_type
    if printer_type not in PRINT_REQUEST_PRINTERS:
        printer_type = row.printer_type
    admin_notes = (request.form.get("admin_notes") or "").strip() or None
    user = current_auth_user()

    row.status = status
    row.printer_type = printer_type
    row.admin_notes = admin_notes
    row.reviewed_by_user_id = user.id
    row.reviewed_at = datetime.utcnow()
    add_audit_log("update_print_request", f"request_id={row.id} status={status}")
    db.session.commit()
    flash("Print request updated.", "success")
    return redirect_to_next("portal_admin_prints_page")


@app.post("/portal/admin/projects/save")
@require_role("admin")
def portal_admin_project_save():
    project_id = parse_positive_int(request.form.get("project_id"), default=0)
    title = (request.form.get("title") or "").strip()
    if not title:
        flash("Project title is required.", "error")
        return redirect_to_next("portal_admin_announcements_page")
    project = db.session.get(Project, project_id) if project_id else None
    if not project:
        project = Project(
            slug=slugify(title) or f"project-{uuid4().hex[:8]}",
            title=title[:200],
            summary="",
            description="",
            status="Active",
        )
        db.session.add(project)

    project.title = title[:200]
    project.slug = slugify(request.form.get("slug") or project.title) or project.slug
    project.summary = (request.form.get("summary") or "").strip()[:320] or "Project summary pending."
    project.description = (request.form.get("description") or "").strip() or "Project details pending."
    project.project_type = (request.form.get("project_type") or "").strip()[:80] or project.project_type or "General"
    project.status = (request.form.get("status") or "").strip()[:80] or "Active"
    timeline_text = (request.form.get("timeline") or "").strip()
    project.timeline = timeline_text or None
    gallery_text = (request.form.get("gallery_json") or "").strip()
    if gallery_text:
        project.gallery_json = json.dumps(parse_json_list(gallery_text))
    elif project.gallery_json is None:
        project.gallery_json = json.dumps([])
    project.lead_name = (request.form.get("lead_name") or "").strip()[:160] or None
    project.image_url = (request.form.get("image_url") or "").strip()[:500] or None
    project.external_link = (request.form.get("external_link") or "").strip()[:500] or None
    add_audit_log("save_project", f"project_id={project.id or 'new'} title={project.title}")
    db.session.commit()
    flash("Project saved.", "success")
    return redirect_to_next("portal_admin_announcements_page")


@app.post("/portal/admin/contact/<int:message_id>/status")
@require_role("admin")
def portal_admin_contact_status(message_id):
    row = db.session.get(ContactMessage, message_id)
    if not row:
        flash("Message not found.", "error")
        return redirect_to_next("portal_admin_announcements_page")
    status = (request.form.get("status") or "new").strip().lower()
    if status not in {"new", "in_progress", "resolved"}:
        status = "new"
    admin_reply = (request.form.get("admin_reply") or "").strip()
    row.status = status
    row.admin_reply = admin_reply[:10000] or None
    add_audit_log("update_contact_status", f"message_id={row.id} status={status}")
    db.session.commit()
    flash("Message status updated.", "success")
    return redirect_to_next("portal_admin_announcements_page")


@app.get("/app")
def app_frontend():
    if not legacy_ops_enabled():
        return redirect("/kiosk")
    return render_template("front/portal.html", **frontend_portal_context())


@app.get("/admin")
def admin_home():
    current = current_auth_user()
    if current and role_allows(current.role, "admin"):
        return redirect(url_for("portal_admin_dashboard"))
    return redirect(url_for("admin_login_page"))


@app.get("/dashboard")
def dashboard_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_router"))
    return render_ops_page(
        template_name="ops/dashboard.html",
        active_page="dashboard",
        page_title="Dashboard",
        page_subtitle="Overview of attendance, stock, and printer queue status.",
        transaction_limit=12,
    )


@app.get("/attendance")
def attendance_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_admin_attendance_page"))
    return render_ops_page(
        template_name="ops/attendance.html",
        active_page="attendance",
        page_title="Attendance",
        page_subtitle="Scan member NFC UIDs and track who is present today.",
        transaction_limit=10,
    )


@app.get("/inventory")
def inventory_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_inventory"))
    return render_ops_page(
        template_name="ops/inventory.html",
        active_page="inventory",
        page_title="Inventory",
        page_subtitle="Checkout, return, and monitor stock health across all items.",
        transaction_limit=12,
    )


@app.get("/prints")
def prints_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_prints"))
    return render_ops_page(
        template_name="ops/prints.html",
        active_page="prints",
        page_title="Printing",
        page_subtitle="Submit jobs and run separate H2S and P1S print queues.",
        transaction_limit=12,
    )


@app.get("/activity")
def activity_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_admin_inventory_page"))
    return render_ops_page(
        template_name="ops/activity.html",
        active_page="activity",
        page_title="Activity Feed",
        page_subtitle="Recent checkout and return history.",
        transaction_limit=80,
    )


@app.get("/settings")
def settings_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_admin_settings_page"))
    return render_ops_page(
        template_name="ops/settings.html",
        active_page="settings",
        page_title="Settings",
        page_subtitle="Customize theme and layout options for this device.",
        transaction_limit=10,
    )


@app.post("/session/member")
def set_active_member_route():
    member = resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    next_url = (request.form.get("next") or "").strip()
    if not member:
        flash("Could not sign in. Select or scan a valid member.", "error")
        return redirect(next_url if next_url.startswith("/") else url_for("scan_page"))

    session["active_member_id"] = member.id
    flash(f"Signed in as {member.name}.", "success")
    return redirect(next_url if next_url.startswith("/") else url_for("scan_page"))


@app.post("/session/member/clear")
def clear_active_member_route():
    next_url = (request.form.get("next") or "").strip()
    session.pop("active_member_id", None)
    flash("Signed out.", "info")
    return redirect(next_url if next_url.startswith("/") else url_for("scan_page"))


@app.get("/api/session/me")
def api_session_me():
    user = current_auth_user()
    if user and not session.get("active_member_id"):
        linked = current_user_member(user)
        if linked:
            session["active_member_id"] = linked.id
    member = get_active_member()
    if not member:
        return jsonify(
            {
                "ok": True,
                "payload": {
                    "member": None,
                    "is_admin": bool(user and role_allows(user.role, "admin")),
                    "user": serialize_user(user),
                },
            }
        )
    return jsonify(
        {
            "ok": True,
            "payload": {
                "member": serialize_member(member),
                "is_admin": is_admin_member(member),
                "user": serialize_user(user),
            },
        }
    )


@app.post("/api/session/member")
def api_set_active_member():
    ensure_inventory_schema_columns()
    member = resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
    if not member:
        return jsonify({"ok": False, "error": "Could not sign in. Select or scan a valid member."}), 404

    session["active_member_id"] = member.id
    user = current_auth_user()
    if user and not user.member_id:
        user.member_id = member.id
        db.session.commit()
    return jsonify(
        {
            "ok": True,
            "message": f"Signed in as {member.name}.",
            "payload": {
                "member": serialize_member(member),
                "is_admin": is_admin_member(member),
                "user": serialize_user(user),
            },
        }
    )


@app.post("/api/session/clear")
def api_clear_active_member():
    session.pop("active_member_id", None)
    return jsonify({"ok": True, "message": "Signed out."})


@app.get("/api/my-items")
def api_my_items():
    member, auth_error = require_active_member_json()
    if auth_error:
        return auth_error

    open_checkouts = (
        Transaction.query.filter_by(member_id=member.id, status="OUT")
        .order_by(Transaction.checkout_time.desc(), Transaction.id.desc())
        .all()
    )
    return jsonify(
        {
            "ok": True,
            "payload": {
                "member": serialize_member(member),
                "items": [serialize_open_checkout(tx) for tx in open_checkouts],
            },
        }
    )


@app.get("/scan")
def scan_page():
    if not legacy_ops_enabled():
        return redirect("/kiosk")
    context = dashboard_context(transaction_limit=25)
    context.update(
        {
            "active_page": "scan",
            "page_title": "NFC Scanner",
            "page_subtitle": "Scan item tags to check out or return equipment.",
            "active_member": get_active_member(),
            "active_member_is_admin": is_admin_member(get_active_member()),
        }
    )
    return render_template("ops/scan.html", **context)


@app.get("/my-items")
def my_items_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_checkouts"))
    active_member = get_active_member()
    open_checkouts = []
    if active_member:
        open_checkouts = (
            Transaction.query.filter_by(member_id=active_member.id, status="OUT")
            .order_by(Transaction.checkout_time.desc(), Transaction.id.desc())
            .all()
        )

    context = dashboard_context(transaction_limit=25)
    context.update(
        {
            "active_page": "my_items",
            "page_title": "My Checked Out Items",
            "page_subtitle": "Your currently checked-out inventory and checkout timestamps.",
            "active_member": active_member,
            "active_member_is_admin": is_admin_member(active_member),
            "open_checkouts": open_checkouts,
        }
    )
    return render_template("ops/my_items.html", **context)


@app.get("/admin/nfc")
def admin_nfc_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_admin_nfc_page"))
    active_member = get_active_member()
    if not is_admin_member(active_member):
        flash("Admin access required for NFC registration and transaction log.", "error")
        return redirect(url_for("scan_page"))

    item_id_raw = (request.args.get("item_id") or "").strip()
    member_id_raw = (request.args.get("member_id") or "").strip()
    date_from_raw = (request.args.get("date_from") or "").strip()
    date_to_raw = (request.args.get("date_to") or "").strip()

    query = Transaction.query.join(Item, Transaction.item_id == Item.id).join(Member, Transaction.member_id == Member.id)
    if item_id_raw.isdigit():
        query = query.filter(Transaction.item_id == int(item_id_raw))
    if member_id_raw.isdigit():
        query = query.filter(Transaction.member_id == int(member_id_raw))
    if date_from_raw:
        date_from = parse_due_date(date_from_raw)
        if date_from:
            query = query.filter(Transaction.timestamp >= datetime.combine(date_from, time.min))
    if date_to_raw:
        date_to = parse_due_date(date_to_raw)
        if date_to:
            query = query.filter(Transaction.timestamp <= datetime.combine(date_to, time.max))

    tx_rows = query.order_by(Transaction.timestamp.desc(), Transaction.id.desc()).limit(250).all()
    tag_rows = ItemTag.query.order_by(ItemTag.created_at.desc(), ItemTag.id.desc()).limit(250).all()

    context = dashboard_context(transaction_limit=25)
    context.update(
        {
            "active_page": "admin_nfc",
            "page_title": "NFC Admin",
            "page_subtitle": "Register tags, review transaction history, and correct inventory values.",
            "active_member": active_member,
            "active_member_is_admin": True,
            "tag_rows": tag_rows,
            "tx_rows": tx_rows,
            "filter_item_id": item_id_raw,
            "filter_member_id": member_id_raw,
            "filter_date_from": date_from_raw,
            "filter_date_to": date_to_raw,
        }
    )
    return render_template("ops/admin_nfc.html", **context)


@app.post("/admin/nfc/register")
def admin_register_item_tag():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_admin_nfc_page"))
    active_member = get_active_member()
    if not is_admin_member(active_member):
        flash("Admin access required.", "error")
        return redirect(url_for("scan_page"))

    item_id_raw = (request.form.get("item_id") or "").strip()
    tag_value = clean_tag_value(request.form.get("tag_value"))
    source = (request.form.get("source") or "manual").strip().lower() or "manual"

    if not item_id_raw.isdigit():
        flash("Select a valid item.", "error")
        return redirect(url_for("admin_nfc_page"))
    if not tag_value:
        flash("Tag value is required.", "error")
        return redirect(url_for("admin_nfc_page"))

    item = db.session.get(Item, int(item_id_raw))
    if not item:
        flash("Item not found.", "error")
        return redirect(url_for("admin_nfc_page"))

    existing = ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag_value.lower()).first()
    if existing and existing.item_id != item.id:
        flash(f"That tag is already assigned to {existing.item.name}.", "error")
        return redirect(url_for("admin_nfc_page"))

    if existing:
        existing.source = source
    else:
        db.session.add(ItemTag(item_id=item.id, tag_value=tag_value, source=source))

    if not clean_tag_value(item.nfc_tag):
        item.nfc_tag = tag_value

    db.session.commit()
    flash(f"Registered tag for {item.name}.", "success")
    return redirect(url_for("admin_nfc_page"))


@app.post("/admin/inventory/correct")
def admin_inventory_correct():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_admin_inventory_page"))
    active_member = get_active_member()
    if not is_admin_member(active_member):
        flash("Admin access required.", "error")
        return redirect(url_for("scan_page"))

    item_id_raw = (request.form.get("item_id") or "").strip()
    total_qty = parse_int(request.form.get("total_qty"), default=0)
    available_qty = parse_int(request.form.get("available_qty"), default=0)
    note = (request.form.get("note") or "").strip() or "Admin inventory correction"

    if not item_id_raw.isdigit():
        flash("Select a valid item for correction.", "error")
        return redirect(url_for("admin_nfc_page"))

    item = db.session.get(Item, int(item_id_raw))
    if not item:
        flash("Item not found.", "error")
        return redirect(url_for("admin_nfc_page"))

    total_qty = max(0, total_qty)
    available_qty = max(0, min(total_qty, available_qty))
    item.total_qty = total_qty
    item.available_qty = available_qty
    db.session.add(
        Transaction(
            member_id=active_member.id,
            item_id=item.id,
            action="return",
            qty=0,
            status="RETURNED",
            timestamp=datetime.utcnow(),
            return_time=datetime.utcnow(),
            return_condition="admin-correction",
            return_notes=note,
            notes=note,
        )
    )
    db.session.commit()
    flash(f"Inventory corrected for {item.name}.", "success")
    return redirect(url_for("admin_nfc_page"))


@app.get("/calendar")
def calendar_page():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_calendar"))
    ensure_meeting_schema_columns()
    active_member = get_active_member()
    upcoming_meetings = get_upcoming_meetings()
    meetings_by_room = {room: [] for room in MEETING_ROOMS}
    for meeting in upcoming_meetings:
        meetings_by_room.setdefault(meeting.room, []).append(meeting)
    outlook_calendar = outlook_calendar_embed_context()
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
            "active_member": active_member,
            "active_member_is_admin": is_admin_member(active_member),
            "meeting_rooms": MEETING_ROOMS,
            "meetings_by_room": meetings_by_room,
            "calendar_default_date": str(date.today()),
            "outlook_calendar_embed_url": outlook_calendar["url"],
            "outlook_calendar_open_url": outlook_calendar["open_url"],
            "outlook_calendar_placeholder": outlook_calendar["placeholder"],
            "calendar_automation_status": automation_status,
            "pending_cancellations": pending_cancellations,
        }
    )
    return render_template("ops/calendar.html", **context)


@app.post("/calendar/book")
def book_meeting():
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_calendar"))
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

    sync_configured = is_outlook_sync_configured()
    config_has_inputs = has_any_outlook_sync_inputs()
    calendar_error = None
    if sync_configured:
        calendar_id, event_id, calendar_error = create_outlook_calendar_event(meeting)
        if not calendar_error:
            meeting.outlook_calendar_id = calendar_id
            meeting.outlook_event_id = event_id

    db.session.add(meeting)
    db.session.commit()

    if calendar_error:
        flash(
            f"Booked {room} for {team_name} on {meeting_date.isoformat()} "
            f"({start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}). "
            f"Outlook sync failed: {calendar_error}",
            "info",
        )
    elif sync_configured:
        flash(
            f"Booked {room} for {team_name} on {meeting_date.isoformat()} "
            f"({start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}) and synced to Outlook Calendar.",
            "success",
        )
    elif config_has_inputs:
        flash(
            f"Booked {room} for {team_name} on {meeting_date.isoformat()} "
            f"({start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}). "
            "Outlook sync is not active yet. Run: python scripts/outlook_sync_doctor.py",
            "info",
        )
    else:
        flash(
            f"Booked {room} for {team_name} on {meeting_date.isoformat()} "
            f"({start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}).",
            "success",
        )
    return redirect_home("calendar")


@app.post("/calendar/meeting/<int:meeting_id>/cancel")
def request_meeting_cancel(meeting_id):
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_calendar"))
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
        "Cancellation request sent. An approval email was sent for confirmation.",
        "info",
    )
    return redirect_home("calendar")


@app.get("/calendar/cancel/confirm/<token>")
def confirm_meeting_cancel(token):
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_calendar"))
    ensure_meeting_schema_columns()
    meeting = Meeting.query.filter_by(cancel_request_token=token).first()
    if not meeting:
        return (
            "<h3>Cancellation link is invalid or already used.</h3>"
            "<p>You can close this tab.</p>",
            404,
        )

    calendar_error = delete_outlook_calendar_event(meeting)
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
        f"<p>{team_name} in {room} on {meeting_date} was removed from Outlook Calendar.</p>"
        "<p>You can close this tab.</p>"
    )


@app.get("/calendar/cancel/reject/<token>")
def reject_meeting_cancel(token):
    if not legacy_ops_enabled():
        return redirect(url_for("portal_member_calendar"))
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
        "<p>The meeting remains on the schedule and in Outlook Calendar.</p>"
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
    ensure_inventory_schema_columns()
    auth_user = current_auth_user()
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
        updated = (
            Item.query.filter(Item.id == item.id, Item.available_qty >= qty)
            .update({Item.available_qty: Item.available_qty - qty}, synchronize_session=False)
        )
        if updated != 1:
            db.session.rollback()
            db.session.refresh(item)
            return api_error(f"Not enough stock. {item.name} has {item.available_qty} available.")
        tx_status = "OUT"
        checkout_time = datetime.utcnow()
        return_time = None
        checkout_notes = notes
        return_condition = None
        return_notes = None
    else:
        item.available_qty = min(item.total_qty, item.available_qty + qty)
        tx_status = "RETURNED"
        checkout_time = None
        return_time = datetime.utcnow()
        checkout_notes = None
        return_condition = "manual-return"
        return_notes = notes

    tx = Transaction(
        member_id=member.id,
        user_id=auth_user.id if auth_user else None,
        item_id=item.id,
        action=action,
        qty=qty,
        status=tx_status,
        checkout_time=checkout_time,
        return_time=return_time,
        due_date=due if action == "checkout" else None,
        notes=notes,
        checkout_notes=checkout_notes,
        return_condition=return_condition,
        return_notes=return_notes,
    )
    db.session.add(tx)
    db.session.commit()

    return api_success(message=f"{action.title()} saved: {qty} x {item.name} for {member.name}.")


@app.get("/api/items/by-tag")
def api_item_by_tag():
    ensure_inventory_schema_columns()
    member, auth_error = require_active_member_json()
    if auth_error:
        return auth_error

    tag_value = clean_tag_value(request.args.get("tag"))
    if not tag_value:
        return jsonify({"ok": False, "error": "Tag value is required."}), 400

    item, resolved_via = find_item_by_tag(tag_value)
    if not item:
        return jsonify({"ok": False, "error": "Tag not registered to an inventory item."}), 404

    open_for_user = get_open_checkout(member.id, item.id)
    checked_out_by_others = (
        Transaction.query.filter(
            Transaction.item_id == item.id,
            Transaction.status == "OUT",
            Transaction.member_id != member.id,
        ).count()
    )

    payload = {
        "item": {
            "id": item.id,
            "name": item.name,
            "description": item.description,
            "category": item.category,
            "location": item.location,
            "available_qty": item.available_qty,
            "total_qty": item.total_qty,
            "resolved_via": resolved_via,
        },
        "active_member": {
            "id": member.id,
            "name": member.name,
            "email": member.email,
        },
        "user_has_open_checkout": bool(open_for_user),
        "user_open_checkout": (
            {
                "transaction_id": open_for_user.id,
                "qty": open_for_user.qty,
                "checkout_time": iso_or_none(open_for_user.checkout_time or open_for_user.timestamp),
                "checkout_notes": open_for_user.checkout_notes or open_for_user.notes,
            }
            if open_for_user
            else None
        ),
        "checked_out_by_others_count": checked_out_by_others,
    }
    return jsonify({"ok": True, "payload": payload})


@app.post("/api/checkout")
def api_checkout_item():
    ensure_inventory_schema_columns()
    member, auth_error = require_active_member_json()
    if auth_error:
        return auth_error
    auth_user = current_auth_user()

    item_id_raw = value_from_request("item_id")
    qty = parse_int(value_from_request("qty"), default=1)
    checkout_notes = (str(value_from_request("notes", "")).strip() or None)
    due = parse_due_date(value_from_request("due_date"))

    try:
        item_id = int(item_id_raw)
    except Exception:
        return jsonify({"ok": False, "error": "item_id is required."}), 400

    item = db.session.get(Item, item_id)
    if not item:
        return jsonify({"ok": False, "error": "Item not found."}), 404

    existing_open = get_open_checkout(member.id, item.id)
    if existing_open:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "You already have this item checked out. Return it before checking out again.",
                }
            ),
            409,
        )

    updated = (
        Item.query.filter(Item.id == item.id, Item.available_qty >= qty)
        .update({Item.available_qty: Item.available_qty - qty}, synchronize_session=False)
    )
    if updated != 1:
        db.session.rollback()
        db.session.refresh(item)
        return (
            jsonify(
                {
                    "ok": False,
                    "error": f"Not enough stock. {item.name} has {item.available_qty} available.",
                }
            ),
            409,
        )

    now = datetime.utcnow()
    tx = Transaction(
        member_id=member.id,
        user_id=auth_user.id if auth_user else None,
        item_id=item.id,
        action="checkout",
        qty=qty,
        status="OUT",
        timestamp=now,
        checkout_time=now,
        checkout_notes=checkout_notes,
        notes=checkout_notes,
        due_date=due,
    )
    db.session.add(tx)
    db.session.commit()
    db.session.refresh(item)

    return jsonify(
        {
            "ok": True,
            "message": f"Checked out {qty} x {item.name}.",
            "payload": {
                "item_id": item.id,
                "available_qty": item.available_qty,
                "transaction_id": tx.id,
                "member_id": member.id,
            },
        }
    )


@app.post("/api/return")
def api_return_item():
    ensure_inventory_schema_columns()
    member, auth_error = require_active_member_json()
    if auth_error:
        return auth_error
    auth_user = current_auth_user()

    item_id_raw = value_from_request("item_id")
    qty = parse_int(value_from_request("qty"), default=1)
    return_condition = (str(value_from_request("condition", "")).strip() or None)
    return_notes = (str(value_from_request("notes", "")).strip() or None)

    try:
        item_id = int(item_id_raw)
    except Exception:
        return jsonify({"ok": False, "error": "item_id is required."}), 400

    if not return_condition:
        return jsonify({"ok": False, "error": "Return condition is required."}), 400

    item = db.session.get(Item, item_id)
    if not item:
        return jsonify({"ok": False, "error": "Item not found."}), 404

    open_tx = get_open_checkout(member.id, item.id)
    if not open_tx:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "No open checkout found for this item under your account.",
                }
            ),
            409,
        )

    if qty != open_tx.qty:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": (
                        "Return quantity must match the checked-out quantity for this item. "
                        f"Checked out quantity: {open_tx.qty}."
                    ),
                }
            ),
            409,
        )

    return_photo_path = None
    return_photo = request.files.get("photo")
    if return_photo and return_photo.filename:
        if not allowed_return_photo(return_photo.filename):
            return jsonify({"ok": False, "error": "Photo must be .jpg, .jpeg, .png, or .webp."}), 400
        safe_name = secure_filename(return_photo.filename)
        stored_name = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{safe_name}"
        photo_path = RETURN_PHOTO_DIR / stored_name
        return_photo.save(photo_path)
        return_photo_path = str(photo_path)

    item.available_qty = min(item.total_qty, item.available_qty + qty)
    now = datetime.utcnow()
    open_tx.status = "RETURNED"
    if auth_user and not open_tx.user_id:
        open_tx.user_id = auth_user.id
    open_tx.return_time = now
    open_tx.return_condition = return_condition
    open_tx.return_notes = return_notes
    open_tx.return_photo_path = return_photo_path
    db.session.commit()

    return jsonify(
        {
            "ok": True,
            "message": f"Returned {qty} x {item.name}.",
            "payload": {
                "item_id": item.id,
                "available_qty": item.available_qty,
                "transaction_id": open_tx.id,
                "member_id": member.id,
            },
        }
    )


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
    ensure_inventory_schema_columns()
    auth_user = current_auth_user()
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
        updated = (
            Item.query.filter(Item.id == item.id, Item.available_qty >= qty)
            .update({Item.available_qty: Item.available_qty - qty}, synchronize_session=False)
        )
        if updated != 1:
            db.session.rollback()
            db.session.refresh(item)
            flash(f"Not enough stock. {item.name} has {item.available_qty} available.", "error")
            return redirect_home("inventory")
        tx_status = "OUT"
        checkout_time = datetime.utcnow()
        return_time = None
        checkout_notes = notes
        return_condition = None
        return_notes = None
    else:
        item.available_qty = min(item.total_qty, item.available_qty + qty)
        tx_status = "RETURNED"
        checkout_time = None
        return_time = datetime.utcnow()
        checkout_notes = None
        return_condition = "manual-return"
        return_notes = notes

    tx = Transaction(
        member_id=member.id,
        user_id=auth_user.id if auth_user else None,
        item_id=item.id,
        action=action,
        qty=qty,
        status=tx_status,
        checkout_time=checkout_time,
        return_time=return_time,
        due_date=due if action == "checkout" else None,
        notes=notes,
        checkout_notes=checkout_notes,
        return_condition=return_condition,
        return_notes=return_notes,
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
        item_tag_conflict = ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag.lower()).first()
        if item_tag_conflict:
            flash("That UID is already assigned to an inventory item tag.", "error")
            return redirect(url_for("pair_member"))
        item_conflict = Item.query.filter(func.lower(Item.nfc_tag) == tag.lower()).first()
        if item_conflict:
            flash("That UID is already assigned to an inventory item.", "error")
            return redirect(url_for("pair_member"))
        user_tag_conflict = NFCTag.query.filter(func.lower(NFCTag.tag_uid) == tag.lower(), NFCTag.active.is_(True)).first()
        if user_tag_conflict:
            flash("That UID is already assigned to a user account tag.", "error")
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
    ensure_inventory_schema_columns()
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
        member_tag_conflict = NFCTag.query.filter(func.lower(NFCTag.tag_uid) == tag.lower()).first()
        if member_tag_conflict:
            flash("That UID is already assigned to a member tag.", "error")
            return redirect(url_for("pair_item"))

        item = db.session.get(Item, item_id)
        if not item:
            flash("Item not found.", "error")
            return redirect(url_for("pair_item"))

        item.nfc_tag = tag
        mapped = ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag.lower()).first()
        if mapped and mapped.item_id != item.id:
            flash("That UID is already assigned in item tag map.", "error")
            return redirect(url_for("pair_item"))
        if not mapped:
            db.session.add(ItemTag(item_id=item.id, tag_value=tag, source="pair_item_page"))
        db.session.commit()
        flash(f"Saved UID for {item.name}.", "success")
        return redirect(url_for("pair_item"))

    items = Item.query.order_by(Item.name.asc()).all()
    paired = Item.query.filter(Item.nfc_tag.isnot(None)).order_by(Item.name.asc()).all()
    return render_template("pair_item.html", items=items, paired=paired)


@app.get("/export")
def export_excel():
    try:
        import pandas as pd
    except Exception:
        flash("Excel export is unavailable because pandas is not loading correctly.", "error")
        return redirect(url_for("dashboard_page"))

    ensure_inventory_schema_columns()
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
                "description": i.description,
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
                "status": t.status,
                "qty": t.qty,
                "checkout_time": t.checkout_time,
                "return_time": t.return_time,
                "due_date": t.due_date,
                "notes": t.notes,
                "checkout_notes": t.checkout_notes,
                "return_condition": t.return_condition,
                "return_notes": t.return_notes,
                "return_photo_path": t.return_photo_path,
            }
            for t in transactions
        ]
    )
    item_tags = ItemTag.query.order_by(ItemTag.id.asc()).all()
    item_tags_df = pd.DataFrame(
        [
            {
                "id": tag.id,
                "item_id": tag.item_id,
                "item_name": tag.item.name if tag.item else "",
                "tag_value": tag.tag_value,
                "source": tag.source,
                "created_at": tag.created_at,
            }
            for tag in item_tags
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
                "outlook_event_id": m.outlook_event_id,
                "outlook_calendar_id": m.outlook_calendar_id,
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
        item_tags_df.to_excel(writer, index=False, sheet_name="ItemTags")
        attendance_df.to_excel(writer, index=False, sheet_name="Attendance")
        jobs_df.to_excel(writer, index=False, sheet_name="PrintJobs")
        meetings_df.to_excel(writer, index=False, sheet_name="Meetings")

    return send_file(export_path, as_attachment=True, download_name="inventory_export.xlsx")


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        ensure_inventory_schema_columns()
        ensure_meeting_schema_columns()
        ensure_portal_schema()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
