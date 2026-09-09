"""Typed application settings.

Every ``ASME_*`` / integration environment variable is resolved exactly once here,
at app creation, and validated. Nothing else in the codebase reads ``os.environ``
for configuration - modules take what they need from ``current_app.config["SETTINGS"]``
(see :func:`asme.config.settings`).
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from uuid import uuid4

from flask import current_app

TRUTHY = {"1", "true", "yes", "on", "y"}
FALSY = {"0", "false", "no", "off", "n"}


def _str(name: str, default: str = "") -> str:
    return (os.environ.get(name) or default).strip()


def _bool(name: str, default: bool = False) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return bool(default)
    if raw in TRUTHY:
        return True
    if raw in FALSY:
        return False
    return bool(default)


def _int(name: str, default: int, minimum: int | None = None) -> int:
    raw = (os.environ.get(name) or "").strip()
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError:
        return default
    if minimum is not None and value < minimum:
        return default
    return value


def load_instance_env_file(path: Path) -> None:
    """Load ``KEY=value`` lines from an instance-local env file into ``os.environ``.

    Kept for backward compatibility with ``instance/print_commands.env`` which
    operators used to configure printer commands without touching the host env.
    """
    if not path.exists():
        return
    try:
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key:
                os.environ[key] = value
    except Exception:
        # A malformed local env file must never stop the app from booting.
        pass


@dataclass(frozen=True)
class Settings:
    # -- runtime -------------------------------------------------------------
    env: str = "production"
    secret_key: str = ""
    database_url: str = "sqlite:///inventory.db"
    port: int = 5000
    auto_migrate: bool = False
    templates_auto_reload: bool = False

    # -- sessions / auth -----------------------------------------------------
    session_cookie_secure: bool = True
    session_cookie_samesite: str = "Lax"
    member_session_idle_minutes: int = 240
    admin_session_idle_minutes: int = 30
    app_boot_token: str = ""
    login_rate_window_seconds: int = 900
    login_rate_max_attempts: int = 8
    bulk_password_hash_method: str = "pbkdf2:sha256:120000"
    password_reset_hours: int = 2

    # -- bootstrap accounts --------------------------------------------------
    default_admin_email: str = "admin@uiowa.edu"
    default_admin_password: str = "ChangeMe123!"
    default_user_password: str = "ChangeMe123!"
    shared_admin_email: str = ""
    shared_admin_password: str = ""
    admin_emails: frozenset[str] = field(default_factory=frozenset)

    # -- feature flags -------------------------------------------------------
    legacy_ops_enabled: bool = False
    onboarding_enforce: bool = False
    outbox_worker_enabled: bool = True
    outbox_poll_seconds: int = 15

    # -- uploads / limits ----------------------------------------------------
    print_max_upload_bytes: int = 50 * 1024 * 1024
    checkin_soon_window_minutes: int = 20

    # -- calendar ------------------------------------------------------------
    calendar_provider: str = "google"
    google_calendar_embed_url: str = ""
    outlook_calendar_embed_url: str = ""
    google_service_account_json: str = ""
    google_calendar_id_robotics: str = ""
    google_calendar_id_fluids: str = ""
    google_calendar_timezone: str = "America/Chicago"
    calendar_scheduling_days: int = 14
    calendar_work_hours_start: str = "08:00"
    calendar_work_hours_end: str = "22:00"

    outlook_tenant_id: str = ""
    outlook_client_id: str = ""
    outlook_client_secret: str = ""
    outlook_calendar_user: str = ""
    outlook_calendar_id: str = ""
    outlook_calendar_robotics_id: str = ""
    outlook_calendar_fluids_id: str = ""
    outlook_calendar_tz: str = "Central Standard Time"

    # -- mail ----------------------------------------------------------------
    smtp_host: str = "smtp.office365.com"
    smtp_port: int = 587
    smtp_user: str = ""
    smtp_pass: str = ""
    cancel_notify_to: str = ""

    # -- printers ------------------------------------------------------------
    h2s_print_cmd: str = ""
    p1s_print_cmd: str = ""

    # -- assistant -----------------------------------------------------------
    anthropic_api_key: str = ""
    anthropic_model: str = "claude-sonnet-5"

    # -- ASME Ops ------------------------------------------------------------
    ops_enabled: bool = True
    ops_web_dist: str = ""
    file_max_mb: int = 25
    file_url_ttl_seconds: int = 900

    @property
    def is_production(self) -> bool:
        return self.env == "production"

    @property
    def is_testing(self) -> bool:
        return self.env == "testing"

    @classmethod
    def from_env(cls, **overrides) -> "Settings":
        env = _str("ASME_ENV", "development" if _bool("FLASK_DEBUG") else "production").lower()
        if env not in {"development", "production", "testing"}:
            env = "production"
        samesite = _str("ASME_SESSION_COOKIE_SAMESITE", "Lax").lower()
        if samesite not in {"lax", "strict", "none"}:
            samesite = "lax"
        admin_emails = frozenset(
            email.strip().lower() for email in _str("ASME_ADMIN_EMAILS").split(",") if email.strip()
        )
        values = dict(
            env=env,
            secret_key=_str("ASME_SECRET_KEY", "asme-dev-secret"),
            database_url=_str("ASME_DATABASE_URL", "sqlite:///inventory.db"),
            port=_int("PORT", 5000, minimum=1),
            auto_migrate=_bool("ASME_AUTO_MIGRATE", default=(env != "production")),
            templates_auto_reload=_bool("ASME_TEMPLATES_AUTO_RELOAD", default=(env == "development")),
            session_cookie_secure=_bool("ASME_SESSION_COOKIE_SECURE", default=bool(os.environ.get("RENDER"))),
            session_cookie_samesite="None" if samesite == "none" else samesite.capitalize(),
            member_session_idle_minutes=_int("ASME_SESSION_IDLE_MINUTES", 240, minimum=0),
            admin_session_idle_minutes=_int("ASME_ADMIN_SESSION_IDLE_MINUTES", 30, minimum=0),
            app_boot_token=_str("ASME_APP_BOOT_TOKEN") or uuid4().hex,
            login_rate_window_seconds=_int("ASME_LOGIN_RATE_WINDOW_SECONDS", 900, minimum=1),
            login_rate_max_attempts=_int("ASME_LOGIN_RATE_MAX_ATTEMPTS", 8, minimum=1),
            bulk_password_hash_method=_str("ASME_BULK_PASSWORD_HASH_METHOD", "pbkdf2:sha256:120000"),
            default_admin_email=_str("ASME_DEFAULT_ADMIN_EMAIL", "admin@uiowa.edu").lower(),
            default_admin_password=_str("ASME_DEFAULT_ADMIN_PASSWORD", "ChangeMe123!"),
            default_user_password=_str("ASME_DEFAULT_USER_PASSWORD", "ChangeMe123!"),
            shared_admin_email=(
                _str("ASME_DEFAULT_ADMIN_EMAIL") or _str("ASME_SHARED_ADMIN_EMAIL") or _str("ASME_ADMIN_EMAIL")
            ).lower(),
            shared_admin_password=(
                _str("ASME_DEFAULT_ADMIN_PASSWORD")
                or _str("ASME_SHARED_ADMIN_PASSWORD")
                or _str("ASME_ADMIN_PASSWORD")
            ),
            admin_emails=admin_emails,
            legacy_ops_enabled=_bool("ASME_ENABLE_LEGACY_OPS", default=False),
            onboarding_enforce=_bool("ASME_ONBOARDING_ENFORCE", default=False),
            outbox_worker_enabled=_bool("ASME_OUTBOX_WORKER", default=True),
            outbox_poll_seconds=_int("ASME_OUTBOX_POLL_SECONDS", 15, minimum=1),
            print_max_upload_bytes=_int("ASME_PRINT_MAX_UPLOAD_MB", 50, minimum=1) * 1024 * 1024,
            checkin_soon_window_minutes=_int("ASME_CHECKIN_SOON_WINDOW_MINUTES", 20, minimum=0),
            calendar_provider=(_str("ASME_CALENDAR_PROVIDER", "google").lower() or "google"),
            google_calendar_embed_url=_str("ASME_GOOGLE_CALENDAR_EMBED_URL"),
            outlook_calendar_embed_url=_str("ASME_OUTLOOK_CALENDAR_EMBED_URL"),
            google_service_account_json=_str("GOOGLE_SERVICE_ACCOUNT_JSON"),
            google_calendar_id_robotics=_str("GOOGLE_CALENDAR_ID_ROBOTICS"),
            google_calendar_id_fluids=_str("GOOGLE_CALENDAR_ID_FLUIDS"),
            google_calendar_timezone=_str("GOOGLE_CALENDAR_TIMEZONE", "America/Chicago"),
            calendar_scheduling_days=_int("CALENDAR_SCHEDULING_DAYS", 14, minimum=1),
            calendar_work_hours_start=_str("CALENDAR_WORK_HOURS_START", "08:00"),
            calendar_work_hours_end=_str("CALENDAR_WORK_HOURS_END", "22:00"),
            outlook_tenant_id=_str("ASME_OUTLOOK_TENANT_ID"),
            outlook_client_id=_str("ASME_OUTLOOK_CLIENT_ID"),
            outlook_client_secret=_str("ASME_OUTLOOK_CLIENT_SECRET"),
            outlook_calendar_user=_str("ASME_OUTLOOK_CALENDAR_USER"),
            outlook_calendar_id=_str("ASME_OUTLOOK_CALENDAR_ID"),
            outlook_calendar_robotics_id=_str("ASME_OUTLOOK_CALENDAR_ROBOTICS_ID"),
            outlook_calendar_fluids_id=_str("ASME_OUTLOOK_CALENDAR_FLUIDS_ID"),
            outlook_calendar_tz=_str("ASME_OUTLOOK_CALENDAR_TZ", "Central Standard Time"),
            smtp_host=_str("ASME_SMTP_HOST", "smtp.office365.com"),
            smtp_port=_int("ASME_SMTP_PORT", 587, minimum=1),
            smtp_user=_str("ASME_SMTP_USER"),
            smtp_pass=_str("ASME_SMTP_PASS"),
            cancel_notify_to=_str("ASME_CANCEL_NOTIFY_TO"),
            h2s_print_cmd=_str("ASME_H2S_PRINT_CMD"),
            p1s_print_cmd=_str("ASME_P1S_PRINT_CMD"),
            anthropic_api_key=_str("ANTHROPIC_API_KEY"),
            anthropic_model=_str("ASME_ASSISTANT_MODEL", "claude-sonnet-5"),
            ops_enabled=_bool("ASME_OPS_ENABLED", default=True),
            ops_web_dist=_str("ASME_OPS_WEB_DIST"),
            file_max_mb=_int("ASME_FILE_MAX_MB", 25, minimum=1),
            file_url_ttl_seconds=_int("ASME_FILE_URL_TTL_SECONDS", 900, minimum=10),
        )
        if values["calendar_provider"] not in {"google", "outlook"}:
            values["calendar_provider"] = "google"
        values.update(overrides)
        return cls(**values)

    def validate(self) -> list[str]:
        """Problems that make the app unsafe to run. In production these refuse startup."""
        problems: list[str] = []
        if self.is_production and self.secret_key in {"", "asme-dev-secret", "change-this-secret"}:
            problems.append("ASME_SECRET_KEY must be set to a real secret in production.")
        if self.session_cookie_samesite == "None" and not self.session_cookie_secure:
            problems.append("SameSite=None cookies require ASME_SESSION_COOKIE_SECURE=1.")
        outlook_inputs = (
            self.outlook_tenant_id,
            self.outlook_client_id,
            self.outlook_client_secret,
            self.outlook_calendar_user,
        )
        if any(outlook_inputs) and not all(outlook_inputs):
            problems.append(
                "Outlook sync is half-configured: set ASME_OUTLOOK_TENANT_ID, ASME_OUTLOOK_CLIENT_ID, "
                "ASME_OUTLOOK_CLIENT_SECRET and ASME_OUTLOOK_CALENDAR_USER together."
            )
        if bool(self.smtp_user) != bool(self.smtp_pass):
            problems.append("ASME_SMTP_USER and ASME_SMTP_PASS must be set together.")
        return problems

    def warnings(self) -> list[str]:
        """Things worth fixing that never block startup."""
        notes: list[str] = []
        if self.is_production and self.default_admin_password == "ChangeMe123!":
            notes.append("ASME_DEFAULT_ADMIN_PASSWORD is still the default; change it and rotate the admin login.")
        if self.is_production and not self.session_cookie_secure:
            notes.append("ASME_SESSION_COOKIE_SECURE is off; set it to 1 when serving over HTTPS.")
        if self.calendar_provider == "google" and not self.google_service_account_json and not self.google_calendar_embed_url:
            notes.append("No Google calendar credentials or embed URL; room scheduling is disabled until configured.")
        return notes


def settings() -> Settings:
    """The active :class:`Settings` for the current app context."""
    return current_app.config["SETTINGS"]
