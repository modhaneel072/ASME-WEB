"""Small, dependency-free parsers for form/query input."""

from __future__ import annotations

import json
from datetime import date, datetime, time, timedelta


def parse_int(value, default=1):
    """Positive-int parser used by the legacy kiosk forms (``0`` and negatives -> default)."""
    try:
        parsed = int(value)
        return parsed if parsed > 0 else default
    except Exception:
        return default


def parse_positive_int(raw_value, default=1):
    text_value = str(raw_value or "").strip().replace(",", "")
    if not text_value:
        return default
    try:
        parsed = int(text_value)
    except Exception:
        try:
            parsed = int(float(text_value))
        except Exception:
            return default
    return parsed if parsed > 0 else default


def parse_non_negative_int(raw_value, default=0):
    text_value = str(raw_value or "").strip().replace(",", "")
    if not text_value:
        return default
    try:
        parsed = int(text_value)
    except Exception:
        try:
            parsed = int(float(text_value))
        except Exception:
            return default
    return parsed if parsed >= 0 else default


def parse_float(raw_value, default=0.0):
    text_value = str(raw_value or "").strip().replace(",", "")
    if not text_value:
        return default
    try:
        return float(text_value)
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
        hour, minute = [int(part) for part in raw.split(":")[:2]]
        return time(hour=hour, minute=minute)
    except Exception:
        return None


def parse_datetime_local(raw_value):
    value = (raw_value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            continue
    return None


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
    return [line.strip() for line in text_value.splitlines() if line.strip()]


def default_due_date(days=7):
    return date.today() + timedelta(days=days)


def iso_or_none(value):
    if value is None:
        return None
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def academic_year_key(today: date | None = None) -> str:
    """``2026-27`` style key; the year rolls over on 1 August."""
    today = today or date.today()
    if today.month >= 8:
        start = today.year
    else:
        start = today.year - 1
    return f"{start}-{(start + 1) % 100:02d}"


def semester_start(today: date | None = None) -> date:
    """First day of the current semester window used for count rules."""
    today = today or date.today()
    if today.month >= 8:
        return date(today.year, 8, 1)
    if today.month >= 1 and today.month < 6:
        return date(today.year, 1, 1)
    return date(today.year, 6, 1)
