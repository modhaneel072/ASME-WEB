"""Google Calendar adapter (service account, freebusy + events.insert)."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlencode, urlparse
from zoneinfo import ZoneInfo

from asme.config import Settings
from asme.constants import MEETING_ROOMS
from asme.integrations.calendar.base import CalendarError, CalendarProvider, CreatedEvent, ProviderStatus

try:  # pragma: no cover - optional dependency in local dev
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except Exception:  # pragma: no cover
    service_account = None
    build = None


def normalize_google_calendar_id(raw_value):
    value = (raw_value or "").strip()
    if not value:
        return ""
    decoded = value
    for _ in range(2):
        maybe_decoded = unquote(decoded)
        if maybe_decoded == decoded:
            break
        decoded = maybe_decoded
    if decoded.lower().startswith("src="):
        return unquote(decoded.split("=", 1)[1]).strip()
    if decoded.lower().startswith(("http://", "https://")):
        try:
            parsed = urlparse(decoded)
            src_values = parse_qs(parsed.query).get("src") or []
            if src_values:
                return unquote((src_values[0] or "").strip()).strip()
        except Exception:
            pass
    return decoded


class GoogleCalendarProvider(CalendarProvider):
    name = "google"

    def __init__(self, cfg: Settings):
        self.cfg = cfg
        self._service = None
        self._service_source = ""

    # -- config ---------------------------------------------------------------

    def room_ids(self) -> dict[str, str]:
        return {
            "Robotics Room": normalize_google_calendar_id(self.cfg.google_calendar_id_robotics),
            "Fluids Lab": normalize_google_calendar_id(self.cfg.google_calendar_id_fluids),
        }

    def tzinfo(self):
        try:
            return ZoneInfo(self.cfg.google_calendar_timezone)
        except Exception:
            return ZoneInfo("UTC")

    def status(self) -> ProviderStatus:
        errors = []
        room_ids = self.room_ids()
        if self.cfg.calendar_provider != "google":
            errors.append("ASME_CALENDAR_PROVIDER is not set to google.")
        if not room_ids.get("Robotics Room"):
            errors.append("GOOGLE_CALENDAR_ID_ROBOTICS is not configured.")
        if not room_ids.get("Fluids Lab"):
            errors.append("GOOGLE_CALENDAR_ID_FLUIDS is not configured.")
        if not self.cfg.google_service_account_json:
            errors.append("GOOGLE_SERVICE_ACCOUNT_JSON is not configured.")
        if service_account is None or build is None:
            errors.append("google-api-python-client is not installed.")
        return ProviderStatus(provider=self.name, enabled=not errors, errors=errors, room_ids=room_ids)

    def _service_account_info(self):
        raw = self.cfg.google_service_account_json
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

    def service(self):
        if service_account is None or build is None:
            return None
        source_key = self.cfg.google_service_account_json
        if not source_key:
            return None
        if self._service and self._service_source == source_key:
            return self._service
        info = self._service_account_info()
        if not info:
            return None
        try:
            creds = service_account.Credentials.from_service_account_info(
                info,
                scopes=["https://www.googleapis.com/auth/calendar"],
            )
            self._service = build("calendar", "v3", credentials=creds, cache_discovery=False)
            self._service_source = source_key
        except Exception:
            return None
        return self._service

    # -- operations -----------------------------------------------------------

    def free_busy(self, time_min: datetime, time_max: datetime):
        status = self.status()
        service = self.service()
        if not status.enabled or not service:
            raise CalendarError("Google Calendar integration is not configured.")
        room_ids = status.room_ids
        body = {
            "timeMin": time_min.isoformat(),
            "timeMax": time_max.isoformat(),
            "timeZone": self.cfg.google_calendar_timezone,
            "items": [{"id": room_ids[room]} for room in MEETING_ROOMS if room_ids.get(room)],
        }
        try:
            response = service.freebusy().query(body=body).execute()
        except Exception as exc:
            raise CalendarError(f"Google freebusy request failed: {exc}") from exc

        tzinfo = self.tzinfo()
        busy_by_room = {room: [] for room in MEETING_ROOMS}
        for room in MEETING_ROOMS:
            calendar_id = room_ids.get(room)
            payload = ((response.get("calendars") or {}).get(calendar_id) or {}) if calendar_id else {}
            for entry in payload.get("busy") or []:
                try:
                    start_dt = datetime.fromisoformat((entry.get("start") or "").replace("Z", "+00:00")).astimezone(tzinfo)
                    end_dt = datetime.fromisoformat((entry.get("end") or "").replace("Z", "+00:00")).astimezone(tzinfo)
                except Exception:
                    continue
                busy_by_room[room].append((start_dt, end_dt))
        return busy_by_room

    def create_event(self, room, summary, description, start_local, end_local) -> CreatedEvent:
        status = self.status()
        service = self.service()
        if not status.enabled or not service:
            raise CalendarError("Google Calendar integration is not configured.")
        calendar_id = status.room_ids.get(room)
        if not calendar_id:
            raise CalendarError(f"No Google calendar ID configured for {room}.")
        body = {
            "summary": summary[:220],
            "description": (description or "")[:4000],
            "location": room,
            "start": {"dateTime": start_local.isoformat(), "timeZone": self.cfg.google_calendar_timezone},
            "end": {"dateTime": end_local.isoformat(), "timeZone": self.cfg.google_calendar_timezone},
        }
        try:
            payload = service.events().insert(calendarId=calendar_id, body=body).execute()
        except Exception as exc:
            raise CalendarError(f"Google event create failed: {exc}") from exc
        return CreatedEvent(
            external_id=payload.get("id") or None,
            calendar_id=calendar_id,
            link=payload.get("htmlLink") or None,
        )

    def delete_event(self, calendar_id, external_id):
        service = self.service()
        if not service:
            raise CalendarError("Google Calendar integration is not configured.")
        if not calendar_id or not external_id:
            return
        try:
            service.events().delete(calendarId=calendar_id, eventId=external_id).execute()
        except Exception as exc:
            raise CalendarError(f"Google event delete failed: {exc}") from exc

    def embed_url(self) -> str:
        if self.cfg.google_calendar_embed_url:
            return self.cfg.google_calendar_embed_url
        ids = [calendar_id for calendar_id in self.room_ids().values() if calendar_id]
        if not ids:
            return ""
        query = [("ctz", self.cfg.google_calendar_timezone), ("mode", "WEEK"), ("showTitle", "0"), ("showPrint", "0")]
        for calendar_id in ids:
            query.append(("src", calendar_id))
        return "https://calendar.google.com/calendar/embed?" + urlencode(query)
