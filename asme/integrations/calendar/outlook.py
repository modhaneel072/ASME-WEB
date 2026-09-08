"""Microsoft Graph (Outlook) calendar adapter - app-only client credentials."""

from __future__ import annotations

import time
from datetime import datetime
from urllib.parse import quote, urlparse

from asme.config import Settings
from asme.constants import KNOWN_NON_MS_MAIL_DOMAINS, MEETING_ROOMS, POSSIBLY_CONSUMER_MS_DOMAINS
from asme.integrations.calendar.base import CalendarError, CalendarProvider, CreatedEvent, ProviderStatus
from asme.integrations.http import http_json_request

GRAPH = "https://graph.microsoft.com/v1.0"


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
            owner_id, publish_id, cid = path_parts[2], path_parts[3], path_parts[4]
            scheme = parsed.scheme or "https"
            return f"{scheme}://{host}/calendar/0/published/{owner_id}/{publish_id}/{cid}/calendar.html/"
    except Exception:
        pass
    return raw_url


def validate_outlook_sync_config(cfg: Settings):
    errors, warnings = [], []
    mailbox_user = cfg.outlook_calendar_user
    if not mailbox_user:
        errors.append("ASME_OUTLOOK_CALENDAR_USER is missing.")
    if not cfg.outlook_tenant_id:
        errors.append("ASME_OUTLOOK_TENANT_ID is missing.")
    elif cfg.outlook_tenant_id.lower() in {"common", "organizations", "consumers"}:
        errors.append("ASME_OUTLOOK_TENANT_ID must be your tenant ID/domain, not common/organizations/consumers.")
    if not cfg.outlook_client_id:
        errors.append("ASME_OUTLOOK_CLIENT_ID is missing.")
    if not cfg.outlook_client_secret:
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


def has_any_outlook_sync_inputs(cfg: Settings) -> bool:
    return bool(cfg.outlook_calendar_user or cfg.outlook_tenant_id or cfg.outlook_client_id or cfg.outlook_client_secret)


class OutlookCalendarProvider(CalendarProvider):
    name = "outlook"

    def __init__(self, cfg: Settings):
        self.cfg = cfg
        self._token_cache = {"access_token": None, "expires_at": 0, "tenant_id": "", "client_id": ""}

    def room_ids(self) -> dict[str, str]:
        return {
            "Robotics Room": self.cfg.outlook_calendar_robotics_id or self.cfg.outlook_calendar_id,
            "Fluids Lab": self.cfg.outlook_calendar_fluids_id or self.cfg.outlook_calendar_id,
        }

    def calendar_id_for_room(self, room: str) -> str:
        return self.room_ids().get(room) or self.cfg.outlook_calendar_id

    def status(self) -> ProviderStatus:
        errors, warnings = validate_outlook_sync_config(self.cfg)
        return ProviderStatus(
            provider=self.name,
            enabled=not errors,
            errors=errors,
            warnings=warnings,
            room_ids=self.room_ids(),
        )

    # -- auth -----------------------------------------------------------------

    def access_token(self) -> str:
        cfg = self.cfg
        if not (cfg.outlook_tenant_id and cfg.outlook_client_id and cfg.outlook_client_secret):
            raise CalendarError(
                "Outlook OAuth is not configured. Set ASME_OUTLOOK_TENANT_ID, "
                "ASME_OUTLOOK_CLIENT_ID, and ASME_OUTLOOK_CLIENT_SECRET."
            )
        now = time.time()
        cache = self._token_cache
        if (
            cache.get("access_token")
            and (cache.get("expires_at") or 0) > (now + 60)
            and cache.get("tenant_id") == cfg.outlook_tenant_id
            and cache.get("client_id") == cfg.outlook_client_id
        ):
            return cache["access_token"]

        token_url = f"https://login.microsoftonline.com/{cfg.outlook_tenant_id}/oauth2/v2.0/token"
        status, payload, error = http_json_request(
            method="POST",
            url=token_url,
            form_data={
                "grant_type": "client_credentials",
                "client_id": cfg.outlook_client_id,
                "client_secret": cfg.outlook_client_secret,
                "scope": "https://graph.microsoft.com/.default",
            },
            retries=2,
        )
        if status != 200:
            error_text = payload.get("error_description") or payload.get("error") or error or "token request failed"
            raise CalendarError(f"Outlook token request failed: {str(error_text)[:250]}")
        access_token = payload.get("access_token")
        if not access_token:
            raise CalendarError("Outlook token response missing access_token.")
        expires_in = int(payload.get("expires_in") or 0)
        cache.update(
            access_token=access_token,
            tenant_id=cfg.outlook_tenant_id,
            client_id=cfg.outlook_client_id,
            expires_at=now + max(60, expires_in - 120) if expires_in else (now + 900),
        )
        return access_token

    # -- operations -----------------------------------------------------------

    def free_busy(self, time_min: datetime, time_max: datetime):
        status = self.status()
        if not status.enabled:
            raise CalendarError(" ".join(status.errors))
        token = self.access_token()
        encoded_user = quote(self.cfg.outlook_calendar_user, safe="")
        busy_by_room = {room: [] for room in MEETING_ROOMS}
        for room in MEETING_ROOMS:
            calendar_id = self.calendar_id_for_room(room)
            if calendar_id:
                endpoint = f"{GRAPH}/users/{encoded_user}/calendars/{quote(calendar_id, safe='')}/calendarView"
            else:
                endpoint = f"{GRAPH}/users/{encoded_user}/calendarView"
            url = (
                f"{endpoint}?startDateTime={quote(time_min.isoformat())}"
                f"&endDateTime={quote(time_max.isoformat())}&$select=start,end&$top=200"
            )
            code, payload, error = http_json_request(
                "GET", url, headers={"Authorization": f"Bearer {token}"}, retries=2
            )
            if code != 200:
                raise CalendarError(f"Outlook calendarView failed for {room}: {error or code}")
            for entry in payload.get("value") or []:
                try:
                    start_dt = datetime.fromisoformat(((entry.get("start") or {}).get("dateTime") or "")[:19])
                    end_dt = datetime.fromisoformat(((entry.get("end") or {}).get("dateTime") or "")[:19])
                except Exception:
                    continue
                busy_by_room[room].append((start_dt.replace(tzinfo=time_min.tzinfo), end_dt.replace(tzinfo=time_min.tzinfo)))
        return busy_by_room

    def create_event(self, room, summary, description, start_local, end_local) -> CreatedEvent:
        status = self.status()
        if not status.enabled:
            raise CalendarError(" ".join(status.errors))
        token = self.access_token()
        calendar_id = self.calendar_id_for_room(room)
        event_body = {
            "subject": summary[:250],
            "body": {"contentType": "Text", "content": (description or "").strip() or "Created from ASME website scheduler."},
            "start": {"dateTime": start_local.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": self.cfg.outlook_calendar_tz},
            "end": {"dateTime": end_local.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": self.cfg.outlook_calendar_tz},
            "location": {"displayName": room},
        }
        encoded_user = quote(self.cfg.outlook_calendar_user, safe="")
        if calendar_id:
            endpoint = f"{GRAPH}/users/{encoded_user}/calendars/{quote(calendar_id, safe='')}/events"
        else:
            endpoint = f"{GRAPH}/users/{encoded_user}/events"
        code, payload, error = http_json_request(
            method="POST",
            url=endpoint,
            headers={"Authorization": f"Bearer {token}"},
            json_data=event_body,
            retries=2,
        )
        if code not in (200, 201):
            msg = payload.get("error", {}).get("message") if isinstance(payload.get("error"), dict) else None
            raise CalendarError(f"Outlook event create failed: {str(msg or error or 'event create failed')[:250]}")
        event_id = payload.get("id")
        if not event_id:
            raise CalendarError("Outlook event created but no event ID returned.")
        return CreatedEvent(external_id=event_id, calendar_id=calendar_id or None, link=payload.get("webLink"))

    def delete_event(self, calendar_id, external_id):
        if not external_id:
            return
        status = self.status()
        if not status.enabled:
            raise CalendarError(" ".join(status.errors))
        token = self.access_token()
        encoded_user = quote(self.cfg.outlook_calendar_user, safe="")
        endpoint = f"{GRAPH}/users/{encoded_user}/events/{quote(external_id, safe='')}"
        code, payload, error = http_json_request(
            method="DELETE", url=endpoint, headers={"Authorization": f"Bearer {token}"}, retries=2
        )
        if code in (200, 202, 204, 404):
            return
        msg = payload.get("error", {}).get("message") if isinstance(payload.get("error"), dict) else None
        raise CalendarError(f"Outlook event delete failed: {str(msg or error or 'event delete failed')[:250]}")

    def embed_url(self) -> str:
        return normalize_outlook_embed_url(self.cfg.outlook_calendar_embed_url)
