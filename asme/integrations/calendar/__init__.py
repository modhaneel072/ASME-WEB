"""Calendar provider factory."""

from __future__ import annotations

from flask import current_app

from asme.config import Settings
from asme.integrations.calendar.base import CalendarError, CalendarProvider, CreatedEvent, NullProvider, ProviderStatus
from asme.integrations.calendar.google import GoogleCalendarProvider
from asme.integrations.calendar.outlook import OutlookCalendarProvider

__all__ = [
    "CalendarError",
    "CalendarProvider",
    "CreatedEvent",
    "NullProvider",
    "ProviderStatus",
    "build_provider",
    "get_provider",
]


def build_provider(cfg: Settings) -> CalendarProvider:
    if cfg.calendar_provider == "outlook":
        return OutlookCalendarProvider(cfg)
    return GoogleCalendarProvider(cfg)


def get_provider() -> CalendarProvider:
    """The app-wide provider instance (created once in the app factory)."""
    provider = current_app.extensions.get("asme_calendar_provider")
    if provider is None:
        provider = build_provider(current_app.config["SETTINGS"])
        current_app.extensions["asme_calendar_provider"] = provider
    return provider
