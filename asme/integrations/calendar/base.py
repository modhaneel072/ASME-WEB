"""Calendar provider interface.

``scheduling.py`` holds one ``CalendarProvider`` chosen from config and never
learns which concrete class it has.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class ProviderStatus:
    provider: str
    enabled: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    room_ids: dict[str, str] = field(default_factory=dict)


@dataclass
class CreatedEvent:
    external_id: str | None
    calendar_id: str | None
    link: str | None


class CalendarError(RuntimeError):
    """Raised by providers for any failure the caller should surface to a human."""


class CalendarProvider:
    name = "none"

    def status(self) -> ProviderStatus:
        raise NotImplementedError

    def free_busy(self, time_min: datetime, time_max: datetime) -> dict[str, list[tuple[datetime, datetime]]]:
        """Busy windows keyed by room name, in the provider's local timezone."""
        raise NotImplementedError

    def create_event(
        self,
        room: str,
        summary: str,
        description: str,
        start_local: datetime,
        end_local: datetime,
    ) -> CreatedEvent:
        raise NotImplementedError

    def delete_event(self, calendar_id: str | None, external_id: str) -> None:
        raise NotImplementedError

    def embed_url(self) -> str:
        return ""


class NullProvider(CalendarProvider):
    """Used when nothing is configured; every operation reports "not configured"."""

    name = "none"

    def __init__(self, reason: str = "No calendar provider is configured."):
        self.reason = reason

    def status(self) -> ProviderStatus:
        return ProviderStatus(provider=self.name, enabled=False, errors=[self.reason])

    def free_busy(self, time_min, time_max):
        raise CalendarError(self.reason)

    def create_event(self, room, summary, description, start_local, end_local):
        raise CalendarError(self.reason)

    def delete_event(self, calendar_id, external_id):
        raise CalendarError(self.reason)
