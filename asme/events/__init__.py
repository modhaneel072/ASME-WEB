"""Tiny in-process domain event bus.

Services ``emit`` after their transaction commits; subscribers (today: the
onboarding engine) react. A failing subscriber is logged and never propagates
into the request that emitted the event - onboarding must not break a checkout.
"""

from __future__ import annotations

import logging
from collections import defaultdict
from typing import Callable

log = logging.getLogger("asme.events")

_subscribers: dict[str, list[Callable[..., None]]] = defaultdict(list)

# Event names emitted by services. Keep this list current; tests assert on it.
LOAN_OPENED = "loan.opened"
LOAN_RETURNED = "loan.returned"
ATTENDANCE_RECORDED = "attendance.recorded"
PRINT_SUBMITTED = "print.submitted"
PRINT_UPDATED = "print.updated"
TRAINING_COMPLETED = "training.completed"
USER_UPDATED = "user.updated"
USER_CREATED = "user.created"
NFC_ASSIGNED = "nfc.assigned"
TEAM_JOINED = "team.joined"
HOURS_LOGGED = "hours.logged"
TASK_SIGNED_OFF = "task.signed_off"
BOOKING_CREATED = "booking.created"
CHAPTER_CHANGED = "chapter.changed"

ALL_EVENTS = (
    LOAN_OPENED,
    LOAN_RETURNED,
    ATTENDANCE_RECORDED,
    PRINT_SUBMITTED,
    PRINT_UPDATED,
    TRAINING_COMPLETED,
    USER_UPDATED,
    USER_CREATED,
    NFC_ASSIGNED,
    TEAM_JOINED,
    HOURS_LOGGED,
    TASK_SIGNED_OFF,
    BOOKING_CREATED,
    CHAPTER_CHANGED,
)


def subscribe(name: str, handler: Callable[..., None]) -> None:
    if handler not in _subscribers[name]:
        _subscribers[name].append(handler)


def unsubscribe(name: str, handler: Callable[..., None]) -> None:
    if handler in _subscribers.get(name, []):
        _subscribers[name].remove(handler)


def clear() -> None:
    _subscribers.clear()


def emit(name: str, **payload) -> int:
    """Dispatch ``name`` to every subscriber. Returns how many handlers ran."""
    ran = 0
    for handler in list(_subscribers.get(name, [])) + list(_subscribers.get("*", [])):
        try:
            handler(name, **payload)
            ran += 1
        except Exception:  # pragma: no cover - defensive, exercised by tests via a bad handler
            log.exception("event handler failed: event=%s handler=%r", name, handler)
    return ran
