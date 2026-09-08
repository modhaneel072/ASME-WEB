"""Launchpad - the onboarding engine.

``register_event_handlers()`` wires the engine to the domain event bus so any
service write that could change a task's state triggers a re-evaluation.
"""

from __future__ import annotations

import logging

from asme import events
from asme.extensions import db
from asme.models import User
from asme.services.onboarding import engine, entitlements, rules, seeds

log = logging.getLogger("asme.onboarding")

USER_EVENTS = (
    events.LOAN_OPENED,
    events.LOAN_RETURNED,
    events.ATTENDANCE_RECORDED,
    events.PRINT_SUBMITTED,
    events.PRINT_UPDATED,
    events.TRAINING_COMPLETED,
    events.USER_UPDATED,
    events.USER_CREATED,
    events.NFC_ASSIGNED,
    events.TEAM_JOINED,
    events.HOURS_LOGGED,
    events.TASK_SIGNED_OFF,
)
CHAPTER_EVENTS = (events.CHAPTER_CHANGED, events.USER_CREATED, events.NFC_ASSIGNED)


def _on_user_event(name, **payload):
    user_id = payload.get("user_id")
    if not user_id:
        return
    user = db.session.get(User, int(user_id))
    if not user:
        return
    engine.evaluate_user(user)


def _on_chapter_event(name, **payload):
    engine.evaluate_chapter()


def register_event_handlers():
    for name in USER_EVENTS:
        events.subscribe(name, _on_user_event)
    for name in CHAPTER_EVENTS:
        events.subscribe(name, _on_chapter_event)


__all__ = ["engine", "entitlements", "rules", "seeds", "register_event_handlers"]
