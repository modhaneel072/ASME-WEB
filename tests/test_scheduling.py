from datetime import datetime, timedelta

import pytest

from asme.integrations.calendar.base import CalendarProvider, CreatedEvent, ProviderStatus
from asme.models import Event, OutboxJob
from asme.services import scheduling
from asme.services.errors import Conflict


class FakeProvider(CalendarProvider):
    name = "google"

    def __init__(self, busy=None):
        self.busy = busy or {}

    def status(self):
        return ProviderStatus(provider="google", enabled=True, room_ids={"Robotics Room": "r", "Fluids Lab": "f"})

    def free_busy(self, time_min, time_max):
        return {"Robotics Room": self.busy.get("Robotics Room", []), "Fluids Lab": self.busy.get("Fluids Lab", [])}

    def create_event(self, room, summary, description, start_local, end_local):
        return CreatedEvent(external_id="x", calendar_id="r", link="https://cal/x")


def test_available_slots_exclude_busy_and_existing(app, db, users):
    app.extensions["asme_calendar_provider"] = FakeProvider()
    slots, error = scheduling.compute_available_slots(60, 2, "Robotics Room")
    assert error is None and slots
    token = slots[0]["token"]
    event = scheduling.book_slot(users["lead"], "Rover", token, notes="kickoff")
    assert event.status == "scheduled" and event.location == "Robotics Room"
    # the booked slot is gone from availability
    slots_after, _ = scheduling.compute_available_slots(60, 2, "Robotics Room")
    assert token not in {s["token"] for s in slots_after}
    # and a calendar job was queued in the same transaction
    assert OutboxJob.query.filter_by(kind="calendar.create_event").count() == 1


def test_double_booking_conflicts(app, db, users):
    app.extensions["asme_calendar_provider"] = FakeProvider()
    slots, _ = scheduling.compute_available_slots(60, 1, "Fluids Lab")
    token = slots[0]["token"]
    scheduling.book_slot(users["lead"], "A", token)
    with pytest.raises(Conflict):
        scheduling.book_slot(users["lead"], "B", token)


def test_bad_token_is_validation(app, users):
    from asme.services.errors import Validation

    with pytest.raises(Validation):
        scheduling.book_slot(users["lead"], "A", "garbage")


def test_admin_event_conflict_detection(app, users):
    start = datetime.now() + timedelta(days=1)
    scheduling.create_event(
        {"title": "GBM", "location": "Robotics Room", "start_time": start.strftime("%Y-%m-%dT%H:%M"), "end_time": (start + timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M")},
        users["admin"],
    )
    with pytest.raises(Conflict):
        scheduling.create_event(
            {"title": "Clash", "location": "robotics room", "start_time": (start + timedelta(minutes=30)).strftime("%Y-%m-%dT%H:%M"), "end_time": (start + timedelta(hours=2)).strftime("%Y-%m-%dT%H:%M")},
            users["admin"],
        )
    assert Event.query.count() == 1
