from datetime import datetime, timedelta

from asme.jobs import outbox
from asme.models import CalendarSync, Event, OutboxJob


def test_enqueue_and_process(app, db):
    calls = []

    @outbox.handler("test.echo")
    def _echo(payload):
        calls.append(payload)

    outbox.enqueue("test.echo", {"x": 1})
    db.session.commit()
    assert outbox.process_pending() == 1
    assert calls == [{"x": 1}]
    assert OutboxJob.query.first().status == "done"


def test_failure_retries_then_fails(app, db):
    @outbox.handler("test.boom")
    def _boom(payload):
        raise RuntimeError("nope")

    outbox.enqueue("test.boom", {}, max_attempts=2)
    db.session.commit()
    assert outbox.process_pending() == 0
    job = OutboxJob.query.first()
    assert job.status == "pending" and job.attempts == 1 and "nope" in job.last_error
    job.run_at = datetime.utcnow() - timedelta(seconds=1)
    db.session.commit()
    assert outbox.process_pending() == 0
    db.session.refresh(job)
    assert job.status == "failed" and job.attempts == 2
    outbox.retry(job)
    assert job.status == "pending" and job.attempts == 0


def test_unknown_kind_fails(app, db):
    outbox.enqueue("nothing.registered", {}, max_attempts=1)
    db.session.commit()
    outbox.process_pending()
    assert OutboxJob.query.first().status == "failed"


def test_ensure_recurring_is_once_per_window(app, db):
    assert outbox.ensure_recurring("stock.reconcile", timedelta(hours=24)) is True
    assert outbox.ensure_recurring("stock.reconcile", timedelta(hours=24)) is False


def test_calendar_job_records_sync(app, db, users, monkeypatch):
    from asme.integrations.calendar.base import CalendarProvider, CreatedEvent, ProviderStatus

    class Fake(CalendarProvider):
        name = "google"

        def status(self):
            return ProviderStatus(provider="google", enabled=True, room_ids={"Robotics Room": "cal-1"})

        def create_event(self, room, summary, description, start_local, end_local):
            return CreatedEvent(external_id="ext-1", calendar_id="cal-1", link="https://cal/ext-1")

    app.extensions["asme_calendar_provider"] = Fake()
    event = Event(title="Meet", location="Robotics Room", status="scheduled", start_time=datetime.now(), end_time=datetime.now() + timedelta(hours=1), requested_by_user_id=users["lead"].id)
    db.session.add(event)
    db.session.flush()
    db.session.add(CalendarSync(subject_type="event", subject_id=event.id, provider="google", status="pending"))
    outbox.enqueue("calendar.create_event", {"event_id": event.id, "organizer": "Lee"})
    db.session.commit()
    assert outbox.process_pending() == 1
    sync = CalendarSync.query.filter_by(subject_id=event.id).first()
    assert sync.status == "synced" and sync.external_id == "ext-1"
    assert event.calendar_event_link == "https://cal/ext-1"
