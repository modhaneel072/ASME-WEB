"""Outbox job handlers. Importing this module registers them."""

from __future__ import annotations

from datetime import datetime

from asme.config import settings
from asme.extensions import db
from asme.integrations import mail
from asme.integrations.calendar import CalendarError, get_provider
from asme.jobs.outbox import handler
from asme.models import CalendarSync, Event, User


@handler("calendar.create_event")
def calendar_create_event(payload: dict):
    event = db.session.get(Event, int(payload["event_id"]))
    if not event:
        return
    sync = CalendarSync.query.filter_by(subject_type="event", subject_id=event.id).order_by(CalendarSync.id.desc()).first()
    if sync is None:
        sync = CalendarSync(subject_type="event", subject_id=event.id, provider=get_provider().name, status="pending")
        db.session.add(sync)
    if sync.status == "synced" and sync.external_id:
        return
    provider = get_provider()
    details = [f"Organizer: {payload.get('organizer') or (event.requested_by.name if event.requested_by else '')}"]
    if event.description:
        details.append(event.description)
    try:
        created = provider.create_event(event.location or "", event.title, "\n".join(details), event.start_time, event.end_time)
    except CalendarError as exc:
        sync.status = "failed"
        sync.last_error = str(exc)[:500]
        db.session.commit()
        raise
    sync.provider = provider.name
    sync.external_id = created.external_id
    sync.calendar_id = created.calendar_id
    sync.link = created.link
    sync.status = "synced"
    sync.last_error = None
    sync.last_synced_at = datetime.utcnow()
    event.calendar_event_link = created.link
    if provider.name == "google":
        event.google_event_id = created.external_id
        event.google_calendar_id = created.calendar_id
    db.session.commit()


@handler("calendar.delete_event")
def calendar_delete_event(payload: dict):
    sync = db.session.get(CalendarSync, int(payload["sync_id"]))
    if not sync or not sync.external_id:
        return
    get_provider().delete_event(sync.calendar_id, sync.external_id)
    sync.status = "deleted"
    sync.last_synced_at = datetime.utcnow()
    db.session.commit()


@handler("mail.send")
def mail_send(payload: dict):
    mail.send_email(settings(), payload["to"], payload.get("subject") or "(no subject)", payload.get("body") or "")


@handler("stock.reconcile")
def stock_reconcile(payload: dict):
    from asme.services import inventory

    inventory.reconcile_stock()


@handler("onboarding.evaluate")
def onboarding_evaluate(payload: dict):
    from asme.services.onboarding import engine

    if payload.get("user_id"):
        user = db.session.get(User, int(payload["user_id"]))
        if user:
            engine.evaluate_user(user)
    if payload.get("chapter"):
        engine.evaluate_chapter()


@handler("onboarding.evaluate_all")
def onboarding_evaluate_all(payload: dict):
    from asme.services.onboarding import engine

    for user in User.query.filter(User.is_active.is_(True)).all():
        engine.evaluate_user(user)
    engine.evaluate_chapter()
