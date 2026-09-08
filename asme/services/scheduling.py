"""Room scheduling: availability, bookings, admin events, legacy meetings.

Bookings commit the ``Event`` and an ``OutboxJob`` together; the worker mirrors
the event to the external calendar and records the result in ``CalendarSync``.
"""

from __future__ import annotations

import secrets
from datetime import date, datetime, time, timedelta

from sqlalchemy import and_, func, or_

from asme import events
from asme.config import settings
from asme.constants import EVENT_STATUSES, MEETING_ROOMS
from asme.extensions import db
from asme.integrations.calendar import CalendarError, get_provider
from asme.models import CalendarSync, Event, Meeting
from asme.services import audit
from asme.services.errors import Conflict, NotFound, Validation
from asme.utils import parse_clock_time


def normalize_meeting_room(raw_location):
    location = (raw_location or "").strip()
    if not location:
        return ""
    return {room.lower(): room for room in MEETING_ROOMS}.get(location.lower(), "")


def work_hours():
    cfg = settings()
    start = parse_clock_time(cfg.calendar_work_hours_start) or time(hour=8)
    end = parse_clock_time(cfg.calendar_work_hours_end) or time(hour=22)
    if end <= start:
        end = time(hour=22)
    return start, end


def find_event_room_conflict(room, start_time, end_time, exclude_event_id=None):
    if not room:
        return None
    query = Event.query.filter(
        func.lower(func.coalesce(Event.location, "")) == room.lower(),
        Event.status.in_(["requested", "scheduled"]),
        Event.start_time < end_time,
        Event.end_time > start_time,
    )
    if exclude_event_id:
        query = query.filter(Event.id != exclude_event_id)
    return query.order_by(Event.start_time.asc(), Event.id.asc()).first()


def _overlaps(start_a, end_a, start_b, end_b):
    return start_a < end_b and end_a > start_b


def compute_available_slots(duration_minutes, day_count, preferred_room=""):
    """Slots free on both the external calendar and our own event table.

    Returns ``(slots, error)`` - ``error`` is a human message when the provider
    is unavailable, in which case ``slots`` is empty.
    """
    provider = get_provider()
    day_count = max(1, min(int(day_count or 14), 30))
    duration_minutes = max(30, min(int(duration_minutes or 60), 180))
    duration_delta = timedelta(minutes=duration_minutes)
    tzinfo = getattr(provider, "tzinfo", None)
    tzinfo = tzinfo() if callable(tzinfo) else None
    now = datetime.now(tzinfo)
    work_start, work_end = work_hours()
    today = now.date()
    range_start = datetime.combine(today, work_start, tzinfo=tzinfo)
    range_end = datetime.combine(today + timedelta(days=day_count - 1), work_end, tzinfo=tzinfo)

    try:
        busy_by_room = provider.free_busy(range_start, range_end)
    except CalendarError as exc:
        return [], str(exc)

    rooms = [preferred_room] if preferred_room in MEETING_ROOMS else list(MEETING_ROOMS)
    slots = []
    for day_offset in range(day_count):
        day_value = today + timedelta(days=day_offset)
        for room in rooms:
            day_start = datetime.combine(day_value, work_start, tzinfo=tzinfo)
            day_end = datetime.combine(day_value, work_end, tzinfo=tzinfo)
            cursor = day_start
            while cursor + duration_delta <= day_end:
                slot_start, slot_end = cursor, cursor + duration_delta
                cursor += timedelta(minutes=30)
                if slot_start < now + timedelta(minutes=2):
                    continue
                if any(_overlaps(slot_start, slot_end, b_start, b_end) for b_start, b_end in busy_by_room.get(room, [])):
                    continue
                local_start, local_end = slot_start.replace(tzinfo=None), slot_end.replace(tzinfo=None)
                if find_event_room_conflict(room, local_start, local_end):
                    continue
                slots.append(
                    {
                        "room": room,
                        "start": slot_start,
                        "end": slot_end,
                        "token": f"{room}|{slot_start.strftime('%Y-%m-%dT%H:%M')}|{duration_minutes}",
                    }
                )
    return slots, None


def parse_slot_token(slot_token):
    try:
        room, start_token, duration_token = (slot_token or "").split("|", 2)
        room = normalize_meeting_room(room)
        start_dt = datetime.strptime(start_token, "%Y-%m-%dT%H:%M")
        duration_minutes = int(duration_token)
    except Exception:
        raise Validation("Select a valid available slot.", field="slot_token")
    if not room:
        raise Validation("Select a valid room.", field="slot_token")
    return room, start_dt, start_dt + timedelta(minutes=duration_minutes)


def book_slot(user, team_name, slot_token, notes="", sync=True) -> Event:
    """Create a scheduled ``Event`` for a room slot and queue the calendar mirror."""
    from asme.jobs.outbox import enqueue

    team_name = (team_name or "").strip() or "Team"
    room, start_dt, end_dt = parse_slot_token(slot_token)
    if find_event_room_conflict(room, start_dt, end_dt):
        raise Conflict("Selected slot is no longer available. Please refresh availability.", code="slot_taken")

    event = Event(
        title=f"ASME - {team_name} Meeting"[:220],
        description=(notes or "")[:4000] or None,
        kind="meeting",
        location=room,
        status="scheduled",
        start_time=start_dt,
        end_time=end_dt,
        requested_by_user_id=user.id,
        created_by_user_id=user.id,
    )
    db.session.add(event)
    db.session.flush()
    # Conflict check inside the same transaction as the insert - a second request
    # that slipped through the pre-check loses here instead of double booking.
    if find_event_room_conflict(room, start_dt, end_dt, exclude_event_id=event.id):
        db.session.rollback()
        raise Conflict("Selected slot was just taken. Please refresh availability.", code="slot_taken")

    provider = get_provider()
    if sync and provider.status().enabled:
        db.session.add(CalendarSync(subject_type="event", subject_id=event.id, provider=provider.name, status="pending"))
        enqueue("calendar.create_event", {"event_id": event.id, "organizer": user.name})
    audit.record("schedule_meeting", f"room={room} start={start_dt.isoformat()} by={user.email}", actor=user)
    db.session.commit()
    events.emit(events.BOOKING_CREATED, event_id=event.id, user_id=user.id)
    return event


def sync_status_for_event(event_id):
    return CalendarSync.query.filter_by(subject_type="event", subject_id=event_id).order_by(CalendarSync.id.desc()).first()


# --------------------------------------------------------------------------- admin events


def create_event(form, actor) -> Event:
    from asme.utils import parse_datetime_local

    title = (form.get("title") or "").strip()
    description = (form.get("description") or "").strip() or None
    room = normalize_meeting_room(form.get("location"))
    start_time = parse_datetime_local(form.get("start_time"))
    end_time = parse_datetime_local(form.get("end_time"))
    status = (form.get("status") or "scheduled").strip().lower()
    kind = (form.get("kind") or "meeting").strip().lower() or "meeting"
    if status not in EVENT_STATUSES:
        status = "scheduled"
    if not title or not room or not start_time or not end_time or end_time <= start_time:
        raise Validation("Enter valid title, room, start, and end times.")
    if status in {"requested", "scheduled"}:
        conflict = find_event_room_conflict(room, start_time, end_time)
        if conflict:
            raise Conflict(
                f"{room} conflicts with '{conflict.title}' "
                f"({conflict.start_time.strftime('%Y-%m-%d %H:%M')} - {conflict.end_time.strftime('%H:%M')}).",
                code="room_conflict",
            )
    event = Event(
        title=title[:220],
        description=description,
        kind=kind[:40],
        location=room,
        status=status,
        start_time=start_time,
        end_time=end_time,
        created_by_user_id=actor.id if actor else None,
    )
    db.session.add(event)
    audit.record("create_event", title[:220], actor=actor)
    db.session.commit()
    events.emit(events.CHAPTER_CHANGED, reason="event_created")
    return event


def update_event(event, form, actor) -> Event:
    from asme.utils import parse_datetime_local

    status = (form.get("status") or event.status or "scheduled").strip().lower()
    if status not in EVENT_STATUSES:
        status = event.status or "scheduled"
    title = (form.get("title") or event.title or "").strip()
    start_time = parse_datetime_local(form.get("start_time")) or event.start_time
    end_time = parse_datetime_local(form.get("end_time")) or event.end_time
    room = normalize_meeting_room(form.get("location") or event.location)
    if not title or not room or not start_time or not end_time or end_time <= start_time:
        raise Validation("Enter valid title, room, start, and end times.")
    if status in {"requested", "scheduled"}:
        conflict = find_event_room_conflict(room, start_time, end_time, exclude_event_id=event.id)
        if conflict:
            raise Conflict(
                f"Cannot update event. {room} conflicts with '{conflict.title}' "
                f"({conflict.start_time.strftime('%Y-%m-%d %H:%M')} - {conflict.end_time.strftime('%H:%M')}).",
                code="room_conflict",
            )
    event.title = title[:220]
    event.location = room
    event.start_time = start_time
    event.end_time = end_time
    event.status = status
    audit.record("update_event_status", f"event_id={event.id} status={status}", actor=actor)
    db.session.commit()
    events.emit(events.CHAPTER_CHANGED, reason="event_updated")
    return event


def upcoming_events(limit=25, since=None):
    since = since or (datetime.now() - timedelta(hours=1))
    return (
        Event.query.filter(Event.status != "cancelled", Event.end_time >= since)
        .order_by(Event.start_time.asc(), Event.id.asc())
        .limit(limit)
        .all()
    )


# --------------------------------------------------------------------------- legacy meetings


def upcoming_meetings():
    now = datetime.now()
    return (
        Meeting.query.filter(
            or_(
                Meeting.meeting_date > now.date(),
                and_(Meeting.meeting_date == now.date(), Meeting.end_time >= now.time()),
            )
        )
        .filter(Meeting.cancel_request_token.is_(None))
        .order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc())
        .all()
    )


def legacy_book_meeting(form) -> tuple[Meeting, str | None, bool, bool]:
    """Legacy ops booking with synchronous Outlook mirror.

    Returns ``(meeting, calendar_error, sync_configured, config_has_inputs)``.
    """
    from asme.integrations.calendar.outlook import OutlookCalendarProvider, has_any_outlook_sync_inputs
    from asme.utils import parse_due_date

    cfg = settings()
    team_name = (form.get("team_name") or "").strip()
    requester_email = (form.get("requester_email") or "").strip() or None
    room = (form.get("room") or "").strip()
    meeting_date = parse_due_date(form.get("meeting_date"))
    start_time = parse_clock_time(form.get("start_time"))
    end_time = parse_clock_time(form.get("end_time"))
    notes = (form.get("notes") or "").strip() or None

    if not team_name:
        raise Validation("Team name is required.", field="team_name")
    if room not in MEETING_ROOMS:
        raise Validation("Please select Robotics Room or Fluids Lab.", field="room")
    if not meeting_date:
        raise Validation("Meeting date is required.", field="meeting_date")
    if meeting_date < date.today():
        raise Validation("Meeting date cannot be in the past.", field="meeting_date")
    if not start_time or not end_time:
        raise Validation("Start and end times are required.")
    if end_time <= start_time:
        raise Validation("End time must be after start time.", field="end_time")

    conflicting = (
        Meeting.query.filter(
            Meeting.room == room,
            Meeting.meeting_date == meeting_date,
            Meeting.start_time < end_time,
            Meeting.end_time > start_time,
            Meeting.cancel_request_token.is_(None),
        )
        .order_by(Meeting.start_time.asc(), Meeting.id.asc())
        .first()
    )
    if conflicting:
        raise Conflict(
            f"{room} is already booked by {conflicting.team_name} from "
            f"{conflicting.start_time.strftime('%H:%M')} to {conflicting.end_time.strftime('%H:%M')}.",
            code="room_conflict",
        )

    meeting = Meeting(
        team_name=team_name,
        requester_email=requester_email,
        room=room,
        meeting_date=meeting_date,
        start_time=start_time,
        end_time=end_time,
        notes=notes,
    )
    outlook = OutlookCalendarProvider(cfg)
    sync_configured = outlook.status().enabled
    config_has_inputs = has_any_outlook_sync_inputs(cfg)
    calendar_error = None
    if sync_configured:
        description_lines = []
        if requester_email:
            description_lines.append(f"Requested by: {requester_email}")
        if notes:
            description_lines.append(f"Notes: {notes}")
        try:
            created = outlook.create_event(
                room,
                f"{team_name} - {room}",
                "\n".join(description_lines),
                datetime.combine(meeting_date, start_time),
                datetime.combine(meeting_date, end_time),
            )
            meeting.outlook_calendar_id = created.calendar_id
            meeting.outlook_event_id = created.external_id
        except CalendarError as exc:
            calendar_error = str(exc)
    db.session.add(meeting)
    db.session.commit()
    return meeting, calendar_error, sync_configured, config_has_inputs


def legacy_request_cancel(meeting) -> str:
    if meeting.cancel_request_token:
        raise Conflict("Cancellation already requested and waiting for email confirmation.", code="pending")
    token = secrets.token_urlsafe(32)
    meeting.cancel_request_token = token
    meeting.cancel_requested_at = datetime.utcnow()
    db.session.commit()
    return token


def legacy_clear_cancel(meeting):
    meeting.cancel_request_token = None
    meeting.cancel_requested_at = None
    db.session.commit()


def legacy_confirm_cancel(meeting) -> str | None:
    """Delete the mirrored Outlook event then the meeting. Returns an error string or None."""
    from asme.integrations.calendar.outlook import OutlookCalendarProvider

    if meeting.outlook_event_id:
        try:
            OutlookCalendarProvider(settings()).delete_event(meeting.outlook_calendar_id, meeting.outlook_event_id)
        except CalendarError as exc:
            return str(exc)
    db.session.delete(meeting)
    db.session.commit()
    return None


def find_meeting_by_cancel_token(token):
    return Meeting.query.filter_by(cancel_request_token=token).first()


def provider_status():
    return get_provider().status()


def calendar_embed_url():
    cfg = settings()
    provider = get_provider()
    url = provider.embed_url()
    if url:
        return url
    if cfg.google_calendar_embed_url:
        return cfg.google_calendar_embed_url
    if cfg.outlook_calendar_embed_url:
        from asme.integrations.calendar.outlook import normalize_outlook_embed_url

        return normalize_outlook_embed_url(cfg.outlook_calendar_embed_url)
    return ""
