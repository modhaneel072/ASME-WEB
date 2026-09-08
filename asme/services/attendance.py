"""Attendance: check-ins against events, RSVPs, the legacy day-based scan."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta

from asme import events
from asme.config import settings
from asme.extensions import db
from asme.models import AttendanceRecord, AttendanceScan, Event, Member
from asme.services import audit
from asme.services.errors import Conflict, NotFound, Validation
from asme.utils import clean_tag_value


def candidate_events(now=None):
    """Events a tap right now could plausibly be for: running, or starting soon."""
    now = now or datetime.now()
    soon = now + timedelta(minutes=settings().checkin_soon_window_minutes)
    return (
        Event.query.filter(Event.status == "scheduled", Event.end_time >= now, Event.start_time <= soon)
        .order_by(Event.start_time.asc(), Event.id.asc())
        .all()
    )


def ensure_open_shop_event(day: date | None = None, created_by=None) -> Event:
    """Auto-opened event that walk-in scans attach to when no meeting is scheduled."""
    day = day or date.today()
    start = datetime.combine(day, time(hour=0, minute=0))
    end = datetime.combine(day, time(hour=23, minute=59))
    existing = Event.query.filter(Event.kind == "open_shop", Event.start_time == start).first()
    if existing:
        return existing
    event = Event(
        title=f"Open shop - {day.isoformat()}",
        kind="open_shop",
        location=None,
        status="scheduled",
        start_time=start,
        end_time=end,
        created_by_user_id=created_by.id if created_by else None,
    )
    db.session.add(event)
    db.session.flush()
    return event


def check_user_into_event(user, event, method="shared_tag", tag_uid=None):
    """Idempotent check-in. Returns ``(record, created)``."""
    from asme.services.identity import member_for_user

    existing = (
        AttendanceRecord.query.filter(AttendanceRecord.event_id == event.id, AttendanceRecord.user_id == user.id)
        .order_by(AttendanceRecord.id.desc())
        .first()
    )
    if existing:
        return existing, False
    member = member_for_user(user)
    row = AttendanceRecord(
        event_id=event.id,
        user_id=user.id,
        member_id=member.id if member else None,
        tag_uid=clean_tag_value(tag_uid) or None,
        checkin_method=(method or "shared_tag")[:40],
        checkin_time=datetime.utcnow(),
    )
    db.session.add(row)
    db.session.commit()
    events.emit(events.ATTENDANCE_RECORDED, user_id=user.id, event_id=event.id, method=row.checkin_method)
    return row, True


def admin_checkin(event, *, tag_uid=None, user_id=None, requested_method=None, actor=None):
    from asme.services.identity import member_for_user, resolve_user_from_tag_uid
    from asme.models import User

    tag_uid = clean_tag_value(tag_uid)
    user = None
    checkin_method = "manual"
    if tag_uid:
        user = resolve_user_from_tag_uid(tag_uid)
        if not user:
            raise NotFound("Tag UID not assigned to an active user.")
        checkin_method = requested_method if requested_method in {"kiosk", "nfc"} else "nfc"
    elif user_id:
        user = db.session.get(User, user_id)
        if not user:
            raise NotFound("Selected user not found.")
    else:
        raise Validation("Provide tag UID or select a user.")

    member = member_for_user(user)
    existing = (
        AttendanceRecord.query.filter_by(event_id=event.id)
        .filter(AttendanceRecord.user_id == user.id)
        .order_by(AttendanceRecord.id.desc())
        .first()
    )
    if existing:
        raise Conflict("This person is already checked in for the selected event.", code="already_checked_in")

    row = AttendanceRecord(
        event_id=event.id,
        user_id=user.id,
        member_id=member.id if member else None,
        tag_uid=tag_uid or None,
        checkin_method=checkin_method,
        checkin_time=datetime.utcnow(),
    )
    db.session.add(row)
    audit.record("attendance_checkin", f"event_id={event.id} method={checkin_method}", actor=actor)
    db.session.commit()
    events.emit(events.ATTENDANCE_RECORDED, user_id=user.id, event_id=event.id, method=checkin_method)
    return row, user


def rsvp(user, event):
    from asme.services.identity import member_for_user

    existing = AttendanceRecord.query.filter_by(event_id=event.id, user_id=user.id, checkin_method="rsvp").first()
    if existing:
        return existing, False
    member = member_for_user(user)
    row = AttendanceRecord(
        event_id=event.id,
        user_id=user.id,
        member_id=member.id if member else None,
        checkin_method="rsvp",
        checkin_time=datetime.utcnow(),
    )
    db.session.add(row)
    db.session.commit()
    return row, True


def legacy_scan(uid) -> tuple[Member, bool]:
    """Legacy day-based attendance: writes ``AttendanceScan`` and also an event-based
    record against the day's open-shop event so the Launchpad sees it."""
    from asme.services.identity import user_for_member

    uid = (uid or "").strip()
    if not uid:
        raise Validation("Scan failed: UID was empty.")
    member = Member.query.filter_by(nfc_tag=uid).first()
    if not member:
        raise NotFound("UID not recognized. Pair this UID to a member first.")
    scan = AttendanceScan(member_id=member.id, scanned_uid=uid, attendance_date=date.today())
    db.session.add(scan)
    db.session.commit()
    first_today = AttendanceScan.query.filter_by(member_id=member.id, attendance_date=date.today()).count() == 1
    user = user_for_member(member)
    if user and first_today:
        event = ensure_open_shop_event()
        db.session.commit()
        check_user_into_event(user, event, method="nfc", tag_uid=uid)
    return member, first_today


def today_unique_scans():
    scans = AttendanceScan.query.filter_by(attendance_date=date.today()).order_by(AttendanceScan.scanned_at.desc()).all()
    unique, seen = [], set()
    for scan in scans:
        if scan.member_id in seen:
            continue
        seen.add(scan.member_id)
        unique.append(scan)
    return unique


def export_rows():
    rows = AttendanceRecord.query.order_by(AttendanceRecord.checkin_time.desc(), AttendanceRecord.id.desc()).all()
    output = []
    for row in rows:
        output.append(
            [
                row.id,
                row.event_id,
                row.event.title if row.event else "",
                row.user_id or "",
                row.user.name if row.user else (row.member.name if row.member else ""),
                row.tag_uid or "",
                row.checkin_method or "",
                row.checkin_time.isoformat() if row.checkin_time else "",
            ]
        )
    return output


EXPORT_HEADERS = ["record_id", "event_id", "event_title", "user_id", "user_name", "tag_uid", "method", "checkin_time"]


def today_count():
    today_start = datetime.combine(date.today(), time.min)
    return AttendanceRecord.query.filter(
        AttendanceRecord.checkin_time >= today_start,
        AttendanceRecord.checkin_time < today_start + timedelta(days=1),
    ).count()
