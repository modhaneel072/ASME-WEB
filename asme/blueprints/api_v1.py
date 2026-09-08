"""``/api/v1`` - the versioned JSON surface.

Conventions:
* one error shape: ``{"ok": false, "code", "error", ...}``
* entitlement failures carry the phase that unlocks them
* cursor pagination on collections, ETag on the catalog
* ``Idempotency-Key`` honoured on kiosk-triggerable writes
"""

from __future__ import annotations

import base64
import hashlib
import json
from datetime import datetime

from flask import Blueprint, Response, jsonify, request

from asme.auth.session import current_auth_user, current_user_member, require_entitlement, require_login, require_role, role_allows
from asme.constants import ENT_EVENT_CHECKIN, ENT_PORTAL_ACCESS, ENT_PRINT_APPROVE, ENT_PRINT_SUBMIT, ENT_ROOM_BOOKING, ENT_SHOP_ACCESS
from asme.extensions import db
from asme.models import AttendanceRecord, Event, Item, Loan, PrintRequest, User
from asme.services import attendance, fabrication, identity, inventory, roster, scheduling
from asme.services.errors import Forbidden, NotFound, ServiceError, Validation
from asme.services.onboarding import engine, entitlements
from asme.services.serializers import (
    serialize_attendance_record,
    serialize_event,
    serialize_item,
    serialize_loan,
    serialize_print_request,
    serialize_user,
)
from asme.utils import clean_tag_value, parse_due_date, parse_positive_int
from asme.utils.http import api_ok, idempotency_key, value_from_request

bp = Blueprint("api_v1", __name__, url_prefix="/api/v1")


@bp.errorhandler(ServiceError)
def _service_error(exc):
    return jsonify(exc.to_dict()), exc.status


@bp.errorhandler(404)
def _not_found(_exc):
    return jsonify({"ok": False, "code": "not_found", "error": "Not found."}), 404


def _json():
    return request.get_json(silent=True) or {}


def _encode_cursor(value) -> str:
    return base64.urlsafe_b64encode(json.dumps(value).encode("utf-8")).decode("ascii").rstrip("=")


def _decode_cursor(raw):
    if not raw:
        return None
    try:
        padded = raw + "=" * (-len(raw) % 4)
        return json.loads(base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8"))
    except Exception:
        raise Validation("Invalid cursor.", field="cursor", code="bad_cursor")


def _limit(default=50, maximum=200):
    return max(1, min(parse_positive_int(request.args.get("limit"), default=default), maximum))


@bp.get("/health")
def health():
    return api_ok({"service": "asme-web", "api": "v1", "time": datetime.utcnow().isoformat()})


# --------------------------------------------------------------------------- identity / onboarding


@bp.get("/me")
@require_login
def me():
    user = current_auth_user()
    summary = engine.summary_for_user(user)
    return api_ok({"user": serialize_user(user, entitlements=summary["entitlements"]), "launchpad": summary})


@bp.get("/onboarding")
@require_login
def onboarding():
    user = current_auth_user()
    target = user
    other_id = parse_positive_int(request.args.get("user_id"), default=0)
    if other_id and other_id != user.id:
        if not role_allows(user.role, "team_leader"):
            raise Forbidden("Only leads can view another member's Launchpad.")
        target = db.session.get(User, other_id)
        if not target:
            raise NotFound("User not found.")
    payload = engine.evaluate_user(target)
    if role_allows(user.role, "admin") and request.args.get("include_chapter"):
        payload["chapter"] = engine.evaluate_chapter()
    return api_ok(payload)


@bp.post("/onboarding/tasks/<task_key>/complete")
@require_login
def complete_task(task_key):
    actor = current_auth_user()
    body = _json()
    subject_type = (body.get("subject_type") or value_from_request("subject_type") or "user").strip().lower()
    note = body.get("note") or value_from_request("note")
    if subject_type == "chapter":
        state = engine.complete_task(task_key, engine.subject_for_chapter(), actor, note=note)
    else:
        target = actor
        other_id = body.get("user_id") or value_from_request("user_id")
        if other_id and int(other_id) != actor.id:
            target = db.session.get(User, int(other_id))
            if not target:
                raise NotFound("User not found.")
        state = engine.complete_task(task_key, engine.subject_for_user(target), actor, note=note)
    return api_ok({"task": task_key, "status": state.status, "completed_at": state.completed_at.isoformat() if state.completed_at else None})


@bp.post("/onboarding/training")
@require_role("team_leader")
def record_training():
    actor = current_auth_user()
    body = _json()
    target = db.session.get(User, int(body.get("user_id") or value_from_request("user_id") or 0))
    if not target:
        raise NotFound("User not found.")
    score = body.get("score", value_from_request("score"))
    row = engine.record_training(target, body.get("module") or value_from_request("module"), score=int(score) if score not in (None, "") else None, actor=actor, note=body.get("note"))
    return api_ok({"id": row.id, "module": row.module.key, "score": row.score, "expires_at": row.expires_at.isoformat() if row.expires_at else None}, status=201)


@bp.get("/entitlements")
@require_login
def my_entitlements():
    return api_ok({"entitlements": sorted(entitlements.user_entitlements(current_auth_user()))})


# --------------------------------------------------------------------------- catalog


@bp.get("/items")
@require_login
@require_entitlement(ENT_PORTAL_ACCESS)
def list_items():
    limit = _limit()
    cursor = _decode_cursor(request.args.get("cursor"))
    query = Item.query.filter(Item.active.is_(True))
    q = (request.args.get("q") or "").strip().lower()
    if q:
        from sqlalchemy import func, or_

        like = f"%{q}%"
        query = query.filter(
            or_(func.lower(Item.name).like(like), func.lower(func.coalesce(Item.category, "")).like(like), func.lower(func.coalesce(Item.location, "")).like(like))
        )
    if cursor:
        query = query.filter(Item.id > int(cursor.get("after", 0)))
    rows = query.order_by(Item.id.asc()).limit(limit + 1).all()
    has_more = len(rows) > limit
    rows = rows[:limit]
    payload = {
        "items": [serialize_item(item) for item in rows],
        "next_cursor": _encode_cursor({"after": rows[-1].id}) if has_more and rows else None,
    }
    body = json.dumps({"ok": True, "payload": payload}, sort_keys=True, default=str)
    etag = hashlib.sha1(body.encode("utf-8")).hexdigest()
    if request.headers.get("If-None-Match") == etag:
        return Response(status=304, headers={"ETag": etag})
    return Response(body, status=200, mimetype="application/json", headers={"ETag": etag, "Cache-Control": "private, max-age=0"})


@bp.get("/items/by-tag/<path:uid>")
@require_login
@require_entitlement(ENT_PORTAL_ACCESS)
def item_by_tag(uid):
    item, resolved_via = inventory.find_item_by_tag(clean_tag_value(uid))
    if not item:
        raise NotFound("Tag not registered to an inventory item.", code="tag_unknown")
    user = current_auth_user()
    open_loan = inventory.get_open_loan(item.id, user=user, member=current_user_member(user))
    return api_ok({"item": serialize_item(item), "resolved_via": resolved_via, "open_loan": serialize_loan(open_loan) if open_loan else None})


# --------------------------------------------------------------------------- loans


@bp.get("/checkouts")
@require_login
@require_entitlement(ENT_PORTAL_ACCESS)
def list_checkouts():
    user = current_auth_user()
    state = (request.args.get("state") or "open").strip().lower()
    query = Loan.query
    if not role_allows(user.role, "admin") or request.args.get("mine"):
        member = current_user_member(user)
        if member:
            query = query.filter((Loan.user_id == user.id) | (Loan.member_id == member.id))
        else:
            query = query.filter(Loan.user_id == user.id)
    if state == "open":
        query = query.filter(Loan.status == "OUT")
    elif state == "returned":
        query = query.filter(Loan.status == "RETURNED")
    limit = _limit()
    cursor = _decode_cursor(request.args.get("cursor"))
    if cursor:
        query = query.filter(Loan.id < int(cursor.get("before", 0)))
    rows = query.order_by(Loan.id.desc()).limit(limit + 1).all()
    has_more = len(rows) > limit
    rows = rows[:limit]
    return api_ok({"checkouts": [serialize_loan(row) for row in rows], "next_cursor": _encode_cursor({"before": rows[-1].id}) if has_more and rows else None})


@bp.post("/checkouts")
@require_login
@require_entitlement(ENT_SHOP_ACCESS)
def create_checkout():
    user = current_auth_user()
    body = _json()
    key = idempotency_key()
    if not key:
        raise Validation("Idempotency-Key header (or idempotency_key field) is required.", field="idempotency_key", code="idempotency_required")
    item = None
    tag = body.get("tag") or value_from_request("tag")
    if tag:
        item, _via = inventory.find_item_by_tag(clean_tag_value(tag))
    if item is None:
        item = db.session.get(Item, int(body.get("item_id") or value_from_request("item_id") or 0))
    if not item:
        raise NotFound("Item not found.", code="item_not_found")
    signed_off_by = None
    signoff_id = body.get("signed_off_by_user_id") or value_from_request("signed_off_by_user_id")
    if signoff_id:
        signed_off_by = db.session.get(User, int(signoff_id))
        if not signed_off_by or not role_allows(signed_off_by.role, "team_leader"):
            raise Validation("signed_off_by_user_id must be a team lead or admin.", field="signed_off_by_user_id")
    loan = inventory.checkout(
        item=item,
        user=user,
        member=current_user_member(user),
        qty=parse_positive_int(body.get("qty", value_from_request("qty")), default=1),
        notes=(body.get("notes") or value_from_request("notes") or "").strip() or None,
        due_date=parse_due_date(body.get("due_date") or value_from_request("due_date")),
        idempotency_key=key,
        signed_off_by=signed_off_by,
        source="api",
    )
    return api_ok({"checkout": serialize_loan(loan)}, status=201)


@bp.post("/checkouts/<int:loan_id>/return")
@require_login
@require_entitlement(ENT_SHOP_ACCESS)
def return_checkout(loan_id):
    user = current_auth_user()
    loan = db.session.get(Loan, loan_id)
    if not loan:
        raise NotFound("Checkout not found.", code="loan_not_found")
    member = current_user_member(user)
    owns = loan.user_id == user.id or (member and loan.member_id == member.id)
    if not owns and not role_allows(user.role, "team_leader"):
        raise Forbidden("That checkout belongs to someone else.")
    body = _json()
    inventory.return_loan(
        loan=loan,
        qty=body.get("qty", value_from_request("qty")),
        condition=(body.get("condition") or value_from_request("condition") or "good").strip(),
        notes=(body.get("notes") or value_from_request("notes") or "").strip() or None,
        user=user,
        source="api",
    )
    return api_ok({"checkout": serialize_loan(loan)})


# --------------------------------------------------------------------------- prints


@bp.get("/print-requests")
@require_login
@require_entitlement(ENT_PORTAL_ACCESS)
def list_print_requests():
    user = current_auth_user()
    query = PrintRequest.query
    if not role_allows(user.role, "team_leader") or request.args.get("mine"):
        query = query.filter(PrintRequest.user_id == user.id)
    status = (request.args.get("status") or "").strip().lower()
    if status:
        query = query.filter(PrintRequest.status == status)
    limit = _limit()
    cursor = _decode_cursor(request.args.get("cursor"))
    if cursor:
        query = query.filter(PrintRequest.id < int(cursor.get("before", 0)))
    rows = query.order_by(PrintRequest.id.desc()).limit(limit + 1).all()
    has_more = len(rows) > limit
    rows = rows[:limit]
    return api_ok({"print_requests": [serialize_print_request(row) for row in rows], "next_cursor": _encode_cursor({"before": rows[-1].id}) if has_more and rows else None})


@bp.post("/print-requests")
@require_login
@require_entitlement(ENT_PRINT_SUBMIT)
def create_print_request():
    user = current_auth_user()
    form = request.form if request.form else _json()
    row = fabrication.submit_print_request(user, form, file_upload=request.files.get("print_file"), member=current_user_member(user))
    return api_ok({"print_request": serialize_print_request(row)}, status=201)


@bp.patch("/print-requests/<int:request_id>")
@require_role("team_leader")
@require_entitlement(ENT_PRINT_APPROVE)
def update_print_request(request_id):
    row = db.session.get(PrintRequest, request_id)
    if not row:
        raise NotFound("Print request not found.")
    body = _json()
    fabrication.update_print_request(row, body.get("status") or row.status, body.get("printer_type"), body.get("admin_notes"), current_auth_user())
    return api_ok({"print_request": serialize_print_request(row)})


# --------------------------------------------------------------------------- attendance


@bp.post("/checkins")
@require_login
@require_entitlement(ENT_EVENT_CHECKIN)
def create_checkin():
    """Tag tap. Resolves the event server-side: explicit id, else running/soon event, else open-shop."""
    user = current_auth_user()
    body = _json()
    tag_uid = clean_tag_value(body.get("tag_uid") or value_from_request("tag_uid"))
    target = user
    if tag_uid and role_allows(user.role, "team_leader"):
        resolved = identity.resolve_user_from_tag_uid(tag_uid)
        if resolved:
            target = resolved
    event_id = int(body.get("event_id") or value_from_request("event_id") or 0)
    event = db.session.get(Event, event_id) if event_id else None
    if event is None:
        candidates = attendance.candidate_events()
        if len(candidates) == 1:
            event = candidates[0]
        elif len(candidates) > 1:
            return jsonify({"ok": False, "code": "choose_event", "error": "Several meetings are running; pass event_id.", "events": [serialize_event(e) for e in candidates]}), 409
        else:
            event = attendance.ensure_open_shop_event(created_by=user)
            db.session.commit()
    row, created = attendance.check_user_into_event(target, event, method="api" if not tag_uid else "nfc", tag_uid=tag_uid)
    return api_ok({"checkin": serialize_attendance_record(row), "event": serialize_event(event), "created": created}, status=201 if created else 200)


@bp.get("/events")
@require_login
def list_events():
    rows = scheduling.upcoming_events(limit=_limit())
    return api_ok({"events": [serialize_event(e, scheduling.sync_status_for_event(e.id)) for e in rows]})


# --------------------------------------------------------------------------- scheduling


@bp.get("/availability")
@require_role("team_leader")
@require_entitlement(ENT_ROOM_BOOKING)
def availability():
    duration = parse_positive_int(request.args.get("duration"), default=60)
    days = parse_positive_int(request.args.get("days"), default=14)
    room = scheduling.normalize_meeting_room(request.args.get("room"))
    slots, error = scheduling.compute_available_slots(duration, days, room)
    if error:
        return jsonify({"ok": False, "code": "calendar_unavailable", "error": error}), 503
    return api_ok({"slots": [{"room": s["room"], "start": s["start"].isoformat(), "end": s["end"].isoformat(), "token": s["token"]} for s in slots]})


@bp.post("/bookings")
@require_role("team_leader")
@require_entitlement(ENT_ROOM_BOOKING)
def create_booking():
    user = current_auth_user()
    body = _json()
    event = scheduling.book_slot(user, body.get("team_name") or value_from_request("team_name"), body.get("slot_token") or value_from_request("slot_token"), notes=body.get("notes") or value_from_request("notes") or "")
    return api_ok({"event": serialize_event(event, scheduling.sync_status_for_event(event.id))}, status=201)


# --------------------------------------------------------------------------- roster


@bp.post("/roster/import")
@require_role("admin")
def roster_import():
    roster_file = request.files.get("roster_file")
    if not roster_file or not roster_file.filename or not roster_file.filename.lower().endswith(".pdf"):
        raise Validation("Upload a roster PDF as 'roster_file'.", field="roster_file")
    entries = roster.parse_roster_pdf_entries(roster_file.read())
    if not entries:
        raise Validation("No roster members were found in the uploaded file.")
    result = roster.import_roster_entries(entries, actor=current_auth_user())
    from asme.services import audit

    audit.record("bulk_roster_import", f"entries={len(entries)} via=api", actor=current_auth_user())
    db.session.commit()
    roster.emit_import_events(result)
    return api_ok(
        {
            "entries": len(entries),
            "created_members": result["created_members"],
            "created_users": result["created_users"],
            "updated_users": result["updated_users"],
            "credentials": result["credentials"],
        },
        status=201,
    )
