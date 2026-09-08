"""Legacy ``/api/*`` JSON surface used by the kiosk front-end and old scripts.

Kept byte-compatible in shape; every write now goes through the services so
there is one copy of the rules. New clients should use ``/api/v1``.
"""

from __future__ import annotations

from datetime import date, datetime
from uuid import uuid4

from flask import Blueprint, current_app, jsonify, request, session
from werkzeug.utils import secure_filename

from asme.auth.session import current_auth_user, current_user_member, get_active_member, is_admin_member, normalize_role, role_allows
from asme.constants import ALLOWED_RETURN_PHOTO_EXTENSIONS
from asme.extensions import db
from asme.models import Item, Member, Transaction
from asme.services import attendance, fabrication, identity, inventory
from asme.services.errors import ServiceError
from asme.services.serializers import (
    serialize_attendance_scan,
    serialize_item,
    serialize_member,
    serialize_open_checkout,
    serialize_print_job,
    serialize_transaction,
    serialize_user,
)
from asme.utils import default_due_date, iso_or_none, parse_due_date, parse_int
from asme.utils.http import api_error, value_from_request

bp = Blueprint("api_legacy", __name__, url_prefix="/api")


@bp.errorhandler(ServiceError)
def _service_error(exc):
    return jsonify(exc.to_dict()), exc.status


def _require_active_member():
    member = get_active_member()
    if not member:
        user = current_auth_user()
        if user:
            linked = current_user_member(user)
            if linked:
                session["active_member_id"] = linked.id
                member = linked
    if member:
        return member, None
    return None, (jsonify({"ok": False, "error": "Login required. Select your member profile first."}), 401)


def build_bootstrap_payload():
    members = Member.query.order_by(Member.name.asc()).all()
    items = Item.query.order_by(Item.name.asc()).all()
    scans = attendance.today_unique_scans()
    recent = Transaction.query.order_by(Transaction.timestamp.desc()).limit(15).all()
    queues = fabrication.queue_snapshot()
    return {
        "today": str(date.today()),
        "default_due": str(default_due_date()),
        "members": [serialize_member(member) for member in members],
        "items": [serialize_item(item) for item in items],
        "attendance_today": [serialize_attendance_scan(scan) for scan in scans],
        "attendance_count": len(scans),
        "recent_transactions": [serialize_transaction(tx) for tx in recent],
        "queues": {
            printer: {
                "active": serialize_print_job(data["active"]) if data["active"] else None,
                "queued": [serialize_print_job(job) for job in data["queued"]],
                "recent_finished": [serialize_print_job(job) for job in data["recent_finished"]],
            }
            for printer, data in queues.items()
        },
    }


def api_success(message=None, status=200):
    return jsonify({"ok": True, "message": message, "payload": build_bootstrap_payload()}), status


@bp.post("/ask")
def api_ask():
    """Admin-only: run a single conversation turn against the analytics assistant."""
    user = current_auth_user()
    if not user or normalize_role(user.role) != "admin":
        return jsonify({"error": "admin only"}), 403
    payload = request.get_json(silent=True) or {}
    message = (payload.get("message") or "").strip()
    history = payload.get("history") or []
    if not message:
        return jsonify({"error": "empty message"}), 400
    if len(message) > 2000:
        return jsonify({"error": "message too long"}), 400
    from asme.integrations.assistant import run_assistant

    result = run_assistant(message, history=history)
    return jsonify({"reply": result.get("reply", ""), "tool_calls": result.get("tool_calls", [])})


@bp.get("/session/me")
def api_session_me():
    user = current_auth_user()
    if user and not session.get("active_member_id"):
        linked = current_user_member(user)
        if linked:
            session["active_member_id"] = linked.id
    member = get_active_member()
    if not member:
        return jsonify({"ok": True, "payload": {"member": None, "is_admin": bool(user and role_allows(user.role, "admin")), "user": serialize_user(user)}})
    return jsonify({"ok": True, "payload": {"member": serialize_member(member), "is_admin": is_admin_member(member), "user": serialize_user(user)}})


@bp.post("/session/member")
def api_set_active_member():
    member = identity.resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
    if not member:
        return jsonify({"ok": False, "error": "Could not sign in. Select or scan a valid member."}), 404
    session["active_member_id"] = member.id
    user = current_auth_user()
    if user and not user.member_id:
        user.member_id = member.id
        db.session.commit()
    return jsonify(
        {
            "ok": True,
            "message": f"Signed in as {member.name}.",
            "payload": {"member": serialize_member(member), "is_admin": is_admin_member(member), "user": serialize_user(user)},
        }
    )


@bp.post("/session/clear")
def api_clear_active_member():
    session.pop("active_member_id", None)
    return jsonify({"ok": True, "message": "Signed out."})


@bp.get("/my-items")
def api_my_items():
    member, auth_error = _require_active_member()
    if auth_error:
        return auth_error
    open_checkouts = inventory.open_loans_for(member=member)
    return jsonify({"ok": True, "payload": {"member": serialize_member(member), "items": [serialize_open_checkout(tx) for tx in open_checkouts]}})


@bp.get("/bootstrap")
def api_bootstrap():
    return jsonify({"ok": True, "payload": build_bootstrap_payload()})


@bp.post("/attendance/scan")
def api_attendance_scan():
    uid = str(value_from_request("uid", "")).strip()
    try:
        member, first_today = attendance.legacy_scan(uid)
    except ServiceError as exc:
        return api_error(exc.message, status=exc.status if exc.status != 400 else 400)
    if first_today:
        message = f"Attendance marked for {member.name}."
    else:
        message = f"{member.name} scanned again. Attendance already marked for today."
    return api_success(message=message)


@bp.post("/inventory/transact")
def api_inventory_transact():
    auth_user = current_auth_user()
    member = identity.resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
    item = inventory.resolve_item(value_from_request("item_tag"), value_from_request("item_id"))
    action = str(value_from_request("action", "")).strip().lower()
    qty = parse_int(value_from_request("qty"), default=1)
    notes = str(value_from_request("notes", "")).strip() or None
    due = parse_due_date(value_from_request("due_date"))
    if not member:
        return api_error("Could not find member. Scan a member UID or select one.")
    if not item:
        return api_error("Could not find item. Scan an item UID or select one.")
    if action not in {"checkout", "return"}:
        return api_error("Invalid inventory action.")
    try:
        inventory.legacy_transact(member=member, item=item, action=action, qty=qty, notes=notes, due_date=due, auth_user=auth_user)
    except ServiceError as exc:
        return api_error(exc.message, status=400)
    return api_success(message=f"{action.title()} saved: {qty} x {item.name} for {member.name}.")


@bp.get("/items/by-tag")
def api_item_by_tag():
    member, auth_error = _require_active_member()
    if auth_error:
        return auth_error
    from asme.utils import clean_tag_value

    tag_value = clean_tag_value(request.args.get("tag"))
    if not tag_value:
        return jsonify({"ok": False, "error": "Tag value is required."}), 400
    item, resolved_via = inventory.find_item_by_tag(tag_value)
    if not item:
        return jsonify({"ok": False, "error": "Tag not registered to an inventory item."}), 404
    open_for_user = inventory.get_open_loan(item.id, member=member)
    checked_out_by_others = Transaction.query.filter(
        Transaction.item_id == item.id, Transaction.status == "OUT", Transaction.member_id != member.id
    ).count()
    return jsonify(
        {
            "ok": True,
            "payload": {
                "item": {
                    "id": item.id,
                    "name": item.name,
                    "description": item.description,
                    "category": item.category,
                    "location": item.location,
                    "available_qty": item.available_qty,
                    "total_qty": item.total_qty,
                    "resolved_via": resolved_via,
                },
                "active_member": {"id": member.id, "name": member.name, "email": member.email},
                "user_has_open_checkout": bool(open_for_user),
                "user_open_checkout": (
                    {
                        "transaction_id": open_for_user.id,
                        "qty": open_for_user.qty,
                        "checkout_time": iso_or_none(open_for_user.checkout_time or open_for_user.timestamp),
                        "checkout_notes": open_for_user.checkout_notes or open_for_user.notes,
                    }
                    if open_for_user
                    else None
                ),
                "checked_out_by_others_count": checked_out_by_others,
            },
        }
    )


@bp.post("/checkout")
def api_checkout_item():
    member, auth_error = _require_active_member()
    if auth_error:
        return auth_error
    auth_user = current_auth_user()
    try:
        item_id = int(value_from_request("item_id"))
    except Exception:
        return jsonify({"ok": False, "error": "item_id is required."}), 400
    item = db.session.get(Item, item_id)
    if not item:
        return jsonify({"ok": False, "error": "Item not found."}), 404
    try:
        loan = inventory.checkout(
            item=item,
            member=member,
            user=auth_user,
            qty=parse_int(value_from_request("qty"), default=1),
            notes=str(value_from_request("notes", "")).strip() or None,
            due_date=parse_due_date(value_from_request("due_date")),
            idempotency_key=str(value_from_request("idempotency_key", "") or request.headers.get("Idempotency-Key") or "").strip() or None,
            source="kiosk",
        )
    except ServiceError as exc:
        return jsonify({"ok": False, "error": exc.message, "code": exc.code}), exc.status
    db.session.refresh(item)
    return jsonify(
        {
            "ok": True,
            "message": f"Checked out {loan.qty} x {item.name}.",
            "payload": {"item_id": item.id, "available_qty": item.available_qty, "transaction_id": loan.id, "member_id": member.id},
        }
    )


@bp.post("/return")
def api_return_item():
    member, auth_error = _require_active_member()
    if auth_error:
        return auth_error
    auth_user = current_auth_user()
    try:
        item_id = int(value_from_request("item_id"))
    except Exception:
        return jsonify({"ok": False, "error": "item_id is required."}), 400
    return_condition = str(value_from_request("condition", "")).strip() or None
    if not return_condition:
        return jsonify({"ok": False, "error": "Return condition is required."}), 400
    item = db.session.get(Item, item_id)
    if not item:
        return jsonify({"ok": False, "error": "Item not found."}), 404
    open_tx = inventory.get_open_loan(item.id, member=member, user=auth_user)
    if not open_tx:
        return jsonify({"ok": False, "error": "No open checkout found for this item under your account."}), 409

    photo_path = None
    return_photo = request.files.get("photo")
    if return_photo and return_photo.filename:
        if "." not in return_photo.filename or return_photo.filename.rsplit(".", 1)[1].lower() not in ALLOWED_RETURN_PHOTO_EXTENSIONS:
            return jsonify({"ok": False, "error": "Photo must be .jpg, .jpeg, .png, or .webp."}), 400
        from pathlib import Path

        photo_dir = Path(current_app.instance_path) / "return_photos"
        photo_dir.mkdir(parents=True, exist_ok=True)
        stored_name = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{secure_filename(return_photo.filename)}"
        photo_file = photo_dir / stored_name
        return_photo.save(photo_file)
        photo_path = str(photo_file)

    try:
        inventory.return_loan(
            loan=open_tx,
            qty=parse_int(value_from_request("qty"), default=1),
            condition=return_condition,
            notes=str(value_from_request("notes", "")).strip() or None,
            photo_path=photo_path,
            user=auth_user,
            source="kiosk",
        )
    except ServiceError as exc:
        return jsonify({"ok": False, "error": exc.message, "code": exc.code}), exc.status
    db.session.refresh(item)
    return jsonify(
        {
            "ok": True,
            "message": f"Returned {open_tx.qty} x {item.name}.",
            "payload": {"item_id": item.id, "available_qty": item.available_qty, "transaction_id": open_tx.id, "member_id": member.id},
        }
    )


@bp.post("/print/submit")
def api_print_submit():
    member = identity.resolve_member(value_from_request("member_tag"), value_from_request("member_id"))
    try:
        job, started = fabrication.submit_print_job(
            member,
            str(value_from_request("printer_type", "")),
            request.files.get("gcode_file"),
            notes=str(value_from_request("notes", "")).strip() or None,
        )
    except ServiceError as exc:
        return api_error(exc.message)
    message = f"{job.printer_type} job submitted and auto-started." if started else f"{job.printer_type} job submitted to queue."
    return api_success(message=message, status=201)


@bp.post("/print/job/<int:job_id>/complete")
def api_complete_print_job(job_id):
    job = fabrication.get_job_or_404(job_id)
    fabrication.finish_print_job(job, "done")
    return api_success(message=f"Marked job #{job.id} done. Next {job.printer_type} job auto-started if available.")


@bp.post("/print/job/<int:job_id>/fail")
def api_fail_print_job(job_id):
    job = fabrication.get_job_or_404(job_id)
    fabrication.finish_print_job(job, "failed")
    return api_success(message=f"Marked job #{job.id} failed. Next {job.printer_type} job auto-started if available.")


@bp.post("/print/job/<int:job_id>/delete")
def api_delete_print_job(job_id):
    job = fabrication.get_job_or_404(job_id)
    result = fabrication.delete_print_job_with_file(job)
    if result["file_error"]:
        message = f"Deleted job for {result['file_name']}, but file removal failed: {result['file_error'][:200]}"
    elif result["file_removed"]:
        message = f"Deleted job and removed {result['file_name']} from storage."
    else:
        message = f"Deleted job for {result['file_name']}. File was already missing."
    return api_success(message=message)
