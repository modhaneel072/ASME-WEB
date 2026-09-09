"""Legacy ops pages (``ASME_ENABLE_LEGACY_OPS=1``). Off by default: every route
redirects into its portal equivalent."""

from __future__ import annotations

import os
from datetime import date, datetime, time
from mimetypes import guess_type
from pathlib import Path

from flask import Blueprint, current_app, flash, redirect, render_template, request, send_file, session, url_for
from sqlalchemy import func

from asme.auth.session import current_auth_user, get_active_member, is_admin_member
from asme.blueprints._context import frontend_portal_context, legacy_dashboard_context, render_ops_page
from asme.blueprints._helpers import flash_error
from asme.config import settings
from asme.constants import MEETING_ROOMS
from asme.extensions import db
from asme.integrations.calendar.outlook import normalize_outlook_embed_url
from asme.models import AttendanceScan, Item, ItemTag, Meeting, Member, PrintJob, Transaction
from asme.services import attendance, fabrication, identity, inventory, notifications, scheduling
from asme.services.errors import ServiceError
from asme.utils import parse_due_date, parse_int
from asme.utils.http import safe_next_url

bp = Blueprint("legacy_ops", __name__)

PAGE_ENDPOINTS = {
    "dashboard": "legacy_ops.dashboard_page",
    "attendance": "legacy_ops.attendance_page",
    "inventory": "legacy_ops.inventory_page",
    "prints": "legacy_ops.prints_page",
    "activity": "legacy_ops.activity_page",
    "calendar": "legacy_ops.calendar_page",
    "scan": "legacy_ops.scan_page",
    "my_items": "legacy_ops.my_items_page",
    "admin_nfc": "legacy_ops.admin_nfc_page",
}


def legacy_enabled():
    return settings().legacy_ops_enabled


def redirect_home(page):
    next_url = safe_next_url(request.form.get("next"))
    if next_url:
        return redirect(next_url)
    return redirect(url_for(PAGE_ENDPOINTS.get(page, "legacy_ops.dashboard_page")))


def _gate(portal_endpoint):
    if not legacy_enabled():
        return redirect(url_for(portal_endpoint))
    return None


@bp.get("/legacy/app")
def app_frontend():
    if not legacy_enabled():
        return redirect("/kiosk")
    return render_template("front/portal.html", **frontend_portal_context())


@bp.get("/dashboard")
def dashboard_page():
    return _gate("portal.portal_router") or render_ops_page(
        "ops/dashboard.html", "dashboard", "Dashboard", "Overview of attendance, stock, and printer queue status.", transaction_limit=12
    )


@bp.get("/attendance")
def attendance_page():
    return _gate("admin.attendance_page") or render_ops_page(
        "ops/attendance.html", "attendance", "Attendance", "Scan member NFC UIDs and track who is present today.", transaction_limit=10
    )


@bp.get("/inventory")
def inventory_page():
    return _gate("portal.member_inventory") or render_ops_page(
        "ops/inventory.html", "inventory", "Inventory", "Checkout, return, and monitor stock health across all items.", transaction_limit=12
    )


@bp.get("/prints")
def prints_page():
    return _gate("portal.member_prints") or render_ops_page(
        "ops/prints.html", "prints", "Printing", "Submit jobs and run separate H2S and P1S print queues.", transaction_limit=12
    )


@bp.get("/activity")
def activity_page():
    return _gate("admin.inventory_page") or render_ops_page(
        "ops/activity.html", "activity", "Activity Feed", "Recent checkout and return history.", transaction_limit=80
    )


@bp.get("/settings")
def settings_page():
    return _gate("admin.settings_page") or render_ops_page(
        "ops/settings.html", "settings", "Settings", "Customize theme and layout options for this device.", transaction_limit=10
    )


@bp.post("/session/member")
def set_active_member_route():
    member = identity.resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    next_url = safe_next_url(request.form.get("next"))
    if not member:
        flash("Could not sign in. Select or scan a valid member.", "error")
        return redirect(next_url or url_for("legacy_ops.scan_page"))
    session["active_member_id"] = member.id
    flash(f"Signed in as {member.name}.", "success")
    return redirect(next_url or url_for("legacy_ops.scan_page"))


@bp.post("/session/member/clear")
def clear_active_member_route():
    next_url = safe_next_url(request.form.get("next"))
    session.pop("active_member_id", None)
    flash("Signed out.", "info")
    return redirect(next_url or url_for("legacy_ops.scan_page"))


@bp.get("/scan")
def scan_page():
    if not legacy_enabled():
        return redirect("/kiosk")
    context = legacy_dashboard_context(transaction_limit=25)
    active_member = get_active_member()
    context.update(
        active_page="scan",
        page_title="NFC Scanner",
        page_subtitle="Scan item tags to check out or return equipment.",
        active_member=active_member,
        active_member_is_admin=is_admin_member(active_member),
    )
    return render_template("ops/scan.html", **context)


@bp.get("/my-items")
def my_items_page():
    gate = _gate("portal.member_checkouts")
    if gate:
        return gate
    active_member = get_active_member()
    context = legacy_dashboard_context(transaction_limit=25)
    context.update(
        active_page="my_items",
        page_title="My Checked Out Items",
        page_subtitle="Your currently checked-out inventory and checkout timestamps.",
        active_member=active_member,
        active_member_is_admin=is_admin_member(active_member),
        open_checkouts=inventory.open_loans_for(member=active_member) if active_member else [],
    )
    return render_template("ops/my_items.html", **context)


@bp.get("/admin/nfc")
def admin_nfc_page():
    gate = _gate("admin.nfc_page")
    if gate:
        return gate
    active_member = get_active_member()
    if not is_admin_member(active_member):
        flash("Admin access required for NFC registration and transaction log.", "error")
        return redirect(url_for("legacy_ops.scan_page"))
    item_id_raw = (request.args.get("item_id") or "").strip()
    member_id_raw = (request.args.get("member_id") or "").strip()
    date_from_raw = (request.args.get("date_from") or "").strip()
    date_to_raw = (request.args.get("date_to") or "").strip()
    query = Transaction.query.join(Item, Transaction.item_id == Item.id)
    if item_id_raw.isdigit():
        query = query.filter(Transaction.item_id == int(item_id_raw))
    if member_id_raw.isdigit():
        query = query.filter(Transaction.member_id == int(member_id_raw))
    date_from = parse_due_date(date_from_raw)
    if date_from:
        query = query.filter(Transaction.timestamp >= datetime.combine(date_from, time.min))
    date_to = parse_due_date(date_to_raw)
    if date_to:
        query = query.filter(Transaction.timestamp <= datetime.combine(date_to, time.max))
    context = legacy_dashboard_context(transaction_limit=25)
    context.update(
        active_page="admin_nfc",
        page_title="NFC Admin",
        page_subtitle="Register tags, review transaction history, and correct inventory values.",
        active_member=active_member,
        active_member_is_admin=True,
        tag_rows=ItemTag.query.order_by(ItemTag.created_at.desc(), ItemTag.id.desc()).limit(250).all(),
        tx_rows=query.order_by(Transaction.timestamp.desc(), Transaction.id.desc()).limit(250).all(),
        filter_item_id=item_id_raw,
        filter_member_id=member_id_raw,
        filter_date_from=date_from_raw,
        filter_date_to=date_to_raw,
    )
    return render_template("ops/admin_nfc.html", **context)


@bp.post("/admin/nfc/register")
def admin_register_item_tag():
    gate = _gate("admin.nfc_page")
    if gate:
        return gate
    if not is_admin_member(get_active_member()):
        flash("Admin access required.", "error")
        return redirect(url_for("legacy_ops.scan_page"))
    item_id_raw = (request.form.get("item_id") or "").strip()
    if not item_id_raw.isdigit():
        flash("Select a valid item.", "error")
        return redirect(url_for("legacy_ops.admin_nfc_page"))
    item = db.session.get(Item, int(item_id_raw))
    if not item:
        flash("Item not found.", "error")
        return redirect(url_for("legacy_ops.admin_nfc_page"))
    try:
        inventory.register_item_tag(item, request.form.get("tag_value"), source=(request.form.get("source") or "manual").strip().lower() or "manual")
    except ServiceError as exc:
        flash_error(exc)
        return redirect(url_for("legacy_ops.admin_nfc_page"))
    flash(f"Registered tag for {item.name}.", "success")
    return redirect(url_for("legacy_ops.admin_nfc_page"))


@bp.post("/admin/inventory/correct")
def admin_inventory_correct():
    gate = _gate("admin.inventory_page")
    if gate:
        return gate
    active_member = get_active_member()
    if not is_admin_member(active_member):
        flash("Admin access required.", "error")
        return redirect(url_for("legacy_ops.scan_page"))
    item_id_raw = (request.form.get("item_id") or "").strip()
    if not item_id_raw.isdigit():
        flash("Select a valid item for correction.", "error")
        return redirect(url_for("legacy_ops.admin_nfc_page"))
    item = db.session.get(Item, int(item_id_raw))
    if not item:
        flash("Item not found.", "error")
        return redirect(url_for("legacy_ops.admin_nfc_page"))
    actor = current_auth_user() or identity.user_for_member(active_member)
    inventory.adjust_counts(
        item,
        parse_int(request.form.get("total_qty"), default=0),
        parse_int(request.form.get("available_qty"), default=0),
        (request.form.get("note") or "").strip() or "Admin inventory correction",
        actor,
    )
    flash(f"Inventory corrected for {item.name}.", "success")
    return redirect(url_for("legacy_ops.admin_nfc_page"))


# --------------------------------------------------------------------------- legacy calendar


@bp.get("/calendar")
def calendar_page():
    gate = _gate("portal.member_calendar")
    if gate:
        return gate
    cfg = settings()
    active_member = get_active_member()
    meetings_by_room = {room: [] for room in MEETING_ROOMS}
    for meeting in scheduling.upcoming_meetings():
        meetings_by_room.setdefault(meeting.room, []).append(meeting)
    embed_url = normalize_outlook_embed_url(cfg.outlook_calendar_embed_url)
    from asme.integrations.calendar.outlook import OutlookCalendarProvider

    outlook = OutlookCalendarProvider(cfg)
    context = legacy_dashboard_context(transaction_limit=20)
    context.update(
        active_page="calendar",
        page_title="Calendar",
        page_subtitle="Schedule team meetings in the Robotics Room or Fluids Lab.",
        active_member=active_member,
        active_member_is_admin=is_admin_member(active_member),
        meeting_rooms=MEETING_ROOMS,
        meetings_by_room=meetings_by_room,
        calendar_default_date=str(date.today()),
        outlook_calendar_embed_url=embed_url,
        outlook_calendar_open_url=embed_url,
        outlook_calendar_placeholder=not embed_url,
        calendar_automation_status={
            "tenant_id_set": bool(cfg.outlook_tenant_id),
            "client_id_set": bool(cfg.outlook_client_id),
            "client_secret_set": bool(cfg.outlook_client_secret),
            "calendar_user": cfg.outlook_calendar_user,
            "default_calendar_id": cfg.outlook_calendar_id,
            "robotics_calendar_id": cfg.outlook_calendar_robotics_id,
            "fluids_calendar_id": cfg.outlook_calendar_fluids_id,
            "has_any_calendar_id": bool(cfg.outlook_calendar_id or cfg.outlook_calendar_robotics_id or cfg.outlook_calendar_fluids_id),
            "smtp_ready": bool(cfg.smtp_user and cfg.smtp_pass),
            "smtp_user": cfg.smtp_user,
            "cancel_notify_to": cfg.cancel_notify_to,
            "timezone": cfg.outlook_calendar_tz,
            "sync_enabled": outlook.status().enabled,
        },
        pending_cancellations=Meeting.query.filter(Meeting.cancel_request_token.isnot(None)).order_by(Meeting.cancel_requested_at.desc(), Meeting.id.desc()).all(),
    )
    return render_template("ops/calendar.html", **context)


@bp.post("/calendar/book")
def book_meeting():
    gate = _gate("portal.member_calendar")
    if gate:
        return gate
    try:
        meeting, calendar_error, sync_configured, config_has_inputs = scheduling.legacy_book_meeting(request.form)
    except ServiceError as exc:
        flash_error(exc)
        return redirect_home("calendar")
    when = f"{meeting.meeting_date.isoformat()} ({meeting.start_time.strftime('%H:%M')} - {meeting.end_time.strftime('%H:%M')})"
    if calendar_error:
        flash(f"Booked {meeting.room} for {meeting.team_name} on {when}. Outlook sync failed: {calendar_error}", "info")
    elif sync_configured:
        flash(f"Booked {meeting.room} for {meeting.team_name} on {when} and synced to Outlook Calendar.", "success")
    elif config_has_inputs:
        flash(f"Booked {meeting.room} for {meeting.team_name} on {when}. Outlook sync is not active yet. Run: python scripts/outlook_sync_doctor.py", "info")
    else:
        flash(f"Booked {meeting.room} for {meeting.team_name} on {when}.", "success")
    return redirect_home("calendar")


@bp.post("/calendar/meeting/<int:meeting_id>/cancel")
def request_meeting_cancel(meeting_id):
    gate = _gate("portal.member_calendar")
    if gate:
        return gate
    meeting = db.session.get(Meeting, meeting_id)
    if not meeting:
        flash("Meeting not found.", "error")
        return redirect_home("calendar")
    try:
        token = scheduling.legacy_request_cancel(meeting)
    except ServiceError as exc:
        flash(exc.message, "info")
        return redirect_home("calendar")
    confirm_url = url_for("legacy_ops.confirm_meeting_cancel", token=token, _external=True)
    reject_url = url_for("legacy_ops.reject_meeting_cancel", token=token, _external=True)
    email_error = notifications.send_meeting_cancel_confirmation_email(meeting, confirm_url=confirm_url, reject_url=reject_url)
    if email_error:
        scheduling.legacy_clear_cancel(meeting)
        flash(f"Cancellation email could not be sent: {email_error}", "error")
        return redirect_home("calendar")
    flash("Cancellation request sent. An approval email was sent for confirmation.", "info")
    return redirect_home("calendar")


@bp.get("/calendar/cancel/confirm/<token>")
def confirm_meeting_cancel(token):
    gate = _gate("portal.member_calendar")
    if gate:
        return gate
    meeting = scheduling.find_meeting_by_cancel_token(token)
    if not meeting:
        return "<h3>Cancellation link is invalid or already used.</h3><p>You can close this tab.</p>", 404
    team_name, room, meeting_date = meeting.team_name, meeting.room, meeting.meeting_date.isoformat()
    error = scheduling.legacy_confirm_cancel(meeting)
    if error:
        return f"<h3>Cancellation could not be completed.</h3><p>{error}</p><p>Please fix configuration and retry.</p>", 500
    return (
        "<h3>Cancellation confirmed.</h3>"
        f"<p>{team_name} in {room} on {meeting_date} was removed from Outlook Calendar.</p>"
        "<p>You can close this tab.</p>"
    )


@bp.get("/calendar/cancel/reject/<token>")
def reject_meeting_cancel(token):
    gate = _gate("portal.member_calendar")
    if gate:
        return gate
    meeting = scheduling.find_meeting_by_cancel_token(token)
    if not meeting:
        return "<h3>Rejection link is invalid or already used.</h3><p>You can close this tab.</p>", 404
    scheduling.legacy_clear_cancel(meeting)
    return (
        "<h3>Cancellation request rejected.</h3>"
        "<p>The meeting remains on the schedule and in Outlook Calendar.</p>"
        "<p>You can close this tab.</p>"
    )


# --------------------------------------------------------------------------- legacy forms


@bp.post("/attendance/scan")
def attendance_scan():
    try:
        member, first_today = attendance.legacy_scan(request.form.get("uid"))
    except ServiceError as exc:
        flash(exc.message, "error")
        return redirect_home("attendance")
    if first_today:
        flash(f"Attendance marked for {member.name}.", "success")
    else:
        flash(f"{member.name} scanned again. Attendance already marked for today.", "info")
    return redirect_home("attendance")


@bp.post("/transact")
def transact():
    auth_user = current_auth_user()
    member = identity.resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    item = inventory.resolve_item(request.form.get("item_tag"), request.form.get("item_id"))
    action = (request.form.get("action") or "").strip().lower()
    qty = parse_int(request.form.get("qty"), default=1)
    if not member:
        flash("Could not find member. Scan a member UID or select one.", "error")
        return redirect_home("inventory")
    if not item:
        flash("Could not find item. Scan an item UID or select one.", "error")
        return redirect_home("inventory")
    if action not in {"checkout", "return"}:
        flash("Invalid inventory action.", "error")
        return redirect_home("inventory")
    try:
        inventory.legacy_transact(
            member=member,
            item=item,
            action=action,
            qty=qty,
            notes=(request.form.get("notes") or "").strip() or None,
            due_date=parse_due_date(request.form.get("due_date")),
            auth_user=auth_user,
        )
    except ServiceError as exc:
        flash_error(exc)
        return redirect_home("inventory")
    flash(f"{action.title()} saved: {qty} x {item.name} for {member.name}.", "success")
    return redirect_home("inventory")


@bp.post("/print/submit")
def submit_print_job():
    member = identity.resolve_member(request.form.get("member_tag"), request.form.get("member_id"))
    try:
        job, started = fabrication.submit_print_job(
            member, request.form.get("printer_type"), request.files.get("gcode_file"), notes=(request.form.get("notes") or "").strip() or None
        )
    except ServiceError as exc:
        flash_error(exc)
        return redirect_home("prints")
    flash(f"{job.printer_type} job submitted and auto-started." if started else f"{job.printer_type} job submitted to queue.", "success")
    return redirect_home("prints")


@bp.post("/print/job/<int:job_id>/complete")
def complete_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
        return redirect_home("prints")
    fabrication.finish_print_job(job, "done")
    flash(f"Marked job #{job.id} done. Next {job.printer_type} job auto-started if available.", "success")
    return redirect_home("prints")


@bp.post("/print/job/<int:job_id>/fail")
def fail_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
        return redirect_home("prints")
    fabrication.finish_print_job(job, "failed")
    flash(f"Marked job #{job.id} failed. Next {job.printer_type} job auto-started if available.", "info")
    return redirect_home("prints")


@bp.get("/print/job/<int:job_id>/download")
def download_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job or not os.path.exists(job.file_path):
        flash("Print file not found for that job.", "error")
        return redirect_home("prints")
    return send_file(job.file_path, as_attachment=True, download_name=job.file_name)


@bp.get("/print/job/<int:job_id>/open")
def open_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job or not os.path.exists(job.file_path):
        flash("Print file not found for that job.", "error")
        return redirect_home("prints")
    mime_type, _ = guess_type(job.file_name)
    return send_file(job.file_path, as_attachment=False, download_name=job.file_name, mimetype=mime_type or "application/octet-stream")


@bp.post("/print/job/<int:job_id>/delete")
def delete_print_job(job_id):
    job = db.session.get(PrintJob, job_id)
    if not job:
        flash("Print job not found.", "error")
        return redirect_home("prints")
    try:
        result = fabrication.delete_print_job_with_file(job)
    except ServiceError as exc:
        flash_error(exc)
        return redirect_home("prints")
    if result["file_error"]:
        flash(f"Deleted job, but file removal failed: {result['file_error'][:200]}", "error")
    elif result["file_removed"]:
        flash(f"Deleted print job and removed {result['file_name']}.", "info")
    else:
        flash(f"Deleted print job for {result['file_name']}. File was already missing.", "info")
    return redirect_home("prints")


@bp.get("/export")
def export_excel():
    try:
        import pandas as pd
    except Exception:
        flash("Excel export is unavailable because pandas is not loading correctly.", "error")
        return redirect(url_for("legacy_ops.dashboard_page"))

    members = Member.query.order_by(Member.id.asc()).all()
    items = Item.query.order_by(Item.id.asc()).all()
    transactions = Transaction.query.order_by(Transaction.timestamp.desc()).all()
    scans = AttendanceScan.query.order_by(AttendanceScan.scanned_at.desc()).all()
    jobs = PrintJob.query.order_by(PrintJob.submitted_at.desc()).all()
    meetings = Meeting.query.order_by(Meeting.meeting_date.asc(), Meeting.start_time.asc(), Meeting.id.asc()).all()
    item_tags = ItemTag.query.order_by(ItemTag.id.asc()).all()

    frames = {
        "Members": pd.DataFrame([{"id": m.id, "name": m.name, "email": m.email, "class": m.member_class, "nfc_tag": m.nfc_tag, "created_at": m.created_at} for m in members]),
        "Items": pd.DataFrame(
            [
                {"id": i.id, "name": i.name, "description": i.description, "category": i.category, "location": i.location, "total_qty": i.total_qty, "available_qty": i.available_qty, "nfc_tag": i.nfc_tag, "created_at": i.created_at}
                for i in items
            ]
        ),
        "Transactions": pd.DataFrame(
            [
                {
                    "id": t.id, "timestamp": t.timestamp, "member": t.borrower_name, "member_email": (t.member.email if t.member else (t.user.email if t.user else "")),
                    "item": t.item.name if t.item else "", "action": t.action, "status": t.status, "qty": t.qty, "checkout_time": t.checkout_time,
                    "return_time": t.return_time, "due_date": t.due_date, "notes": t.notes, "checkout_notes": t.checkout_notes,
                    "return_condition": t.return_condition, "return_notes": t.return_notes, "return_photo_path": t.return_photo_path,
                }
                for t in transactions
            ]
        ),
        "ItemTags": pd.DataFrame([{"id": tag.id, "item_id": tag.item_id, "item_name": tag.item.name if tag.item else "", "tag_value": tag.tag_value, "source": tag.source, "created_at": tag.created_at} for tag in item_tags]),
        "Attendance": pd.DataFrame([{"id": s.id, "member": s.member.name, "member_email": s.member.email, "uid": s.scanned_uid, "attendance_date": s.attendance_date, "scanned_at": s.scanned_at} for s in scans]),
        "PrintJobs": pd.DataFrame(
            [
                {"id": j.id, "member": j.member.name, "member_email": j.member.email, "printer_type": j.printer_type, "file_name": j.file_name, "file_path": j.file_path, "status": j.status, "notes": j.notes, "submitted_at": j.submitted_at, "started_at": j.started_at, "completed_at": j.completed_at}
                for j in jobs
            ]
        ),
        "Meetings": pd.DataFrame(
            [
                {
                    "id": m.id, "team_name": m.team_name, "requester_email": m.requester_email, "room": m.room, "meeting_date": m.meeting_date, "start_time": m.start_time, "end_time": m.end_time,
                    "notes": m.notes, "google_event_id": m.google_event_id, "google_calendar_id": m.google_calendar_id, "outlook_event_id": m.outlook_event_id, "outlook_calendar_id": m.outlook_calendar_id,
                    "cancel_request_token": m.cancel_request_token, "cancel_requested_at": m.cancel_requested_at, "created_at": m.created_at,
                }
                for m in meetings
            ]
        ),
    }
    export_path = Path(current_app.instance_path) / "inventory_export.xlsx"
    with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
        for sheet, frame in frames.items():
            frame.to_excel(writer, index=False, sheet_name=sheet)
    return send_file(export_path, as_attachment=True, download_name="inventory_export.xlsx")
