"""Admin portal (HTML)."""

from __future__ import annotations

from datetime import datetime, time, timedelta

from flask import Blueprint, Response, flash, redirect, render_template, request, send_file, session, url_for
from io import BytesIO
from pathlib import Path

from asme.auth.session import current_auth_user, normalize_role, require_role, role_allows
from asme.blueprints._context import admin_dashboard_context
from asme.blueprints._helpers import flash_error, on_service_error
from asme.config import settings
from asme.constants import ALL_ENTITLEMENTS, MEETING_ROOMS, PRINT_REQUEST_PRINTERS
from asme.extensions import db
from asme.jobs import outbox
from asme.models import (
    Announcement,
    AttendanceRecord,
    ContactMessage,
    Event,
    Item,
    Loan,
    NFCTag,
    OutboxJob,
    PrintRequest,
    StockDiscrepancy,
    TrainingModule,
    User,
)
from asme.services import attendance, bootstrap, content, fabrication, identity, inventory, roster, scheduling
from asme.services.errors import ServiceError
from asme.services.onboarding import engine, entitlements
from asme.utils import csv_stream_from_rows, parse_bool_flag, parse_non_negative_int, parse_positive_int
from asme.utils.http import redirect_to_next

bp = Blueprint("admin", __name__, url_prefix="/portal/admin")


def _page(template, title, active_page, **extra):
    context = admin_dashboard_context()
    context["page_title"] = title
    context["active_page"] = active_page
    context.update(extra)
    return render_template(template, **context)


# --------------------------------------------------------------------------- pages


@bp.get("")
@require_role("admin")
def dashboard():
    return _page("portal/admin_dashboard_home.html", "Admin Home", "admin_dashboard")


@bp.get("/members")
@require_role("admin")
def members_page():
    show_inactive = parse_bool_flag(request.args.get("show_inactive"), default=False)
    users_query = User.query.order_by(User.created_at.desc(), User.id.desc())
    if not show_inactive:
        users_query = users_query.filter(User.is_active.is_(True))
    return _page(
        "portal/admin_members.html",
        "Members / Roles",
        "admin_members",
        users=users_query.all(),
        show_inactive=show_inactive,
        inactive_user_count=User.query.filter(User.is_active.is_(False)).count(),
        latest_roster_credentials_file=(session.get("latest_roster_credentials_file") or "").strip(),
    )


@bp.get("/nfc")
@require_role("admin")
def nfc_page():
    return redirect(url_for("admin.dashboard"))


@bp.get("/attendance")
@require_role("admin")
def attendance_page():
    context = admin_dashboard_context()
    active_tab = (request.args.get("tab") or "live").strip().lower()
    if active_tab not in {"live", "history"}:
        active_tab = "live"
    checkin_event_id = parse_positive_int(request.args.get("checkin_event_id") or request.args.get("event_id"), default=0)
    checkin_event = db.session.get(Event, checkin_event_id) if checkin_event_id else None
    if not checkin_event:
        checkin_event = (
            Event.query.filter(Event.end_time >= datetime.now(), Event.status != "cancelled")
            .order_by(Event.start_time.asc(), Event.id.asc())
            .first()
        )
    if not checkin_event and context["events"]:
        checkin_event = context["events"][0]
    history_event_id = parse_positive_int(request.args.get("history_event_id"), default=0)
    history_date_raw = (request.args.get("history_date") or "").strip()
    history_date = None
    if history_date_raw:
        try:
            history_date = datetime.strptime(history_date_raw, "%Y-%m-%d").date()
        except Exception:
            history_date = None
    history_query = AttendanceRecord.query
    if history_event_id:
        history_query = history_query.filter(AttendanceRecord.event_id == history_event_id)
    if history_date:
        history_start = datetime.combine(history_date, time.min)
        history_query = history_query.filter(
            AttendanceRecord.checkin_time >= history_start, AttendanceRecord.checkin_time < history_start + timedelta(days=1)
        )
    context.update(
        page_title="Attendance",
        active_page="admin_attendance",
        active_tab=active_tab,
        checkin_event=checkin_event,
        history_records=history_query.order_by(AttendanceRecord.checkin_time.desc(), AttendanceRecord.id.desc()).limit(350).all(),
        history_event_id=history_event_id,
        history_date=history_date_raw,
    )
    return render_template("portal/admin_attendance.html", **context)


@bp.get("/inventory")
@require_role("admin")
def inventory_page():
    return _page(
        "portal/admin_inventory.html",
        "Inventory Admin",
        "admin_inventory",
        low_stock_items=inventory.low_stock_items(),
        overdue_transactions=inventory.overdue_loans(),
        stock_discrepancies=StockDiscrepancy.query.filter_by(status="open").order_by(StockDiscrepancy.id.desc()).all(),
    )


@bp.get("/prints")
@require_role("admin")
def prints_page():
    context = admin_dashboard_context()
    grouped = {printer: [] for printer in PRINT_REQUEST_PRINTERS}
    for row in context.get("print_requests", []):
        grouped.setdefault(row.printer_type, []).append(row)
    context.update(page_title="Prints Admin", active_page="admin_prints", print_requests_by_printer=grouped)
    return render_template("portal/admin_prints.html", **context)


@bp.route("/announcements", methods=["GET", "POST"])
@require_role("admin")
def announcements_page():
    return redirect(url_for("admin.dashboard"))


@bp.route("/content", methods=["GET", "POST"])
@require_role("admin")
def content_page():
    return redirect(url_for("admin.dashboard"))


@bp.get("/calendar")
@require_role("admin")
def calendar_page():
    cfg = settings()
    context = admin_dashboard_context()
    status = scheduling.provider_status()
    work_start, work_end = scheduling.work_hours()
    requested_events = Event.query.filter(Event.status == "requested").order_by(Event.start_time.asc(), Event.id.asc()).limit(120).all()
    room_events = {}
    from sqlalchemy import func

    for room in MEETING_ROOMS:
        room_events[room] = (
            Event.query.filter(
                func.lower(func.coalesce(Event.location, "")) == room.lower(),
                Event.status.in_(["requested", "scheduled"]),
            )
            .order_by(Event.start_time.asc(), Event.id.asc())
            .limit(180)
            .all()
        )
    calendar_embed = scheduling.calendar_embed_url()
    env_status = [
        {"name": "ASME_CALENDAR_PROVIDER", "value": cfg.calendar_provider, "configured": bool(cfg.calendar_provider)},
    ]
    if cfg.calendar_provider == "google":
        env_status += [
            {"name": "GOOGLE_CALENDAR_ID_ROBOTICS", "value": "set" if status.room_ids.get("Robotics Room") else "missing", "configured": bool(status.room_ids.get("Robotics Room"))},
            {"name": "GOOGLE_CALENDAR_ID_FLUIDS", "value": "set" if status.room_ids.get("Fluids Lab") else "missing", "configured": bool(status.room_ids.get("Fluids Lab"))},
            {"name": "GOOGLE_SERVICE_ACCOUNT_JSON", "value": "set" if cfg.google_service_account_json else "missing", "configured": bool(cfg.google_service_account_json)},
        ]
    else:
        env_status += [
            {"name": "ASME_OUTLOOK_TENANT_ID", "value": "set" if cfg.outlook_tenant_id else "missing", "configured": bool(cfg.outlook_tenant_id)},
            {"name": "ASME_OUTLOOK_CLIENT_ID", "value": "set" if cfg.outlook_client_id else "missing", "configured": bool(cfg.outlook_client_id)},
            {"name": "ASME_OUTLOOK_CALENDAR_USER", "value": "set" if cfg.outlook_calendar_user else "missing", "configured": bool(cfg.outlook_calendar_user)},
        ]
    env_status.append({"name": "CALENDAR_EMBED_URL", "value": "set" if calendar_embed else "missing", "configured": bool(calendar_embed)})
    context.update(
        page_title="Calendar",
        active_page="admin_calendar",
        meeting_rooms=MEETING_ROOMS,
        calendar_provider=cfg.calendar_provider,
        calendar_embed_url=calendar_embed,
        google_schedule_enabled=status.enabled,
        calendar_work_start=work_start.strftime("%H:%M"),
        calendar_work_end=work_end.strftime("%H:%M"),
        calendar_days_default=min(max(cfg.calendar_scheduling_days, 1), 30),
        calendar_warnings=list(status.errors),
        google_env_status=env_status,
        requested_events=requested_events,
        room_events=room_events,
        calendar_sync_failures=OutboxJob.query.filter(OutboxJob.kind.like("calendar.%"), OutboxJob.status == "failed").order_by(OutboxJob.id.desc()).limit(20).all(),
    )
    return render_template("portal/admin_calendar.html", **context)


@bp.get("/exports")
@require_role("admin")
def exports_page():
    return redirect(url_for("admin.dashboard"))


@bp.get("/exports/download.zip")
@require_role("admin")
def exports_zip():
    return redirect(url_for("admin.dashboard"))


@bp.get("/settings")
@require_role("admin")
def settings_page():
    cfg = settings()
    status = scheduling.provider_status()
    embed_configured = bool((scheduling.calendar_embed_url() or "").strip())
    return _page(
        "portal/admin_settings.html",
        "System Status",
        "admin_settings",
        system_status={
            "google_calendar_configured": status.enabled or embed_configured,
            "printer_automation_configured": bool(cfg.h2s_print_cmd and cfg.p1s_print_cmd),
            "legacy_ops_enabled": cfg.legacy_ops_enabled,
            "onboarding_enforced": cfg.onboarding_enforce,
            "outbox_worker_enabled": cfg.outbox_worker_enabled,
            "schema_revision": bootstrap.current_revision(),
        },
        calendar_provider=cfg.calendar_provider,
        calendar_errors=status.errors,
        config_problems=cfg.validate() + cfg.warnings(),
        outbox_stats=outbox.stats(),
        outbox_failures=outbox.recent_failures(),
    )


# --------------------------------------------------------------------------- launchpad (admin)


def launchpad_admin_context(is_admin_view=True):
    chapter = engine.evaluate_chapter()
    overview = engine.member_overview()
    pending_signoffs = []
    for row in overview:
        phase = row.get("current_phase")
        if not phase:
            continue
        for task in phase["tasks"]:
            if task["human"] and task["status"] != "complete":
                pending_signoffs.append({"user": row["user"], "phase": phase, "task": task})
    unsigned_loans = (
        Loan.query.filter(Loan.signed_off_by_user_id.is_(None), Loan.status == "OUT")
        .order_by(Loan.checkout_time.desc(), Loan.id.desc())
        .limit(50)
        .all()
    )
    return {
        "is_admin_view": is_admin_view,
        "chapter": chapter,
        "chapter_track": chapter["tracks"][0] if chapter["tracks"] else None,
        "member_overview": overview,
        "pending_signoffs": pending_signoffs,
        "unsigned_loans": unsigned_loans,
        "training_modules": TrainingModule.query.order_by(TrainingModule.id.asc()).all(),
        "all_entitlements": ALL_ENTITLEMENTS,
        "onboarding_enforced": settings().onboarding_enforce,
        "outbox_stats": outbox.stats(),
    }


@bp.get("/launchpad")
@require_role("admin")
def launchpad_page():
    return _page("portal/admin_launchpad.html", "Launchpad", "admin_launchpad", **launchpad_admin_context(is_admin_view=True))


@bp.post("/launchpad/tasks/<task_key>/complete")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_complete_chapter_task(task_key):
    engine.complete_task(task_key, engine.subject_for_chapter(), current_auth_user(), note=request.form.get("note"))
    flash("Chapter task marked complete.", "success")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/tasks/<task_key>/reopen")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_reopen_chapter_task(task_key):
    engine.reopen_task(task_key, engine.subject_for_chapter(), current_auth_user(), note=request.form.get("note"))
    flash("Chapter task reopened.", "info")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/users/<int:user_id>/tasks/<task_key>/signoff")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_signoff_task(user_id, task_key):
    target = db.session.get(User, user_id)
    if not target:
        flash("User not found.", "error")
        return redirect_to_next("admin.launchpad_page")
    engine.complete_task(task_key, engine.subject_for_user(target), current_auth_user(), note=request.form.get("note"))
    flash(f"Signed off '{task_key}' for {target.name}.", "success")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/users/<int:user_id>/tasks/<task_key>/reopen")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_reopen_task(user_id, task_key):
    target = db.session.get(User, user_id)
    if not target:
        flash("User not found.", "error")
        return redirect_to_next("admin.launchpad_page")
    engine.reopen_task(task_key, engine.subject_for_user(target), current_auth_user(), note=request.form.get("note"))
    flash(f"Reopened '{task_key}' for {target.name}.", "info")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/users/<int:user_id>/training")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_record_training(user_id):
    target = db.session.get(User, user_id)
    if not target:
        flash("User not found.", "error")
        return redirect_to_next("admin.launchpad_page")
    score_raw = (request.form.get("score") or "").strip()
    engine.record_training(
        target,
        request.form.get("module_key"),
        score=int(score_raw) if score_raw.isdigit() else None,
        actor=current_auth_user(),
        note=request.form.get("note"),
    )
    flash(f"Training recorded for {target.name}.", "success")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/users/<int:user_id>/entitlements/grant")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_grant_entitlement(user_id):
    target = db.session.get(User, user_id)
    key = (request.form.get("key") or "").strip()
    if not target or key not in ALL_ENTITLEMENTS:
        flash("Pick a user and a valid entitlement.", "error")
        return redirect_to_next("admin.launchpad_page")
    days = parse_non_negative_int(request.form.get("expires_days"), default=0)
    expires_at = datetime.utcnow() + timedelta(days=days) if days else None
    entitlements.grant(
        target,
        key,
        source="override",
        source_ref="admin",
        granted_by=current_auth_user(),
        expires_at=expires_at,
        reason=request.form.get("reason"),
        commit=True,
    )
    flash(f"Granted {key} to {target.name}.", "success")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/users/<int:user_id>/entitlements/revoke")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_revoke_entitlement(user_id):
    target = db.session.get(User, user_id)
    key = (request.form.get("key") or "").strip()
    if not target or key not in ALL_ENTITLEMENTS:
        flash("Pick a user and a valid entitlement.", "error")
        return redirect_to_next("admin.launchpad_page")
    count = entitlements.revoke(target, key, revoked_by=current_auth_user(), reason=request.form.get("reason"), commit=True)
    flash(f"Revoked {count} grant(s) of {key} from {target.name}.", "info")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/loans/<int:loan_id>/signoff")
@require_role("admin")
@on_service_error("admin.launchpad_page")
def launchpad_signoff_loan(loan_id):
    loan = db.session.get(Loan, loan_id)
    if not loan:
        flash("Loan not found.", "error")
        return redirect_to_next("admin.launchpad_page")
    inventory.sign_off_loan(loan, current_auth_user())
    flash("Supervised checkout signed off.", "success")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/launchpad/reevaluate")
@require_role("admin")
def launchpad_reevaluate():
    outbox.enqueue("onboarding.evaluate_all", {})
    db.session.commit()
    if not settings().outbox_worker_enabled:
        outbox.process_pending()
    flash("Re-evaluation queued for every active member.", "info")
    return redirect_to_next("admin.launchpad_page")


@bp.post("/jobs/<int:job_id>/retry")
@require_role("admin")
def retry_job(job_id):
    job = db.session.get(OutboxJob, job_id)
    if not job:
        flash("Job not found.", "error")
        return redirect_to_next("admin.settings_page")
    outbox.retry(job)
    flash(f"Job #{job.id} queued for retry.", "info")
    return redirect_to_next("admin.settings_page")


@bp.post("/stock/discrepancies/<int:row_id>/resolve")
@require_role("admin")
def resolve_discrepancy(row_id):
    row = db.session.get(StockDiscrepancy, row_id)
    if not row:
        flash("Discrepancy not found.", "error")
        return redirect_to_next("admin.inventory_page")
    inventory.resolve_discrepancy(
        row,
        current_auth_user(),
        note=request.form.get("note"),
        apply_derived=parse_bool_flag(request.form.get("apply_derived"), default=False),
    )
    flash("Discrepancy resolved.", "success")
    return redirect_to_next("admin.inventory_page")


@bp.post("/stock/reconcile")
@require_role("admin")
def reconcile_now():
    filed = inventory.reconcile_stock(actor=current_auth_user())
    flash(f"Reconciliation complete: {len(filed)} new discrepancy(ies) filed.", "info")
    return redirect_to_next("admin.inventory_page")


# --------------------------------------------------------------------------- announcements / projects / contact


@bp.post("/announcements/<int:announcement_id>/update")
@require_role("admin")
def announcement_update(announcement_id):
    row = db.session.get(Announcement, announcement_id)
    if not row:
        flash("Announcement not found.", "error")
        return redirect_to_next("admin.announcements_page")
    content.update_announcement(row, request.form, current_auth_user())
    flash("Announcement updated.", "success")
    return redirect_to_next("admin.announcements_page")


@bp.post("/announcements/<int:announcement_id>/delete")
@require_role("admin")
def announcement_delete(announcement_id):
    row = db.session.get(Announcement, announcement_id)
    if not row:
        flash("Announcement not found.", "error")
        return redirect_to_next("admin.announcements_page")
    content.delete_announcement(row, current_auth_user())
    flash("Announcement deleted.", "success")
    return redirect_to_next("admin.announcements_page")


@bp.post("/projects/save")
@require_role("admin")
@on_service_error("admin.announcements_page")
def project_save():
    content.save_project(request.form, current_auth_user())
    flash("Project saved.", "success")
    return redirect_to_next("admin.announcements_page")


@bp.post("/contact/<int:message_id>/status")
@require_role("admin")
def contact_status(message_id):
    row = db.session.get(ContactMessage, message_id)
    if not row:
        flash("Message not found.", "error")
        return redirect_to_next("admin.announcements_page")
    content.update_contact_status(row, request.form.get("status"), request.form.get("admin_reply"), current_auth_user())
    flash("Message status updated.", "success")
    return redirect_to_next("admin.announcements_page")


# --------------------------------------------------------------------------- people


@bp.post("/members/create")
@require_role("admin")
@on_service_error("admin.members_page")
def create_member():
    member, linked_user = identity.admin_create_member(request.form, current_auth_user())
    if linked_user:
        flash(f"Member created with ID #{member.id} and linked user account.", "success")
    else:
        flash(f"Member created with ID #{member.id}.", "success")
    return redirect_to_next("admin.members_page")


@bp.post("/users/create")
@require_role("admin")
@on_service_error("admin.members_page")
def create_user():
    identity.admin_create_user(request.form, current_auth_user())
    flash("User created.", "success")
    return redirect_to_next("admin.members_page")


@bp.post("/users/<int:user_id>/role")
@require_role("admin")
@on_service_error("admin.members_page")
def update_user_role(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("admin.members_page")
    identity.admin_update_user(user, request.form, current_auth_user())
    flash("User updated.", "success")
    return redirect_to_next("admin.members_page")


@bp.post("/users/<int:user_id>/reset-password")
@require_role("admin")
@on_service_error("admin.members_page")
def reset_user_password(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("admin.members_page")
    new_password = identity.admin_reset_password(user, request.form.get("new_password"), current_auth_user())
    flash(f"Password reset for {user.email}. Temporary password: {new_password}", "info")
    return redirect_to_next("admin.members_page")


@bp.post("/users/<int:user_id>/invite-link")
@require_role("admin")
def user_invite_link(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("admin.members_page")
    token = identity.admin_invite_link_token(user, current_auth_user())
    reset_link = url_for("auth.reset_password_page", token=token, _external=True)
    flash(f"Invite/reset link for {user.email}: {reset_link}", "info")
    return redirect_to_next("admin.members_page")


@bp.post("/users/<int:user_id>/delete")
@require_role("admin")
@on_service_error("admin.members_page")
def delete_user(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("User not found.", "error")
        return redirect_to_next("admin.members_page")
    outcome = identity.admin_delete_user(user, current_auth_user())
    if outcome == "deleted":
        flash("User deleted.", "success")
    else:
        flash("User has history, so the account was deactivated instead of hard-deleted.", "info")
    return redirect_to_next("admin.members_page")


@bp.post("/members/import-roster")
@require_role("admin")
def import_roster():
    roster_file = request.files.get("roster_file")
    if not roster_file or not roster_file.filename:
        flash("Upload a roster PDF file.", "error")
        return redirect_to_next("admin.members_page")
    if not (roster_file.filename or "").strip().lower().endswith(".pdf"):
        flash("Roster import currently supports PDF files only.", "error")
        return redirect_to_next("admin.members_page")
    try:
        entries = roster.parse_roster_pdf_entries(roster_file.read())
    except Exception as exc:
        flash(f"Failed to parse roster PDF: {str(exc)[:220]}", "error")
        return redirect_to_next("admin.members_page")
    if not entries:
        flash("No roster members were found in the uploaded file.", "error")
        return redirect_to_next("admin.members_page")

    role = normalize_role(request.form.get("role") or "member")
    result = roster.import_roster_entries(
        entries,
        default_role=role,
        member_class=(request.form.get("member_class") or "Member").strip() or "Member",
        reset_existing_passwords=True,
        actor=current_auth_user(),
    )
    credentials = result.get("credentials") or []
    if not credentials:
        db.session.rollback()
        flash("Roster processed, but no account credentials were generated.", "info")
        return redirect_to_next("admin.members_page")

    from asme.services import audit

    audit.record(
        "bulk_roster_import",
        (
            f"entries={len(entries)} created_members={result['created_members']} "
            f"created_users={result['created_users']} updated_users={result['updated_users']} "
            f"reset_passwords={result['reset_passwords']} file=inline_csv_download"
        ),
        actor=current_auth_user(),
    )
    db.session.commit()
    roster.emit_import_events(result)
    csv_content = csv_stream_from_rows(
        ["name", "member_id", "username", "password"],
        [[row.get("name") or "", row.get("member_id") or "", row.get("username") or "", row.get("password") or ""] for row in credentials],
    )
    download_name = f"member_credentials_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(csv_content, mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename={download_name}"})


@bp.get("/members/import-credentials/<filename>")
@require_role("admin")
def download_roster_credentials(filename):
    safe_name = Path(filename or "").name
    if not safe_name:
        flash("Invalid credentials file name.", "error")
        return redirect(url_for("admin.members_page"))
    file_path = roster.roster_credentials_dir() / safe_name
    if not file_path.exists() or not file_path.is_file():
        flash(
            "Credentials file not found. On Render free instances, uploaded/generated files can be lost on restart. "
            "Use 'Download Fresh Credentials CSV' to regenerate working passwords.",
            "error",
        )
        return redirect(url_for("admin.members_page"))
    return send_file(file_path, as_attachment=True, download_name=safe_name, mimetype="text/csv")


@bp.post("/members/export-credentials")
@require_role("admin")
def export_fresh_credentials():
    from asme.services import audit

    rows = identity.build_fresh_member_credentials()
    if not rows:
        flash("No active member/team_leader users found.", "error")
        return redirect_to_next("admin.members_page")
    audit.record("export_fresh_member_credentials", f"user_count={len(rows)}", actor=current_auth_user())
    db.session.commit()
    csv_content = csv_stream_from_rows(
        ["name", "member_id", "username", "password"],
        [[row.get("name") or "", row.get("member_id") or "", row.get("username") or "", row.get("password") or ""] for row in rows],
    )
    download_name = f"member_credentials_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(csv_content, mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename={download_name}"})


@bp.post("/members/export-credentials-pdf")
@require_role("admin")
def export_fresh_credentials_pdf():
    from asme.services import audit

    rows = identity.build_fresh_member_credentials()
    if not rows:
        flash("No active member/team_leader users found.", "error")
        return redirect_to_next("admin.members_page")
    try:
        pdf_bytes = roster.credentials_pdf_bytes(rows)
    except Exception:
        db.session.rollback()
        flash("PDF export requires reportlab. Add reportlab to requirements and redeploy.", "error")
        return redirect_to_next("admin.members_page")
    audit.record("export_fresh_member_credentials_pdf", f"user_count={len(rows)}", actor=current_auth_user())
    db.session.commit()
    download_name = f"member_credentials_name_last_nfc_user_pass_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(BytesIO(pdf_bytes), as_attachment=True, download_name=download_name, mimetype="application/pdf")


@bp.post("/members/reset")
@require_role("admin")
def reset_members():
    if (request.form.get("confirmation") or "").strip().upper() != "RESET MEMBERS":
        flash("Type RESET MEMBERS to confirm.", "error")
        return redirect_to_next("admin.members_page")
    counts = identity.reset_members(current_auth_user())
    flash(
        f"Members reset complete: {counts['deleted_users']} user accounts deleted, "
        f"{counts['deactivated_users']} deactivated, {counts['deleted_members']} member profiles deleted, "
        f"{counts['kept_members']} retained for history.",
        "success",
    )
    return redirect_to_next("admin.members_page")


@bp.post("/nfc/assign")
@require_role("admin")
@on_service_error("admin.members_page")
def assign_nfc():
    user = db.session.get(User, parse_positive_int(request.form.get("user_id"), default=0))
    if not user:
        flash("User and tag UID are required.", "error")
        return redirect_to_next("admin.members_page")
    identity.assign_nfc(user, request.form.get("tag_uid"), (request.form.get("notes") or "").strip() or None, current_auth_user())
    flash("NFC tag assignment updated.", "success")
    return redirect_to_next("admin.members_page")


@bp.post("/nfc/unassign/<int:tag_id>")
@require_role("admin")
def unassign_nfc(tag_id):
    row = db.session.get(NFCTag, tag_id)
    if not row:
        flash("NFC tag assignment not found.", "error")
        return redirect_to_next("admin.members_page")
    identity.unassign_nfc(row, current_auth_user())
    flash("Tag unassigned.", "success")
    return redirect_to_next("admin.members_page")


# --------------------------------------------------------------------------- events / attendance


@bp.post("/events/create")
@require_role("admin")
@on_service_error("admin.attendance_page")
def create_event():
    scheduling.create_event(request.form, current_auth_user())
    flash("Event saved.", "success")
    return redirect_to_next("admin.attendance_page")


@bp.post("/events/<int:event_id>/status")
@require_role("admin")
@on_service_error("admin.attendance_page")
def event_status_update(event_id):
    event = db.session.get(Event, event_id)
    if not event:
        flash("Event not found.", "error")
        return redirect_to_next("admin.attendance_page")
    scheduling.update_event(event, request.form, current_auth_user())
    flash("Event updated.", "success")
    return redirect_to_next("admin.attendance_page")


@bp.post("/attendance/checkin")
@require_role("admin")
def attendance_checkin():
    event = db.session.get(Event, parse_positive_int(request.form.get("event_id"), default=0))
    if not event:
        flash("Event not found.", "error")
        return redirect_to_next("admin.attendance_page")
    try:
        attendance.admin_checkin(
            event,
            tag_uid=request.form.get("tag_uid"),
            user_id=parse_positive_int(request.form.get("user_id"), default=0),
            requested_method=(request.form.get("checkin_method") or "").strip().lower(),
            actor=current_auth_user(),
        )
    except ServiceError as exc:
        flash(exc.message, "info" if exc.status == 409 else "error")
        return redirect_to_next("admin.attendance_page")
    flash(f"You have been marked present for {event.title}.", "success")
    return redirect_to_next("admin.attendance_page")


@bp.get("/attendance/export.csv")
@require_role("admin")
def attendance_export():
    content_csv = csv_stream_from_rows(attendance.EXPORT_HEADERS, attendance.export_rows())
    return Response(content_csv, mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=attendance_records.csv"})


# --------------------------------------------------------------------------- inventory


@bp.post("/inventory/item/save")
@require_role("admin")
@on_service_error("admin.inventory_page")
def inventory_item_save():
    inventory.save_item(request.form, current_auth_user())
    flash("Inventory item saved.", "success")
    return redirect_to_next("admin.inventory_page")


@bp.post("/inventory/adjust")
@require_role("admin")
def inventory_adjust():
    item = db.session.get(Item, parse_positive_int(request.form.get("item_id"), default=0))
    if not item:
        flash("Item not found.", "error")
        return redirect_to_next("admin.inventory_page")
    inventory.adjust_counts(
        item,
        parse_non_negative_int(request.form.get("total_qty"), default=0),
        parse_non_negative_int(request.form.get("available_qty"), default=0),
        request.form.get("notes"),
        current_auth_user(),
    )
    flash("Inventory counts updated.", "success")
    return redirect_to_next("admin.inventory_page")


@bp.post("/inventory/bootstrap-counts")
@require_role("admin")
def inventory_bootstrap_counts():
    updated = inventory.bootstrap_counts(parse_positive_int(request.form.get("starter_qty"), default=1), current_auth_user())
    if updated == 0:
        flash("No items needed starter count updates.", "info")
    else:
        flash(f"Starter counts applied to {updated} item(s).", "success")
    return redirect_to_next("admin.inventory_page")


@bp.post("/inventory/reset")
@require_role("admin")
def inventory_reset():
    if (request.form.get("confirmation") or "").strip().upper() != "RESET INVENTORY":
        flash("Type RESET INVENTORY to confirm the reset.", "error")
        return redirect_to_next("admin.inventory_page")
    counts = inventory.reset_inventory(current_auth_user())
    flash(
        f"Inventory reset complete: deleted {counts['items']} items, {counts['tags']} item tags, and {counts['transactions']} inventory transactions.",
        "success",
    )
    return redirect_to_next("admin.inventory_page")


@bp.post("/inventory/import")
@require_role("admin")
def inventory_bulk_import():
    try:
        rows = inventory.parse_bulk_inventory_rows(request.files.get("inventory_file"))
        result = inventory.bulk_import(rows, current_auth_user())
    except ServiceError as exc:
        flash_error(exc)
        return redirect_to_next("admin.inventory_page")
    flash(f"Bulk import complete: {result['created']} created, {result['updated']} updated, {result['skipped']} skipped.", "success")
    if result["errors"]:
        preview = "; ".join(result["errors"][:4])
        if len(result["errors"]) > 4:
            preview += f"; +{len(result['errors']) - 4} more"
        flash(preview, "info")
    return redirect_to_next("admin.inventory_page")


@bp.get("/inventory/items-tags.pdf")
@require_role("admin")
def inventory_item_nfc_pdf():
    try:
        pdf_bytes = inventory.item_nfc_pdf_bytes()
    except Exception:
        flash("PDF export requires reportlab. Add reportlab to requirements and redeploy.", "error")
        return redirect(url_for("admin.inventory_page", tab="items"))
    from datetime import date

    return send_file(BytesIO(pdf_bytes), as_attachment=True, download_name=f"inventory_item_nfc_{date.today().isoformat()}.pdf", mimetype="application/pdf")


@bp.get("/inventory/treasury-report.csv")
@require_role("admin")
def inventory_treasury_report():
    content_csv = csv_stream_from_rows(inventory.TREASURY_REPORT_HEADERS, inventory.treasury_report_rows())
    return Response(content_csv, mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=inventory_treasury_report.csv"})


# --------------------------------------------------------------------------- prints


@bp.post("/print/<int:request_id>/status")
@require_role("admin")
@on_service_error("admin.prints_page")
def print_status(request_id):
    row = db.session.get(PrintRequest, request_id)
    if not row:
        flash("Print request not found.", "error")
        return redirect_to_next("admin.prints_page")
    fabrication.update_print_request(row, request.form.get("status"), request.form.get("printer_type"), request.form.get("admin_notes"), current_auth_user())
    flash("Print request updated.", "success")
    return redirect_to_next("admin.prints_page")
