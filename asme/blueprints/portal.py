"""Member and team-leader portal (HTML)."""

from __future__ import annotations

from datetime import datetime

from flask import Blueprint, flash, redirect, render_template, request, url_for

from asme.auth.session import current_auth_user, current_user_member, require_entitlement, require_login, require_role, role_allows
from asme.blueprints._context import member_dashboard_context
from asme.blueprints._helpers import flash_error, on_service_error
from asme.config import settings
from asme.constants import ENT_PRINT_SUBMIT, ENT_ROOM_BOOKING, ENT_SHOP_ACCESS, MEETING_ROOMS
from asme.extensions import db
from asme.models import Event, Item, Loan, NFCTag, PrintRequest, Project, TrainingModule, User
from asme.services import attendance, content, fabrication, identity, inventory, scheduling
from asme.services.errors import ServiceError
from asme.services.onboarding import engine, entitlements
from asme.utils import parse_due_date, parse_positive_int
from asme.utils.http import redirect_to_next

bp = Blueprint("portal", __name__, url_prefix="/portal")


@bp.get("")
@require_login
def portal_router():
    user = current_auth_user()
    if role_allows(user.role, "admin"):
        return redirect(url_for("admin.dashboard"))
    return redirect(url_for("portal.member_dashboard"))


def _page(template, title, active_page, **extra):
    context = member_dashboard_context()
    context["page_title"] = title
    context["active_page"] = active_page
    context.update(extra)
    return render_template(template, **context)


# --------------------------------------------------------------------------- member pages


@bp.get("/member")
@require_role("member")
def member_dashboard():
    context = member_dashboard_context()
    context["page_title"] = "Member Dashboard"
    context["active_page"] = "member_dashboard"
    context["tile_counts"] = {
        "inventory": len(context.get("items", [])),
        "my_items": len(context.get("open_checkouts", [])),
        "prints": len(context.get("print_requests", [])),
        "upcoming_events": len(context.get("upcoming_events", [])),
        "announcements": len(context.get("announcements", [])),
    }
    return render_template("portal/member_dashboard_home.html", **context)


@bp.get("/member/inventory")
@require_role("member")
def member_inventory():
    context = member_dashboard_context()
    context["page_title"] = "Inventory"
    context["active_page"] = "member_inventory"
    query = (request.args.get("q") or "").strip().lower()
    if query:
        context["items"] = [
            item
            for item in context["items"]
            if query in " ".join([item.name or "", item.category or "", item.location or "", item.description or "", item.notes or ""]).lower()
        ]
    context["query"] = query
    return render_template("portal/member_inventory.html", **context)


@bp.get("/member/items/<int:item_id>")
@bp.get("/member/inventory/<int:item_id>")
@require_role("member")
def member_item_detail(item_id):
    item = db.session.get(Item, item_id)
    if not item:
        flash("Item not found.", "error")
        return redirect(url_for("portal.member_inventory"))
    user = current_auth_user()
    member = current_user_member(user)
    open_tx = inventory.get_open_loan(item.id, user=user, member=member)
    return _page("portal/member_item_detail.html", "Item Detail", "member_inventory", item=item, open_tx=open_tx)


@bp.get("/member/checkouts")
@bp.get("/member/my-items")
@require_role("member")
def member_checkouts():
    return _page("portal/member_checkouts.html", "My Items", "member_my_items")


@bp.get("/member/prints")
@require_role("member")
def member_prints():
    return _page("portal/member_prints.html", "3D Printing", "member_prints")


@bp.get("/member/prints/<int:request_id>")
@require_role("member")
def member_print_detail(request_id):
    row = db.session.get(PrintRequest, request_id)
    user = current_auth_user()
    if not row or row.user_id != user.id:
        flash("Print request not found.", "error")
        return redirect(url_for("portal.member_prints"))
    return _page("portal/member_print_detail.html", "Print Request Detail", "member_prints", print_request=row)


@bp.get("/member/calendar")
@require_role("member")
def member_calendar():
    user = current_auth_user()
    return _page(
        "portal/member_calendar.html",
        "Calendar",
        "member_calendar",
        team_mode=role_allows(user.role, "team_leader"),
        meeting_rooms=MEETING_ROOMS,
        calendar_provider=settings().calendar_provider,
    )


@bp.route("/member/schedule", methods=["GET", "POST"])
@require_role("team_leader")
@require_entitlement(ENT_ROOM_BOOKING)
def member_schedule():
    cfg = settings()
    context = member_dashboard_context()
    context["page_title"] = "Schedule Meeting"
    context["active_page"] = "member_schedule"
    context["meeting_rooms"] = MEETING_ROOMS
    context["calendar_provider"] = cfg.calendar_provider
    status = scheduling.provider_status()
    context["google_schedule_enabled"] = status.enabled
    context["google_schedule_errors"] = status.errors
    context["duration_options"] = [30, 60, 90]
    context["range_options"] = [7, 14]
    form_values = {
        "team_name": "",
        "duration": 60,
        "days": min(max(cfg.calendar_scheduling_days, 7), 14),
        "location": "",
        "notes": "",
    }
    slots = []

    if request.method == "POST":
        form_values["team_name"] = (request.form.get("team_name") or "").strip()
        form_values["duration"] = parse_positive_int(request.form.get("duration"), default=60)
        if form_values["duration"] not in {30, 60, 90}:
            form_values["duration"] = 60
        form_values["days"] = parse_positive_int(request.form.get("days"), default=14)
        if form_values["days"] not in {7, 14}:
            form_values["days"] = 14
        form_values["location"] = scheduling.normalize_meeting_room(request.form.get("location"))
        form_values["notes"] = (request.form.get("notes") or "").strip()
        action = (request.form.get("action") or "preview").strip().lower()
        if not form_values["team_name"]:
            flash("Team name is required.", "error")
            return render_template("portal/member_schedule.html", **context, form_values=form_values, slots=slots)

        if action == "book":
            if not status.enabled:
                flash("Admin needs to configure the calendar integration before scheduling.", "error")
                return render_template("portal/member_schedule.html", **context, form_values=form_values, slots=slots)
            try:
                scheduling.book_slot(
                    context["current_user"],
                    form_values["team_name"],
                    request.form.get("slot_token"),
                    notes=form_values["notes"],
                )
            except ServiceError as exc:
                flash_error(exc)
                slots, _err = scheduling.compute_available_slots(form_values["duration"], form_values["days"], form_values["location"])
                return render_template("portal/member_schedule.html", **context, form_values=form_values, slots=slots)
            flash("Meeting scheduled. The calendar link will appear once it syncs.", "success")
            return redirect(url_for("portal.member_calendar"))

        slots, slot_error = scheduling.compute_available_slots(form_values["duration"], form_values["days"], form_values["location"])
        if slot_error:
            flash(slot_error, "error")

    context["form_values"] = form_values
    context["slots"] = slots
    return render_template("portal/member_schedule.html", **context)


@bp.get("/leader/schedule")
@require_role("team_leader")
def leader_schedule_alias():
    return redirect(url_for("portal.member_schedule"))


@bp.get("/team")
@require_role("team_leader")
def team_dashboard():
    return redirect(url_for("portal.leader_launchpad"))


@bp.route("/member/help", methods=["GET", "POST"])
@require_role("member")
def member_help():
    context = member_dashboard_context()
    context["page_title"] = "Help / Ask Leads"
    context["active_page"] = "member_help"
    user = context["current_user"]
    if request.method == "POST":
        message = (request.form.get("message") or "").strip()
        if not message:
            flash("Message is required.", "error")
            return render_template("portal/member_help.html", **context)
        content.create_contact_message(
            name=user.name,
            email=user.email,
            message=message,
            kind="help",
            user=user,
            target=(request.form.get("target") or "").strip() or "team_leads",
            subject=(request.form.get("subject") or "").strip() or "Help request",
        )
        flash("Help request sent.", "success")
        return redirect(url_for("portal.member_help"))
    return render_template("portal/member_help.html", **context)


@bp.route("/member/profile", methods=["GET", "POST"])
@require_role("member")
def member_profile():
    context = member_dashboard_context()
    context["page_title"] = "Profile / Account Settings"
    context["active_page"] = "member_profile"
    user = context["current_user"]
    context["active_tag"] = (
        NFCTag.query.filter_by(user_id=user.id, active=True).order_by(NFCTag.assigned_at.desc(), NFCTag.id.desc()).first()
    )
    if request.method == "POST":
        action = (request.form.get("action") or "").strip().lower()
        try:
            if action == "profile":
                grad_raw = (request.form.get("graduation_year") or "").strip()
                identity.update_profile(
                    user,
                    request.form.get("name"),
                    request.form.get("email"),
                    major=request.form.get("major") if "major" in request.form else None,
                    graduation_year=(parse_positive_int(grad_raw, default=0) if grad_raw else 0) if "graduation_year" in request.form else None,
                    phone=request.form.get("phone") if "phone" in request.form else None,
                )
                flash("Profile updated.", "success")
                return redirect(url_for("portal.member_profile"))
            if action == "password":
                identity.change_password(
                    user,
                    request.form.get("current_password"),
                    request.form.get("new_password"),
                    request.form.get("confirm_password"),
                )
                flash("Password updated.", "success")
                return redirect(url_for("portal.member_profile"))
        except ServiceError as exc:
            flash_error(exc)
            return render_template("portal/member_profile.html", **context)
    return render_template("portal/member_profile.html", **context)


# --------------------------------------------------------------------------- launchpad (member)


@bp.get("/member/launchpad")
@require_role("member")
def member_launchpad():
    user = current_auth_user()
    progress = engine.evaluate_user(user)
    memberships = content.active_memberships(user)
    return _page(
        "portal/member_launchpad.html",
        "Launchpad",
        "member_launchpad",
        progress=progress,
        track=progress["tracks"][0] if progress["tracks"] else None,
        memberships=memberships,
        joinable_projects=content.joinable_projects(),
        training_modules=TrainingModule.query.order_by(TrainingModule.id.asc()).all(),
        hours_total=content.total_hours(user),
        entitlement_keys=sorted(entitlements.user_entitlements(user)),
    )


@bp.post("/member/launchpad/tasks/<task_key>/complete")
@require_role("member")
@on_service_error("portal.member_launchpad")
def member_complete_task(task_key):
    user = current_auth_user()
    engine.complete_task(task_key, engine.subject_for_user(user), user, note=request.form.get("note"))
    flash("Task marked complete.", "success")
    return redirect_to_next("portal.member_launchpad")


@bp.post("/member/launchpad/teams/join")
@require_role("member")
@on_service_error("portal.member_launchpad")
def member_join_team():
    user = current_auth_user()
    project = db.session.get(Project, parse_positive_int(request.form.get("project_id"), default=0))
    if not project:
        flash("Select a team.", "error")
        return redirect_to_next("portal.member_launchpad")
    content.join_project(user, project)
    flash(f"You joined {project.title}.", "success")
    return redirect_to_next("portal.member_launchpad")


@bp.post("/member/launchpad/teams/leave")
@require_role("member")
@on_service_error("portal.member_launchpad")
def member_leave_team():
    user = current_auth_user()
    project = db.session.get(Project, parse_positive_int(request.form.get("project_id"), default=0))
    if not project:
        flash("Select a team.", "error")
        return redirect_to_next("portal.member_launchpad")
    content.leave_project(user, project)
    flash(f"You left {project.title}.", "info")
    return redirect_to_next("portal.member_launchpad")


@bp.post("/member/launchpad/hours")
@require_role("member")
@on_service_error("portal.member_launchpad")
def member_log_hours():
    user = current_auth_user()
    project = db.session.get(Project, parse_positive_int(request.form.get("project_id"), default=0))
    logged_for = parse_due_date(request.form.get("logged_for"))
    content.log_hours(user, project, request.form.get("hours"), note=request.form.get("note"), logged_for=logged_for)
    flash("Hours logged.", "success")
    return redirect_to_next("portal.member_launchpad")


# --------------------------------------------------------------------------- launchpad (team lead)


@bp.get("/leader/launchpad")
@require_role("team_leader")
def leader_launchpad():
    from asme.blueprints.admin import launchpad_admin_context

    context = member_dashboard_context()
    context.update(launchpad_admin_context(is_admin_view=False))
    context["page_title"] = "Team Launchpad"
    context["active_page"] = "member_launchpad_lead"
    return render_template("portal/admin_launchpad.html", **context)


@bp.post("/leader/users/<int:user_id>/tasks/<task_key>/signoff")
@require_role("team_leader")
@on_service_error("portal.leader_launchpad")
def leader_signoff_task(user_id, task_key):
    actor = current_auth_user()
    target = db.session.get(User, user_id)
    if not target:
        flash("User not found.", "error")
        return redirect_to_next("portal.leader_launchpad")
    engine.complete_task(task_key, engine.subject_for_user(target), actor, note=request.form.get("note"))
    flash(f"Signed off '{task_key}' for {target.name}.", "success")
    return redirect_to_next("portal.leader_launchpad")


@bp.post("/leader/users/<int:user_id>/training")
@require_role("team_leader")
@on_service_error("portal.leader_launchpad")
def leader_record_training(user_id):
    actor = current_auth_user()
    target = db.session.get(User, user_id)
    if not target:
        flash("User not found.", "error")
        return redirect_to_next("portal.leader_launchpad")
    score_raw = (request.form.get("score") or "").strip()
    engine.record_training(
        target,
        request.form.get("module_key"),
        score=int(score_raw) if score_raw.isdigit() else None,
        actor=actor,
        note=request.form.get("note"),
    )
    flash(f"Training recorded for {target.name}.", "success")
    return redirect_to_next("portal.leader_launchpad")


@bp.post("/leader/loans/<int:loan_id>/signoff")
@require_role("team_leader")
@on_service_error("portal.leader_launchpad")
def leader_signoff_loan(loan_id):
    actor = current_auth_user()
    loan = db.session.get(Loan, loan_id)
    if not loan:
        flash("Loan not found.", "error")
        return redirect_to_next("portal.leader_launchpad")
    inventory.sign_off_loan(loan, actor)
    flash("Supervised checkout signed off.", "success")
    return redirect_to_next("portal.leader_launchpad")


# --------------------------------------------------------------------------- inventory / prints / events (write)


@bp.post("/inventory/checkout")
@require_role("member")
@require_entitlement(ENT_SHOP_ACCESS)
@on_service_error("portal.member_inventory")
def checkout():
    user = current_auth_user()
    member = current_user_member(user)
    item = db.session.get(Item, parse_positive_int(request.form.get("item_id"), default=0))
    if not item:
        flash("Item not found.", "error")
        return redirect_to_next("portal.member_inventory")
    qty = parse_positive_int(request.form.get("qty"), default=1)
    loan = inventory.checkout(
        item=item,
        user=user,
        member=member,
        qty=qty,
        notes=(request.form.get("notes") or "").strip() or None,
        due_date=parse_due_date(request.form.get("due_date")),
        idempotency_key=(request.form.get("idempotency_key") or "").strip() or None,
        source="portal",
    )
    flash(f"Checked out {loan.qty} x {item.name}.", "success")
    return redirect_to_next("portal.member_checkouts")


@bp.post("/inventory/return")
@require_role("member")
@on_service_error("portal.member_checkouts")
def return_item():
    user = current_auth_user()
    member = current_user_member(user)
    item = db.session.get(Item, parse_positive_int(request.form.get("item_id"), default=0))
    if not item:
        flash("Item not found.", "error")
        return redirect_to_next("portal.member_checkouts")
    open_loan = inventory.get_open_loan(item.id, user=user, member=member)
    if not open_loan:
        flash("No open checkout found for this item.", "error")
        return redirect_to_next("portal.member_checkouts")
    inventory.return_loan(
        loan=open_loan,
        qty=parse_positive_int(request.form.get("qty"), default=open_loan.qty),
        condition=(request.form.get("condition") or "").strip() or "good",
        notes=(request.form.get("notes") or "").strip() or None,
        user=user,
        source="portal",
    )
    flash(f"Returned {open_loan.qty} x {item.name}.", "success")
    return redirect_to_next("portal.member_checkouts")


@bp.post("/print/request")
@require_role("member")
@require_entitlement(ENT_PRINT_SUBMIT)
@on_service_error("portal.member_prints")
def print_request():
    user = current_auth_user()
    member = current_user_member(user)
    fabrication.submit_print_request(user, request.form, file_upload=request.files.get("print_file"), member=member)
    flash("3D print request submitted.", "success")
    return redirect_to_next("portal.member_prints")


@bp.post("/events/request")
@require_role("team_leader")
def event_request():
    flash("Use Schedule Meeting for slot-based availability and room conflict prevention.", "info")
    return redirect(url_for("portal.member_schedule"))


@bp.post("/events/<int:event_id>/rsvp")
@require_role("member")
def event_rsvp(event_id):
    user = current_auth_user()
    event = db.session.get(Event, event_id)
    if not event:
        flash("Event not found.", "error")
        return redirect_to_next("portal.member_calendar")
    _row, created = attendance.rsvp(user, event)
    flash("RSVP saved." if created else "RSVP already recorded.", "success" if created else "info")
    return redirect_to_next("portal.member_calendar")
