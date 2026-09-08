"""Shared-device flows: kiosk login tag, meeting check-in tag, NFC pairing pages."""

from __future__ import annotations

from flask import Blueprint, flash, redirect, render_template, request, session, url_for
from sqlalchemy import func

from asme.auth.session import current_auth_user, require_login
from asme.blueprints._context import public_site_context
from asme.extensions import db
from asme.models import Event, Item, ItemTag, Member, NFCTag
from asme.services import attendance
from asme.utils import parse_positive_int

bp = Blueprint("kiosk", __name__)


@bp.get("/kiosk")
def kiosk_entry():
    if current_auth_user():
        return redirect(url_for("portal.member_inventory"))
    context = public_site_context("Kiosk Login")
    context["inventory_next"] = url_for("portal.member_inventory")
    context["admin_next"] = url_for("admin.dashboard")
    return render_template("site/kiosk.html", **context)


@bp.get("/checkin")
@require_login
def checkin_entry():
    user = current_auth_user()
    candidates = attendance.candidate_events()
    if not candidates:
        context = public_site_context("Meeting Check-in")
        context["active_event"] = None
        return render_template("site/checkin.html", **context)
    if len(candidates) == 1:
        event = candidates[0]
        _row, created = attendance.check_user_into_event(user, event, method="shared_tag")
        return redirect(url_for("kiosk.checkin_success", event_id=event.id, status="new" if created else "existing"))
    session["checkin_candidate_ids"] = [row.id for row in candidates]
    return redirect(url_for("kiosk.checkin_select"))


@bp.route("/checkin/select", methods=["GET", "POST"])
@require_login
def checkin_select():
    candidate_ids = session.get("checkin_candidate_ids") or []
    if not candidate_ids:
        flash("No active meeting choices. Try scanning again.", "info")
        return redirect(url_for("kiosk.checkin_entry"))
    candidates = Event.query.filter(Event.id.in_(candidate_ids)).order_by(Event.start_time.asc(), Event.id.asc()).all()
    if not candidates:
        session.pop("checkin_candidate_ids", None)
        flash("No active meeting choices. Try scanning again.", "info")
        return redirect(url_for("kiosk.checkin_entry"))
    if request.method == "POST":
        event_id = parse_positive_int(request.form.get("event_id"), default=0)
        event = next((row for row in candidates if row.id == event_id), None)
        if not event:
            flash("Select a valid meeting.", "error")
            return redirect(url_for("kiosk.checkin_select"))
        _row, created = attendance.check_user_into_event(current_auth_user(), event, method="shared_tag")
        session.pop("checkin_candidate_ids", None)
        return redirect(url_for("kiosk.checkin_success", event_id=event.id, status="new" if created else "existing"))
    context = public_site_context("Select Meeting")
    context["candidate_events"] = candidates
    return render_template("site/checkin_select.html", **context)


@bp.get("/checkin/success")
@require_login
def checkin_success():
    event_id = parse_positive_int(request.args.get("event_id"), default=0)
    status = (request.args.get("status") or "new").strip().lower()
    event = db.session.get(Event, event_id)
    if not event:
        flash("Meeting not found.", "error")
        return redirect(url_for("kiosk.checkin_entry"))
    context = public_site_context("Check-in Success")
    context["event"] = event
    context["status"] = status
    return render_template("site/checkin_success.html", **context)


# --------------------------------------------------------------------------- legacy pairing pages


@bp.route("/pair/member", methods=["GET", "POST"])
def pair_member():
    if request.method == "POST":
        member_id_raw = (request.form.get("member_id") or "").strip()
        tag = (request.form.get("tag") or "").strip()
        if not member_id_raw or not tag:
            flash("Member and UID are required.", "error")
            return redirect(url_for("kiosk.pair_member"))
        try:
            member_id = int(member_id_raw)
        except Exception:
            flash("Invalid member selection.", "error")
            return redirect(url_for("kiosk.pair_member"))
        existing = Member.query.filter_by(nfc_tag=tag).first()
        if existing and existing.id != member_id:
            flash("That UID is already assigned to another member.", "error")
            return redirect(url_for("kiosk.pair_member"))
        if ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag.lower()).first():
            flash("That UID is already assigned to an inventory item tag.", "error")
            return redirect(url_for("kiosk.pair_member"))
        if Item.query.filter(func.lower(Item.nfc_tag) == tag.lower()).first():
            flash("That UID is already assigned to an inventory item.", "error")
            return redirect(url_for("kiosk.pair_member"))
        if NFCTag.query.filter(func.lower(NFCTag.tag_uid) == tag.lower(), NFCTag.active.is_(True)).first():
            flash("That UID is already assigned to a user account tag.", "error")
            return redirect(url_for("kiosk.pair_member"))
        member = db.session.get(Member, member_id)
        if not member:
            flash("Member not found.", "error")
            return redirect(url_for("kiosk.pair_member"))
        member.nfc_tag = tag
        db.session.commit()
        flash(f"Saved UID for {member.name}.", "success")
        return redirect(url_for("kiosk.pair_member"))
    members = Member.query.order_by(Member.name.asc()).all()
    paired = Member.query.filter(Member.nfc_tag.isnot(None)).order_by(Member.name.asc()).all()
    return render_template("pair_member.html", members=members, paired=paired)


@bp.route("/pair/item", methods=["GET", "POST"])
def pair_item():
    if request.method == "POST":
        item_id_raw = (request.form.get("item_id") or "").strip()
        tag = (request.form.get("tag") or "").strip()
        if not item_id_raw or not tag:
            flash("Item and UID are required.", "error")
            return redirect(url_for("kiosk.pair_item"))
        try:
            item_id = int(item_id_raw)
        except Exception:
            flash("Invalid item selection.", "error")
            return redirect(url_for("kiosk.pair_item"))
        existing = Item.query.filter_by(nfc_tag=tag).first()
        if existing and existing.id != item_id:
            flash("That UID is already assigned to another item.", "error")
            return redirect(url_for("kiosk.pair_item"))
        if NFCTag.query.filter(func.lower(NFCTag.tag_uid) == tag.lower()).first():
            flash("That UID is already assigned to a member tag.", "error")
            return redirect(url_for("kiosk.pair_item"))
        item = db.session.get(Item, item_id)
        if not item:
            flash("Item not found.", "error")
            return redirect(url_for("kiosk.pair_item"))
        mapped = ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag.lower()).first()
        if mapped and mapped.item_id != item.id:
            flash("That UID is already assigned in item tag map.", "error")
            return redirect(url_for("kiosk.pair_item"))
        item.nfc_tag = tag
        if not mapped:
            db.session.add(ItemTag(item_id=item.id, tag_value=tag, source="pair_item_page"))
        db.session.commit()
        flash(f"Saved UID for {item.name}.", "success")
        return redirect(url_for("kiosk.pair_item"))
    items = Item.query.order_by(Item.name.asc()).all()
    paired = Item.query.filter(Item.nfc_tag.isnot(None)).order_by(Item.name.asc()).all()
    return render_template("pair_item.html", items=items, paired=paired)
