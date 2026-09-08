"""Public marketing site."""

from __future__ import annotations

import os
from pathlib import Path

from flask import Blueprint, current_app, flash, jsonify, redirect, render_template, request, send_from_directory, url_for

from asme.blueprints._context import public_site_context
from asme.blueprints._helpers import flash_error
from asme.content_data import FRONT_ABOUT_ASME_FACTS, FRONT_CLUB_HIGHLIGHTS, FRONT_CLUB_MISSION, FRONT_UIOWA_CURRENT_PROJECTS
from asme.models import Project
from asme.services import content
from asme.services.errors import ServiceError
from asme.utils import parse_json_list

bp = Blueprint("public", __name__)


@bp.get("/")
def public_home():
    return render_template("site/landing.html")


@bp.get("/healthz")
def health_check():
    return jsonify({"ok": True, "service": "asme-web"})


@bp.get("/events")
def public_events():
    return render_template("site/events.html", **public_site_context("Events"))


@bp.get("/gallery")
def public_gallery():
    exts = {".jpg", ".jpeg", ".png", ".webp", ".gif"}

    def _load(subdir: str):
        folder = Path(current_app.static_folder) / "images" / subdir
        if not folder.is_dir():
            return []
        thumbs_dir = folder / "_thumbs"
        files = sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in exts), key=lambda p: p.name.lower())
        result = []
        for p in files:
            thumb_name = p.stem + ".jpg"
            has_thumb = (thumbs_dir / thumb_name).is_file()
            result.append(
                {
                    "url": url_for("static", filename=f"images/{subdir}/{p.name}"),
                    "thumb": url_for("static", filename=f"images/{subdir}/_thumbs/{thumb_name}")
                    if has_thumb
                    else url_for("static", filename=f"images/{subdir}/{p.name}"),
                }
            )
        return result

    context = public_site_context("Gallery")
    context["gallery_images"] = _load("gallery")
    context["makeathon_images"] = _load("makeathon")
    return render_template("site/gallery.html", **context)


@bp.route("/arm-sim/<path:subpath>")
def arm_sim_assets(subpath):
    return send_from_directory(os.path.join(current_app.static_folder, "arm-sim"), subpath)


@bp.route("/arm-sim")
def arm_sim_index():
    return send_from_directory(os.path.join(current_app.static_folder, "arm-sim"), "index.html")


@bp.route("/stl/<path:subpath>")
def arm_sim_stl(subpath):
    return send_from_directory(os.path.join(current_app.static_folder, "arm-sim", "stl"), subpath)


@bp.route("/hdri/<path:subpath>")
def arm_sim_hdri(subpath):
    return send_from_directory(os.path.join(current_app.static_folder, "arm-sim", "hdri"), subpath)


@bp.get("/who-we-are")
def public_who_we_are():
    context = public_site_context("Who We Are")
    context["mission"] = FRONT_CLUB_MISSION
    context["highlights"] = FRONT_CLUB_HIGHLIGHTS
    context["asme_facts"] = FRONT_ABOUT_ASME_FACTS
    context["uiowa_projects"] = FRONT_UIOWA_CURRENT_PROJECTS
    return render_template("site/who_we_are.html", **context)


@bp.get("/about")
def public_about_alias():
    return redirect(url_for("public.public_who_we_are"))


@bp.get("/executive-team")
def public_executive_team():
    return render_template("site/executive_team.html", **public_site_context("Executive Team"))


@bp.get("/exec")
def public_exec_alias():
    return redirect(url_for("public.public_executive_team"))


@bp.get("/projects")
def public_projects():
    return render_template("site/projects.html", **public_site_context("Projects"))


@bp.get("/projects/<slug>")
def public_project_detail(slug):
    project = Project.query.filter_by(slug=(slug or "").strip().lower()).first()
    if not project:
        flash("Project not found.", "error")
        return redirect(url_for("public.public_projects"))
    context = public_site_context(project.title)
    context["project"] = project
    context["project_gallery"] = parse_json_list(project.gallery_json)
    context["project_timeline"] = parse_json_list(project.timeline)
    slug_clean = (project.slug or "").strip().lower()
    per_slug_glb = os.path.join(current_app.static_folder, "models", "projects", f"{slug_clean}.glb")
    if os.path.isfile(per_slug_glb):
        context["project_model_url"] = url_for("static", filename=f"models/projects/{slug_clean}.glb")
    else:
        context["project_model_url"] = url_for("static", filename="models/hero/humanoid-soldering.glb")
    return render_template("site/project_detail.html", **context)


def _inbox_form(template, page_title, kind, target, subject_fn, success_message, endpoint):
    context = public_site_context(page_title)
    if request.method == "POST":
        try:
            content.create_contact_message(
                name=request.form.get("name"),
                email=request.form.get("email"),
                message=request.form.get("message") or subject_fn(request.form)[1],
                kind=kind,
                target=target,
                subject=subject_fn(request.form)[0],
            )
        except ServiceError as exc:
            flash_error(exc)
            return render_template(template, **context)
        flash(success_message, "success")
        return redirect(url_for(endpoint))
    return render_template(template, **context)


@bp.route("/contact", methods=["GET", "POST"])
def public_contact():
    return _inbox_form(
        "site/contact.html",
        "Socials + Contact",
        "contact",
        None,
        lambda form: ((form.get("subject") or "").strip() or None, ""),
        "Message sent. Our admin team will follow up.",
        "public.public_contact",
    )


@bp.get("/socials")
def public_socials():
    return redirect(url_for("public.public_contact"))


@bp.route("/join", methods=["GET", "POST"])
def public_join():
    context = public_site_context("Join / Get Involved")
    if request.method == "POST":
        interest = (request.form.get("interest") or "").strip()
        if not interest:
            flash("Name, email, and interest area are required.", "error")
            return render_template("site/join.html", **context)
        try:
            content.create_contact_message(
                name=request.form.get("name"),
                email=request.form.get("email"),
                message=(request.form.get("message") or "").strip() or f"Interested in: {interest}",
                kind="join",
                target="membership",
                subject=f"Join Interest: {interest[:120]}",
            )
        except ServiceError as exc:
            flash("Name, email, and interest area are required.", "error")
            return render_template("site/join.html", **context)
        flash("Interest form submitted. We will contact you with onboarding details.", "success")
        return redirect(url_for("public.public_join"))
    return render_template("site/join.html", **context)


@bp.route("/sponsors", methods=["GET", "POST"])
def public_sponsors():
    return _inbox_form(
        "site/sponsors.html",
        "Sponsors / Partners",
        "sponsor",
        "sponsorship",
        lambda form: ("Sponsorship Inquiry", ""),
        "Sponsorship inquiry sent. Thank you for supporting ASME at Iowa.",
        "public.public_sponsors",
    )
