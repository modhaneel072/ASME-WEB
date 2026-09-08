"""Login, signup, password reset, admin login."""

from __future__ import annotations

from flask import Blueprint, flash, redirect, render_template, request, url_for

from asme.auth.session import current_auth_user, rate_limiter, role_allows, sign_in_user, sign_out_user
from asme.blueprints._context import public_site_context
from asme.blueprints._helpers import flash_error
from asme.config import settings
from asme.models import User
from asme.services import identity
from asme.services.errors import ServiceError
from asme.utils.http import request_client_ip, safe_next_url
from sqlalchemy import func

bp = Blueprint("auth", __name__)


@bp.route("/signup", methods=["GET", "POST"])
def signup_page():
    context = public_site_context("Sign Up")
    if request.method == "POST":
        try:
            identity.signup(
                name=request.form.get("name"),
                email=request.form.get("email"),
                password=request.form.get("password") or "",
                role=request.form.get("role") or "member",
            )
        except ServiceError as exc:
            flash_error(exc)
            return render_template("site/signup.html", **context)
        flash("Account created. You can now log in.", "success")
        return redirect(url_for("auth.login_page"))
    return render_template("site/signup.html", **context)


@bp.route("/login", methods=["GET", "POST"])
def login_page():
    context = public_site_context("Login")
    next_url = safe_next_url(request.args.get("next") or request.form.get("next"))
    context["next_url"] = next_url
    if request.method == "POST":
        identifier = (request.form.get("identifier") or request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""
        limiter = rate_limiter()
        ip = request_client_ip()
        blocked, retry_after = limiter.is_limited(ip, identifier)
        if blocked:
            flash(f"Too many login attempts. Try again in about {retry_after} seconds.", "error")
            return render_template("site/login.html", **context)
        user = identity.authenticate(identifier, password)
        if not user:
            limiter.record_failure(ip, identifier)
            flash("Invalid email or password.", "error")
            return render_template("site/login.html", **context)
        limiter.clear(ip, identifier)
        sign_in_user(user)
        if next_url:
            return redirect(next_url)
        return redirect(url_for("portal.portal_router"))
    return render_template("site/login.html", **context)


@bp.route("/admin-login", methods=["GET", "POST"])
def admin_login_page():
    cfg = settings()
    context = public_site_context("Admin Login")
    next_url = safe_next_url(request.args.get("next") or request.form.get("next"))
    context["next_url"] = next_url
    context["shared_admin_email"] = cfg.shared_admin_email
    current = current_auth_user()
    if current and role_allows(current.role, "admin"):
        return redirect(url_for("admin.dashboard"))

    if request.method == "POST":
        identifier = (request.form.get("identifier") or request.form.get("email") or "").strip().lower()
        if not identifier and cfg.shared_admin_email:
            identifier = cfg.shared_admin_email
        password = request.form.get("password") or ""
        limiter = rate_limiter()
        ip = request_client_ip()
        blocked, retry_after = limiter.is_limited(ip, identifier)
        if blocked:
            flash(f"Too many login attempts. Try again in about {retry_after} seconds.", "error")
            return render_template("site/admin_login.html", **context)

        # One shared admin login from environment variables.
        if cfg.shared_admin_email and cfg.shared_admin_password and identifier == cfg.shared_admin_email and password == cfg.shared_admin_password:
            user = identity.ensure_shared_admin_user(cfg.shared_admin_email, cfg.shared_admin_password)
            if user:
                limiter.clear(ip, identifier)
                sign_in_user(user)
                return redirect(next_url or url_for("admin.dashboard"))

        user = identity.authenticate(identifier, password)
        if not user:
            limiter.record_failure(ip, identifier)
            flash("Invalid admin credentials.", "error")
            return render_template("site/admin_login.html", **context)
        if not role_allows(user.role, "admin"):
            limiter.record_failure(ip, identifier)
            flash("This account is not an admin account.", "error")
            return render_template("site/admin_login.html", **context)
        limiter.clear(ip, identifier)
        sign_in_user(user)
        return redirect(next_url or url_for("admin.dashboard"))
    return render_template("site/admin_login.html", **context)


@bp.get("/admin")
def admin_home():
    current = current_auth_user()
    if current and role_allows(current.role, "admin"):
        return redirect(url_for("admin.dashboard"))
    return redirect(url_for("auth.admin_login_page"))


@bp.get("/logout")
@bp.post("/logout")
def logout_page():
    sign_out_user()
    flash("You have been logged out.", "info")
    response = redirect(url_for("public.public_home"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0, private"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@bp.route("/forgot-password", methods=["GET", "POST"])
def forgot_password_page():
    context = public_site_context("Forgot Password")
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        user = User.query.filter(func.lower(User.email) == email).first()
        if user and user.is_active:
            token = identity.create_password_reset(user)
            reset_link = url_for("auth.reset_password_page", token=token, _external=True)
            flash(f"Reset link generated: {reset_link}", "info")
        else:
            flash("If this email exists, a reset link has been generated.", "info")
        return redirect(url_for("auth.forgot_password_page"))
    return render_template("site/forgot_password.html", **context)


@bp.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password_page(token):
    context = public_site_context("Reset Password")
    reset_row = identity.find_valid_reset(token)
    if not reset_row:
        flash("Reset link is invalid or expired.", "error")
        return redirect(url_for("auth.forgot_password_page"))
    if request.method == "POST":
        try:
            identity.consume_password_reset(reset_row, request.form.get("password") or "", request.form.get("confirm_password") or "")
        except ServiceError as exc:
            flash_error(exc)
            if exc.status == 404:
                return redirect(url_for("auth.forgot_password_page"))
            return render_template("site/reset_password.html", **context)
        flash("Password updated. You can now log in.", "success")
        return redirect(url_for("auth.login_page"))
    return render_template("site/reset_password.html", **context)
