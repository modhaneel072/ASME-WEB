"""Session-backed authentication, role checks and entitlement gates.

Two orthogonal axes are checked on protected routes:

* ``role`` - what job the person holds (``member < team_leader < admin``).
* ``entitlement`` - what they have proven, granted by the Launchpad engine.

``require_entitlement`` runs in *shadow mode* until ``ASME_ONBOARDING_ENFORCE=1``:
it records what it would have blocked and lets the request through.
"""

from __future__ import annotations

import logging
import time as time_module
from functools import wraps

from flask import current_app, flash, g, has_request_context, jsonify, redirect, request, session, url_for
from sqlalchemy import func

from asme.config import settings
from asme.constants import ROLE_ORDER
from asme.extensions import db
from asme.models import Member, User

log = logging.getLogger("asme.auth")

AUTH_SESSION_BOOT_KEY = "auth_boot_token"
AUTH_SESSION_LAST_SEEN_KEY = "auth_last_seen_ts"
AUTH_SESSION_LOGIN_TS_KEY = "auth_login_ts"


# --------------------------------------------------------------------------- roles


def normalize_role(role):
    cleaned = (role or "").strip().lower()
    if cleaned in ROLE_ORDER:
        return cleaned
    return "member"


def role_allows(user_role, required_role):
    return ROLE_ORDER.get(normalize_role(user_role), 0) >= ROLE_ORDER.get(normalize_role(required_role), 0)


# --------------------------------------------------------------------------- current user


def _cache_user(user):
    if not has_request_context():
        return
    g._current_auth_user_cached = user
    g._current_auth_user_resolved = True


def current_auth_user():
    """Resolve the logged-in ``User`` (or ``None``), enforcing boot-token and idle rules."""
    if has_request_context() and getattr(g, "_current_auth_user_resolved", False):
        return getattr(g, "_current_auth_user_cached", None)

    cfg = settings()
    user_id = session.get("auth_user_id")
    if not user_id:
        _cache_user(None)
        return None
    if (session.get(AUTH_SESSION_BOOT_KEY) or "").strip() != cfg.app_boot_token:
        sign_out_user()
        return None
    try:
        user = db.session.get(User, int(user_id))
    except Exception:
        sign_out_user()
        return None
    if not user or not user.is_active:
        sign_out_user()
        return None

    now_ts = int(time_module.time())
    try:
        last_seen_ts = int(session.get(AUTH_SESSION_LAST_SEEN_KEY) or 0)
    except Exception:
        last_seen_ts = 0
    idle_minutes = cfg.admin_session_idle_minutes if role_allows(user.role, "admin") else cfg.member_session_idle_minutes
    if idle_minutes > 0 and last_seen_ts and (now_ts - last_seen_ts) > (idle_minutes * 60):
        sign_out_user()
        return None

    session[AUTH_SESSION_BOOT_KEY] = cfg.app_boot_token
    session[AUTH_SESSION_LAST_SEEN_KEY] = now_ts
    _cache_user(user)
    return user


def current_user_member(user=None):
    """Legacy ``Member`` profile linked to a user (auto-links by email once)."""
    user = user or current_auth_user()
    if not user:
        return None
    if user.member:
        return user.member
    if user.email:
        linked = Member.query.filter(func.lower(Member.email) == user.email.strip().lower()).first()
        if linked:
            user.member_id = linked.id
            db.session.commit()
            return linked
    return None


def get_active_member():
    """Kiosk-style "active member" chosen on a shared device (legacy ops flows)."""
    member_id = session.get("active_member_id")
    if not member_id:
        return None
    try:
        return db.session.get(Member, int(member_id))
    except Exception:
        return None


def sign_in_user(user):
    from datetime import datetime

    cfg = settings()
    session.clear()
    session.permanent = False
    now_ts = int(time_module.time())
    session["auth_user_id"] = user.id
    session["auth_user_role"] = normalize_role(user.role)
    session[AUTH_SESSION_BOOT_KEY] = cfg.app_boot_token
    session[AUTH_SESSION_LOGIN_TS_KEY] = now_ts
    session[AUTH_SESSION_LAST_SEEN_KEY] = now_ts
    member = current_user_member(user)
    if member:
        session["active_member_id"] = member.id
    else:
        session.pop("active_member_id", None)
    _cache_user(user)
    user.last_login_at = datetime.utcnow()
    db.session.commit()


def sign_out_user():
    session.clear()
    _cache_user(None)


# --------------------------------------------------------------------------- decorators


def _wants_json():
    if request.path.startswith("/api/"):
        return True
    accept = request.headers.get("Accept", "")
    return request.is_json or ("application/json" in accept and "text/html" not in accept)


def _deny_login():
    if _wants_json():
        return jsonify({"ok": False, "code": "login_required", "error": "Login required."}), 401
    flash("Please log in to continue.", "error")
    return redirect(url_for("auth.login_page", next=request.path))


def _deny_role():
    if _wants_json():
        return jsonify({"ok": False, "code": "forbidden", "error": "You do not have permission to do that."}), 403
    flash("You do not have permission to view that page.", "error")
    return redirect(url_for("portal.portal_router"))


def require_login(view_func):
    @wraps(view_func)
    def wrapped(*args, **kwargs):
        if not current_auth_user():
            return _deny_login()
        return view_func(*args, **kwargs)

    return wrapped


def require_role(min_role):
    def decorator(view_func):
        @wraps(view_func)
        def wrapped(*args, **kwargs):
            user = current_auth_user()
            if not user:
                return _deny_login()
            if not role_allows(user.role, min_role):
                return _deny_role()
            return view_func(*args, **kwargs)

        return wrapped

    return decorator


def entitlement_denied_response(key, user):
    """403 payload that tells the client which Launchpad phase unlocks ``key``."""
    from asme.services.onboarding import entitlements as ent

    phase = ent.phase_that_grants(key)
    detail = {
        "ok": False,
        "code": "entitlement_required",
        "entitlement": key,
        "phase": phase.key if phase else None,
        "phase_name": phase.name if phase else None,
        "error": f"You need '{key}' to do that. Complete the Launchpad phase that unlocks it.",
    }
    if _wants_json():
        return jsonify(detail), 403
    if phase:
        flash(f"Finish the \"{phase.name}\" phase in Launchpad to unlock that.", "error")
    else:
        flash("You are not yet cleared to do that. Check Launchpad for the next step.", "error")
    return redirect(url_for("portal.member_launchpad"))


def require_entitlement(key):
    """Gate a route on a Launchpad entitlement. Admins always pass.

    In shadow mode (``ASME_ONBOARDING_ENFORCE=0``) the check is evaluated and
    recorded on ``g.shadow_entitlement_blocks`` but the request proceeds.
    """

    def decorator(view_func):
        @wraps(view_func)
        def wrapped(*args, **kwargs):
            from asme.services.onboarding import entitlements as ent

            user = current_auth_user()
            if not user:
                return _deny_login()
            if ent.has_entitlement(user, key):
                return view_func(*args, **kwargs)
            if settings().onboarding_enforce:
                return entitlement_denied_response(key, user)
            blocks = getattr(g, "shadow_entitlement_blocks", None)
            if blocks is None:
                blocks = g.shadow_entitlement_blocks = []
            blocks.append(key)
            log.info("shadow-block entitlement=%s user_id=%s path=%s", key, user.id, request.path)
            return view_func(*args, **kwargs)

        return wrapped

    return decorator


def is_admin_member(member):
    """Legacy kiosk admin check based on roster class / configured admin emails."""
    if not member:
        return False
    admin_emails = settings().admin_emails
    if admin_emails and member.email and member.email.strip().lower() in admin_emails:
        return True
    role = (member.member_class or "").strip().lower()
    return any(marker in role for marker in ("admin", "officer", "lead", "president", "chair"))


def rate_limiter():
    return current_app.extensions["asme_login_rate_limiter"]
