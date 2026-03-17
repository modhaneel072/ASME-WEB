from functools import wraps

from flask import flash, redirect, request, url_for
from flask_login import LoginManager, current_user, login_required

login_manager = LoginManager()
login_manager.login_view = "login"
login_manager.login_message = "Please sign in to continue."
login_manager.login_message_category = "error"


def init_auth(app):
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        from models import Member

        try:
            member = Member.query.get(int(user_id))
        except (TypeError, ValueError):
            return None

        if member is None or not member.is_active:
            return None
        return member

    @login_manager.unauthorized_handler
    def unauthorized():
        next_url = request.full_path if request.query_string else request.path
        endpoint = "admin_portal_entry" if request.path.startswith("/admin") else "member_portal_entry"
        return redirect(url_for(endpoint, next=next_url))


def member_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        return f(*args, **kwargs)

    return decorated


def admin_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if current_user.role != "admin":
            flash("Admin access required.", "error")
            return redirect(url_for("member_dashboard"))
        return f(*args, **kwargs)

    return decorated


def elevated_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if current_user.role not in ("team_lead", "project_manager", "admin"):
            flash("Elevated access required.", "error")
            return redirect(url_for("member_dashboard"))
        return f(*args, **kwargs)

    return decorated
