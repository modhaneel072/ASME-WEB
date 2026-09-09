"""Session, organisation, setup center, dashboard, notifications, audit."""

from __future__ import annotations

from flask import request

from asme.auth.session import rate_limiter, sign_in_user, sign_out_user
from asme.ops import authz
from asme.ops.api import body, bp, contract, ctx, ok
from asme.ops.models import Organization
from asme.ops.schemas import LoginRequest, MemberUpdate, OrgProfileUpdate, SessionOut, SetupBannerUpdate, UserInvite
from asme.ops.services import audit_log, dashboard, notifications, organizations, setup_center
from asme.ops.services.common import user_ref
from asme.ops.tenancy import load_context
from asme.services import identity
from asme.services.errors import Validation
from asme.utils import parse_positive_int
from asme.utils.http import request_client_ip


def _session_payload(c) -> dict:
    org: Organization = c.organization
    return {
        "user": user_ref(c.user, c.membership),
        "organization": {"id": org.id, "name": org.name, "slug": org.slug, "timezone": org.timezone, "academic_year_start_month": org.academic_year_start_month, "logo_url": org.logo_url, "settings": org.settings},
        "role": {"key": c.role.system_key, "name": c.role.name, "is_custom": bool(c.role.is_custom)},
        "permissions": c.permission_keys,
        "grants": c.grants,
        "setup_banner_dismissed": bool(c.membership.settings.get("setup_banner_dismissed")),
        "unread_notifications": notifications.unread_count(c),
        "title": c.membership.title,
    }


@bp.get("/ops/health")
def health():
    return ok({"service": "asme-ops", "api": "v1"})


@bp.post("/ops/session/login")
@contract(LoginRequest, SessionOut, summary="Sign in with email or username", tags=("session",))
def session_login():
    data = body(LoginRequest)
    limiter = rate_limiter()
    ip = request_client_ip()
    blocked, retry_after = limiter.is_limited(ip, data.identifier)
    if blocked:
        return ok({"retry_after": retry_after}, status=429, ok_override=False) if False else ({"ok": False, "code": "rate_limited", "error": f"Too many attempts. Try again in about {retry_after} seconds.", "retry_after": retry_after}, 429)
    user = identity.authenticate(data.identifier, data.password)
    if user is None:
        limiter.record_failure(ip, data.identifier)
        return {"ok": False, "code": "invalid_credentials", "error": "Invalid email/username or password."}, 401
    limiter.clear(ip, data.identifier)
    sign_in_user(user)
    c = load_context(user)
    if c is None:
        sign_out_user()
        return {"ok": False, "code": "membership_inactive", "error": "Your membership is not active."}, 403
    return ok(_session_payload(c))


@bp.post("/ops/session/logout")
def session_logout():
    sign_out_user()
    return ok({"signed_out": True})


@bp.get("/ops/session")
@contract(None, SessionOut, summary="Current user, organisation, role and permissions", tags=("session",))
def session_current():
    return ok(_session_payload(ctx()))


# --------------------------------------------------------------------------- organisation & members


@bp.get("/org")
def org_get():
    c = ctx()
    org = c.organization
    return ok({"id": org.id, "name": org.name, "slug": org.slug, "timezone": org.timezone, "academic_year_start_month": org.academic_year_start_month, "logo_url": org.logo_url, "settings": org.settings, "created_at": org.created_at.isoformat()})


@bp.patch("/org")
@contract(OrgProfileUpdate, summary="Update chapter profile", tags=("organization",))
@authz.permission_required("org.manage")
def org_update():
    data = body(OrgProfileUpdate)
    org = organizations.update_profile(ctx().organization, data, data.model_fields_set, ctx().user)
    setup_center.confirm_profile(ctx())
    return ok({"id": org.id, "name": org.name, "timezone": org.timezone, "academic_year_start_month": org.academic_year_start_month, "logo_url": org.logo_url})


@bp.get("/roles")
def roles_list():
    return ok([{"id": r.id, "key": r.system_key, "name": r.name, "rank": r.rank, "is_custom": bool(r.is_custom), "permissions": sorted({rp.permission.key: rp.scope_type for rp in r.permissions}.items())} for r in organizations.list_roles(ctx().organization)])


@bp.get("/users")
@authz.permission_required("team.read")
def users_list():
    include_inactive = request.args.get("include_inactive") in {"1", "true"}
    if include_inactive and not authz.can(ctx(), "user.manage"):
        include_inactive = False
    rows = organizations.list_members(ctx().organization, include_inactive=include_inactive, include_invited=True)
    return ok([organizations.serialize_member(m) for m in rows])


@bp.post("/users")
@contract(UserInvite, summary="Invite a member", tags=("users",))
@authz.permission_required("user.manage")
def users_invite():
    data = body(UserInvite)
    membership = organizations.invite_user(ctx().organization, data, ctx().user)
    return ok(organizations.serialize_member(membership), status=201)


@bp.patch("/users/<int:user_id>")
@contract(MemberUpdate, summary="Change a member's role, status or title", tags=("users",))
@authz.permission_required("user.manage")
def users_update(user_id):
    data = body(MemberUpdate)
    membership = organizations.membership_for(ctx().organization, user_id)
    membership = organizations.update_member(ctx().organization, membership, data, data.model_fields_set, ctx().user)
    return ok(organizations.serialize_member(membership))


# --------------------------------------------------------------------------- setup center


@bp.get("/ops/setup")
def setup_get():
    return ok(setup_center.evaluate(ctx()))


@bp.post("/ops/setup/banner")
@contract(SetupBannerUpdate, summary="Dismiss or restore the setup banner for the current user", tags=("setup",))
def setup_banner():
    data = body(SetupBannerUpdate)
    organizations.set_setup_banner(ctx().membership, data.dismissed)
    return ok({"dismissed": data.dismissed})


@bp.post("/ops/setup/guide-read")
def setup_guide_read():
    setup_center.mark_guide_read(ctx())
    return ok(setup_center.evaluate(ctx()))


@bp.post("/ops/setup/confirm-profile")
@authz.permission_required("setup.manage")
def setup_confirm_profile():
    setup_center.confirm_profile(ctx())
    return ok(setup_center.evaluate(ctx()))


# --------------------------------------------------------------------------- dashboard, notifications, audit


@bp.get("/ops/dashboard/operations")
@authz.permission_required("report.view")
def dashboard_operations():
    range_key = (request.args.get("range") or "30d").strip()
    project_id = (request.args.get("project_id") or "").strip() or None
    if project_id:
        from asme.ops.models import OpsProject
        from asme.ops.tenancy import get_or_404

        get_or_404(OpsProject, project_id, ctx(), "Project")
    return ok(dashboard.operations(ctx(), range_key, project_id))


@bp.get("/notifications")
def notifications_list():
    unread_only = request.args.get("unread") in {"1", "true"}
    rows = notifications.list_for(ctx(), unread_only=unread_only)
    return ok({"items": [notifications.serialize(r) for r in rows], "unread": notifications.unread_count(ctx())})


@bp.post("/notifications/read-all")
def notifications_read_all():
    return ok({"marked": notifications.mark_read(ctx())})


@bp.post("/notifications/<notification_id>/read")
def notifications_read(notification_id):
    marked = notifications.mark_read(ctx(), notification_id)
    if not marked:
        raise Validation("Notification not found or already read.", code="not_found")
    return ok({"marked": marked})


@bp.get("/ops/audit")
@authz.permission_required("audit.read")
def audit_list():
    limit = min(parse_positive_int(request.args.get("page[limit]"), default=50), 200)
    offset = max(0, parse_positive_int(request.args.get("offset"), default=0) if request.args.get("offset") else 0)
    rows, has_more = audit_log.list_events(
        ctx(),
        entity_type=(request.args.get("entity_type") or "").strip() or None,
        entity_id=(request.args.get("entity_id") or "").strip() or None,
        event_type=(request.args.get("event_type") or "").strip() or None,
        actor_id=parse_positive_int(request.args.get("actor_id"), default=0) or None,
        limit=limit,
        offset=offset,
    )
    return ok({"items": [audit_log.serialize_event(r) for r in rows], "next_offset": (offset + limit) if has_more else None})
