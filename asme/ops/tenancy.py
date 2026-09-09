"""Organisation resolution and tenant-safe lookups.

Every ops request runs through ``load_context`` which resolves the signed-in user's
membership in the default organisation (creating one from the legacy role on first
use) and stores an ``AuthzContext`` on ``g``. ``get_or_404`` is the only way services
fetch a record by id: it always filters by the caller's organisation, so an id from
another tenant is indistinguishable from a missing one.
"""

from __future__ import annotations

from flask import g

from asme.extensions import db
from asme.ops import authz
from asme.ops.models import Membership, Organization, Role
from asme.services.errors import NotFound

DEFAULT_ORG_SLUG = "asme-uiowa"
DEFAULT_ORG_NAME = "American Society of Mechanical Engineers at the University of Iowa"


def default_organization() -> Organization:
    org = Organization.query.filter_by(slug=DEFAULT_ORG_SLUG).first()
    if org is None:
        org = Organization(name=DEFAULT_ORG_NAME, slug=DEFAULT_ORG_SLUG, timezone="America/Chicago")
        db.session.add(org)
        db.session.flush()
    return org


def ensure_membership(user, organization: Organization | None = None) -> Membership:
    """Membership for ``user``; created from the legacy ``users.role`` when missing."""
    organization = organization or default_organization()
    membership = Membership.query.filter_by(organization_id=organization.id, user_id=user.id).first()
    if membership is not None:
        return membership
    roles = authz.ensure_system_roles(organization)
    system_key = authz.LEGACY_ROLE_MAP.get((user.role or "member").strip().lower(), "full_member")
    membership = Membership(
        organization_id=organization.id,
        user_id=user.id,
        role_id=roles[system_key].id,
        member_status="active" if user.is_active else "suspended",
    )
    db.session.add(membership)
    db.session.flush()
    return membership


def load_context(user) -> authz.AuthzContext | None:
    if user is None:
        g.ops_ctx = None
        return None
    cached = getattr(g, "ops_ctx", None)
    if cached is not None and cached.user.id == user.id:
        return cached
    organization = default_organization()
    membership = ensure_membership(user, organization)
    if membership.member_status != "active":
        g.ops_ctx = None
        return None
    db.session.commit()
    ctx = authz.build_context(user, organization, membership)
    g.ops_ctx = ctx
    return ctx


def ctx_or_none() -> authz.AuthzContext | None:
    return getattr(g, "ops_ctx", None)


def org_id() -> str:
    ctx = ctx_or_none()
    if ctx is None:
        raise RuntimeError("no ops context on this request")
    return ctx.organization.id


def scoped(model, ctx: authz.AuthzContext | None = None):
    """``Model.query`` filtered to the caller's organisation."""
    ctx = ctx or ctx_or_none()
    return model.query.filter(model.organization_id == ctx.organization.id)


def get_or_404(model, record_id, ctx: authz.AuthzContext | None = None, label: str | None = None):
    ctx = ctx or ctx_or_none()
    if not record_id:
        raise NotFound(f"{label or model.__name__} not found.")
    row = model.query.filter(model.id == record_id, model.organization_id == ctx.organization.id).first()
    if row is None:
        raise NotFound(f"{label or model.__name__} not found.")
    return row


def role_by_key(organization: Organization, system_key: str) -> Role | None:
    return Role.query.filter_by(organization_id=organization.id, system_key=system_key).first()
