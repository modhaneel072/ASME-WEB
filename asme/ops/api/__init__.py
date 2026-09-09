"""``/api/v1`` ops surface: one blueprint, typed errors, tenant context per request."""

from __future__ import annotations

import logging

from flask import Blueprint, g, jsonify, request
from pydantic import ValidationError

from asme.auth.session import current_auth_user
from asme.ops import authz
from asme.ops.tenancy import load_context
from asme.services.errors import ServiceError

log = logging.getLogger("asme.ops.api")

bp = Blueprint("ops_api", __name__, url_prefix="/api/v1")

PUBLIC_ENDPOINTS = {"ops_api.session_login", "ops_api.openapi", "ops_api.file_download", "ops_api.health"}
MUTATING = {"POST", "PUT", "PATCH", "DELETE"}
CSRF_HEADER = "X-Requested-With"
CSRF_VALUE = "ASME-Ops"


def _error(code: str, message: str, status: int, **extra):
    body = {"ok": False, "code": code, "error": message}
    body.update(extra)
    return jsonify(body), status


@bp.before_request
def _load_ops_context():
    if request.endpoint in PUBLIC_ENDPOINTS:
        return None
    if request.method in MUTATING and request.headers.get(CSRF_HEADER) != CSRF_VALUE:
        return _error("csrf", f"Send the {CSRF_HEADER}: {CSRF_VALUE} header on state-changing requests.", 403)
    user = current_auth_user()
    if user is None:
        return _error("login_required", "Login required.", 401)
    ctx = load_context(user)
    if ctx is None:
        return _error("membership_inactive", "Your membership is not active.", 403)
    g.ops_ctx = ctx
    return None


@bp.errorhandler(ServiceError)
def _service_error(exc: ServiceError):
    return jsonify(exc.to_dict()), exc.status


@bp.errorhandler(authz.Forbidden)
def _forbidden(exc: authz.Forbidden):
    return authz.forbidden_response(exc)


@bp.errorhandler(ValidationError)
def _validation(exc: ValidationError):
    details = []
    for err in exc.errors():
        loc = ".".join(str(p) for p in err.get("loc", ()) if p != "body")
        details.append({"field": loc or None, "message": err.get("msg", "Invalid value."), "type": err.get("type")})
    return _error("validation", "Check the highlighted fields.", 400, details=details)


@bp.errorhandler(404)
def _not_found(_exc):
    return _error("not_found", "Not found.", 404)


@bp.errorhandler(405)
def _method_not_allowed(_exc):
    return _error("method_not_allowed", "Method not allowed.", 405)


@bp.errorhandler(413)
def _too_large(_exc):
    return _error("file_size", "The upload is too large.", 413)


@bp.errorhandler(Exception)
def _unexpected(exc: Exception):
    log.exception("unhandled ops api error")
    return _error("server_error", "Something went wrong on our side. The error has been logged.", 500, request_id=getattr(g, "request_id", None))


def ctx() -> authz.AuthzContext:
    return g.ops_ctx


def body(model):
    payload = request.get_json(force=True, silent=True)
    if payload is None:
        payload = {}
    if not isinstance(payload, dict):
        raise ValidationError.from_exception_data("body", [{"type": "dict_type", "loc": ("body",), "input": payload}])
    return model.model_validate(payload)


def ok(payload=None, status=200, **extra):
    out = {"ok": True}
    if payload is not None:
        out["payload"] = payload
    out.update(extra)
    return jsonify(out), status


def contract(request_model=None, response_model=None, summary: str | None = None, tags: tuple[str, ...] = ()):
    """Annotate a view with its contract so the OpenAPI document can describe it."""

    def decorator(fn):
        fn._contract = {"request": request_model, "response": response_model, "summary": summary or (fn.__doc__ or "").strip().splitlines()[0] if (summary or fn.__doc__) else fn.__name__, "tags": list(tags)}
        return fn

    return decorator


from asme.ops.api import assets, directory, files, openapi, projects, saved_filters, session, work_orders  # noqa: E402,F401
