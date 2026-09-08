"""Request/response helpers used by blueprints."""

from __future__ import annotations

from flask import jsonify, redirect, request, url_for


def request_client_ip():
    forwarded = (request.headers.get("X-Forwarded-For") or "").strip()
    if forwarded:
        return forwarded.split(",")[0].strip()[:120]
    return (request.remote_addr or "unknown")[:120]


def value_from_request(key, default=None):
    """Read a field from JSON body or form data, whichever the client sent."""
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        return payload.get(key, default)
    return request.form.get(key, default)


def safe_next_url(raw):
    value = (raw or "").strip()
    if value.startswith("/") and not value.startswith("//"):
        return value
    return ""


def redirect_to_next(default_endpoint, **values):
    next_url = safe_next_url(request.form.get("next") or request.args.get("next"))
    if next_url:
        return redirect(next_url)
    return redirect(url_for(default_endpoint, **values))


def api_error(message, status=400, code="bad_request", **extra):
    body = {"ok": False, "error": message, "code": code}
    if extra:
        body.update(extra)
    return jsonify(body), status


def api_ok(payload=None, message=None, status=200):
    body = {"ok": True}
    if message is not None:
        body["message"] = message
    if payload is not None:
        body["payload"] = payload
    return jsonify(body), status


def idempotency_key():
    value = (request.headers.get("Idempotency-Key") or "").strip()
    if not value and request.is_json:
        value = str((request.get_json(silent=True) or {}).get("idempotency_key") or "").strip()
    if not value:
        value = (request.form.get("idempotency_key") or "").strip()
    return value[:120] or None
