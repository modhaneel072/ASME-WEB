"""Helpers for HTML blueprints: translate service errors into flashes."""

from __future__ import annotations

from functools import wraps

from flask import flash

from asme.services.errors import ServiceError
from asme.utils.http import redirect_to_next


def flash_error(exc: ServiceError):
    flash(exc.message, "error")
    for extra in exc.extra.get("errors") or []:
        flash(extra, "info")


def on_service_error(default_endpoint, **values):
    """Decorator: a ``ServiceError`` becomes a flash + redirect to ``next``/default."""

    def decorator(view_func):
        @wraps(view_func)
        def wrapped(*args, **kwargs):
            try:
                return view_func(*args, **kwargs)
            except ServiceError as exc:
                flash_error(exc)
                return redirect_to_next(default_endpoint, **values)

        return wrapped

    return decorator
