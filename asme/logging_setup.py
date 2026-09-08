"""Structured request logging: one JSON line per request."""

from __future__ import annotations

import json
import logging
import sys
import time
import uuid

from flask import g, request, session


class JsonFormatter(logging.Formatter):
    def format(self, record):
        payload = {
            "ts": self.formatTime(record, "%Y-%m-%dT%H:%M:%S"),
            "level": record.levelname,
            "logger": record.name,
            "msg": record.getMessage(),
        }
        extra = getattr(record, "asme", None)
        if isinstance(extra, dict):
            payload.update(extra)
        if record.exc_info:
            payload["exc"] = self.formatException(record.exc_info)
        return json.dumps(payload, default=str)


def configure_logging(app):
    root = logging.getLogger()
    if not any(getattr(h, "_asme_json", False) for h in root.handlers):
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(JsonFormatter())
        handler._asme_json = True
        root.addHandler(handler)
    level = logging.DEBUG if app.debug else logging.INFO
    root.setLevel(level)
    logging.getLogger("asme").setLevel(level)
    logging.getLogger("werkzeug").setLevel(logging.WARNING)

    request_log = logging.getLogger("asme.request")

    @app.before_request
    def _start_timer():
        g._req_started = time.perf_counter()
        g.request_id = (request.headers.get("X-Request-Id") or uuid.uuid4().hex[:12]).strip()[:64]

    @app.after_request
    def _log_request(response):
        if request.path.startswith("/static/") or request.path in {"/healthz"}:
            response.headers["X-Request-Id"] = getattr(g, "request_id", "")
            return response
        started = getattr(g, "_req_started", None)
        duration_ms = round((time.perf_counter() - started) * 1000, 1) if started else None
        request_log.info(
            "request",
            extra={
                "asme": {
                    "request_id": getattr(g, "request_id", None),
                    "method": request.method,
                    "path": request.path,
                    "status": response.status_code,
                    "duration_ms": duration_ms,
                    "user_id": session.get("auth_user_id"),
                    "endpoint": request.endpoint,
                    "shadow_blocks": getattr(g, "shadow_entitlement_blocks", None),
                }
            },
        )
        response.headers["X-Request-Id"] = getattr(g, "request_id", "")
        return response
