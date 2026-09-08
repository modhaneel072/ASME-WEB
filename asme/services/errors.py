"""Service-level errors carry a machine code, a human message and an HTTP status.

Blueprints translate them: HTML routes ``flash(err.message)``; JSON routes return
``{"ok": false, "code": err.code, "error": err.message}`` with ``err.status``.
"""


class ServiceError(Exception):
    def __init__(self, message: str, code: str = "bad_request", status: int = 400, **extra):
        super().__init__(message)
        self.message = message
        self.code = code
        self.status = status
        self.extra = extra

    def to_dict(self) -> dict:
        body = {"ok": False, "code": self.code, "error": self.message}
        body.update(self.extra)
        return body


class NotFound(ServiceError):
    def __init__(self, message="Not found.", code="not_found", **extra):
        super().__init__(message, code=code, status=404, **extra)


class Conflict(ServiceError):
    def __init__(self, message="Conflict.", code="conflict", **extra):
        super().__init__(message, code=code, status=409, **extra)


class Forbidden(ServiceError):
    def __init__(self, message="Not allowed.", code="forbidden", **extra):
        super().__init__(message, code=code, status=403, **extra)


class Validation(ServiceError):
    def __init__(self, message, code="validation", field=None, **extra):
        if field:
            extra["field"] = field
        super().__init__(message, code=code, status=400, **extra)
