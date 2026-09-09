"""Blueprint registration. Order matters only for URL rule precedence, which
these do not overlap."""

from __future__ import annotations


def register_blueprints(app):
    from asme.blueprints import admin, api_legacy, api_v1, auth, kiosk, legacy_ops, portal, public

    for module in (public, auth, kiosk, portal, admin, api_v1, api_legacy, legacy_ops):
        app.register_blueprint(module.bp)

    if app.config["SETTINGS"].ops_enabled:
        from asme.ops import web as ops_web
        from asme.ops.api import bp as ops_api_bp

        app.register_blueprint(ops_api_bp)
        app.register_blueprint(ops_web.bp)
