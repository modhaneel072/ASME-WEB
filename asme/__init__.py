"""ASME @ UIowa web platform - application factory."""

from __future__ import annotations

import logging
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, g, request, session

from asme.auth.rate_limit import LoginRateLimiter
from asme.config import Settings, load_instance_env_file
from asme.extensions import db, migrate

ROOT = Path(__file__).resolve().parent.parent
log = logging.getLogger("asme")

__version__ = "2.0.0"


def create_app(settings: Settings | None = None, **overrides) -> Flask:
    """Build the Flask app.

    ``overrides`` are applied on top of the environment-derived settings, which
    is how tests get an in-memory database and no background worker.
    """
    load_dotenv(ROOT / ".env")
    instance_path = ROOT / "instance"
    instance_path.mkdir(parents=True, exist_ok=True)
    load_instance_env_file(instance_path / "print_commands.env")

    cfg = settings or Settings.from_env(**overrides)

    app = Flask(
        __name__,
        template_folder=str(ROOT / "templates"),
        static_folder=str(ROOT / "static"),
        instance_path=str(instance_path),
    )
    app.config.update(
        SETTINGS=cfg,
        SECRET_KEY=cfg.secret_key,
        SQLALCHEMY_DATABASE_URI=cfg.database_url,
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        SESSION_COOKIE_SAMESITE=cfg.session_cookie_samesite,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SECURE=cfg.session_cookie_secure,
        SESSION_PERMANENT=False,
        TEMPLATES_AUTO_RELOAD=cfg.templates_auto_reload,
        TESTING=cfg.is_testing,
        MAX_CONTENT_LENGTH=max(cfg.print_max_upload_bytes, 64 * 1024 * 1024),
    )
    app.jinja_env.auto_reload = cfg.templates_auto_reload

    problems = cfg.validate()
    if problems:
        if cfg.is_production:
            raise RuntimeError("Refusing to start with invalid configuration:\n- " + "\n- ".join(problems))
        for problem in problems:
            log.warning("config: %s", problem)
    for note in cfg.warnings():
        log.warning("config: %s", note)

    # -- extensions ------------------------------------------------------------
    db.init_app(app)
    migrate.init_app(app, db, directory=str(ROOT / "migrations"), render_as_batch=True)
    app.extensions["asme_login_rate_limiter"] = LoginRateLimiter(cfg.login_rate_window_seconds, cfg.login_rate_max_attempts)

    from asme.integrations.calendar import build_provider

    app.extensions["asme_calendar_provider"] = build_provider(cfg)

    # -- wiring ----------------------------------------------------------------
    from asme.blueprints import register_blueprints
    from asme.logging_setup import configure_logging
    from asme.services.onboarding import register_event_handlers
    import asme.jobs  # noqa: F401  (registers outbox handlers)

    configure_logging(app)
    register_blueprints(app)
    register_event_handlers()
    _register_request_hooks(app, cfg)
    _register_cli(app)

    # -- schema / seeds / worker ----------------------------------------------
    if cfg.auto_migrate and not cfg.is_testing:
        from asme.services import bootstrap

        with app.app_context():
            try:
                outcome = bootstrap.sync_schema()
                bootstrap.seed_defaults()
                log.info("schema sync: %s", outcome)
            except Exception:
                log.exception("automatic schema sync failed; run `python manage.py upgrade`")

    return app


def _register_request_hooks(app: Flask, cfg: Settings):
    from asme.auth.session import current_auth_user

    skip_prefixes = ("/static/", "/arm-sim")

    @app.before_request
    def _start_worker_once():
        if cfg.outbox_worker_enabled and not cfg.is_testing and not app.extensions.get("asme_worker_started"):
            from asme.jobs import start_worker

            start_worker(app)
            app.extensions["asme_worker_started"] = True

    @app.before_request
    def _enforce_auth_session_guardrails():
        if request.path.startswith(skip_prefixes) or request.path == "/healthz":
            return None
        current_auth_user()
        return None

    @app.after_request
    def _no_cache_for_authenticated(response):
        if session.get("auth_user_id"):
            response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0, private"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
            existing_vary = response.headers.get("Vary", "")
            if "Cookie" not in existing_vary:
                response.headers["Vary"] = f"{existing_vary}, Cookie".strip(", ")
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
        response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
        return response

    @app.teardown_appcontext
    def _shutdown_session(_exc):
        db.session.remove()

    @app.context_processor
    def _inject_globals():
        return {"app_version": __version__, "onboarding_enforced": cfg.onboarding_enforce}


def _register_cli(app: Flask):
    import click

    @app.cli.command("upgrade-db")
    def upgrade_db():
        """Bring the schema to head (stamps legacy databases first) and seed defaults."""
        from asme.services import bootstrap

        click.echo(f"schema: {bootstrap.sync_schema()}")
        click.echo(f"seed: {bootstrap.seed_defaults()}")

    @app.cli.command("seed")
    def seed():
        """Seed default projects, announcement, admin and Launchpad tracks."""
        from asme.services import bootstrap

        click.echo(bootstrap.seed_defaults())

    @app.cli.command("evaluate-launchpad")
    def evaluate_launchpad():
        """Re-run the onboarding engine for every active user and the chapter."""
        from asme.models import User
        from asme.services.onboarding import engine

        count = 0
        for user in User.query.filter(User.is_active.is_(True)).all():
            engine.evaluate_user(user)
            count += 1
        engine.evaluate_chapter()
        click.echo(f"evaluated {count} users + chapter")

    @app.cli.command("reconcile-stock")
    def reconcile_stock():
        from asme.services import inventory

        filed = inventory.reconcile_stock()
        click.echo(f"discrepancies filed: {len(filed)}")

    @app.cli.command("run-jobs")
    @click.option("--loop/--once", default=False)
    def run_jobs(loop):
        """Process outbox jobs once, or loop forever (use for a dedicated worker process)."""
        import time

        from asme.jobs import process_pending

        while True:
            done = process_pending()
            click.echo(f"processed {done}")
            if not loop:
                break
            time.sleep(app.config["SETTINGS"].outbox_poll_seconds)
