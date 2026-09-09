"""Operational commands.

    python manage.py upgrade      # migrate to head (stamps legacy DBs) + seed defaults
    python manage.py seed         # seed defaults only
    python manage.py seed-demo    # development: Crater Cruncher Rover demo data (ASME Ops)
    python manage.py serve-e2e    # throwaway seeded server for Playwright
    python manage.py evaluate     # re-run the Launchpad engine for everyone
    python manage.py reconcile    # file stock discrepancies
    python manage.py worker       # dedicated outbox worker loop
    python manage.py serve        # upgrade, then run the dev server
    python manage.py routes       # list URL rules
"""

from __future__ import annotations

import sys
import time

from asme import create_app


def serve_e2e():
    """Fresh SQLite database + demo seed, served on ASME_E2E_PORT (default 5055) for Playwright."""
    import os
    import tempfile
    from pathlib import Path

    port = int(os.environ.get("ASME_E2E_PORT", "5055"))
    db_path = Path(tempfile.gettempdir()) / "asme_e2e.db"
    if db_path.exists():
        db_path.unlink()
    app = create_app(
        env="testing", secret_key="e2e-secret", database_url=f"sqlite:///{db_path.as_posix()}", auto_migrate=False,
        outbox_worker_enabled=False, onboarding_enforce=False, session_cookie_secure=False, login_rate_max_attempts=1000,
    )
    with app.app_context():
        from asme.services import bootstrap
        from asme.ops.seeds import seed_ops_demo

        bootstrap.sync_schema()
        bootstrap.seed_defaults()
        print("demo:", seed_ops_demo(), flush=True)
    print(f"e2e server ready on http://127.0.0.1:{port}", flush=True)
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)
    return 0


def main(argv):
    command = (argv[1] if len(argv) > 1 else "help").strip().lower()
    if command in {"help", "-h", "--help"}:
        print(__doc__)
        return 0

    if command == "serve-e2e":
        return serve_e2e()

    # Never spin up the background worker for one-shot commands.
    app = create_app(outbox_worker_enabled=(command in {"serve"}), auto_migrate=False)
    with app.app_context():
        from asme.services import bootstrap

        if command == "upgrade":
            print("schema:", bootstrap.sync_schema())
            print("seed:", bootstrap.seed_defaults())
            return 0
        if command == "seed":
            print("seed:", bootstrap.seed_defaults())
            return 0
        if command == "seed-demo":
            from asme.ops.seeds import seed_ops_demo

            print("schema:", bootstrap.sync_schema())
            print("seed:", bootstrap.seed_defaults())
            print("demo:", seed_ops_demo())
            return 0
        if command == "evaluate":
            from asme.models import User
            from asme.services.onboarding import engine

            users = User.query.filter(User.is_active.is_(True)).all()
            for user in users:
                engine.evaluate_user(user)
            engine.evaluate_chapter()
            print(f"evaluated {len(users)} users + chapter")
            return 0
        if command == "reconcile":
            from asme.services import inventory

            print("discrepancies filed:", len(inventory.reconcile_stock()))
            return 0
        if command == "worker":
            from asme.jobs import process_pending

            poll = app.config["SETTINGS"].outbox_poll_seconds
            print(f"outbox worker running (poll {poll}s); Ctrl+C to stop")
            while True:
                done = process_pending()
                if done:
                    print("processed", done)
                time.sleep(poll)
        if command == "routes":
            for rule in sorted(app.url_map.iter_rules(), key=lambda r: (r.rule, r.endpoint)):
                methods = ",".join(sorted(m for m in rule.methods if m not in {"HEAD", "OPTIONS"}))
                print(f"{methods:12} {rule.rule:55} {rule.endpoint}")
            return 0
        if command == "serve":
            print("schema:", bootstrap.sync_schema())
            print("seed:", bootstrap.seed_defaults())
    if command == "serve":
        settings = app.config["SETTINGS"]
        app.run(host="0.0.0.0", port=settings.port, debug=settings.env == "development")
        return 0

    print(f"unknown command: {command}")
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
