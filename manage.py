"""Operational commands.

    python manage.py upgrade      # migrate to head (stamps legacy DBs) + seed defaults
    python manage.py seed         # seed defaults only
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


def main(argv):
    command = (argv[1] if len(argv) > 1 else "help").strip().lower()
    if command in {"help", "-h", "--help"}:
        print(__doc__)
        return 0

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
