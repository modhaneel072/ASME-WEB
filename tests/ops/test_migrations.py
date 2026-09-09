"""Migration 0003 applies on a fresh database and on a legacy 0002 database with rows."""

import sqlite3
import tempfile
from pathlib import Path

from flask_migrate import upgrade
from werkzeug.security import generate_password_hash

from tests.conftest import make_app

OPS_TABLES = {
    "organizations", "permissions", "roles", "role_permissions", "memberships", "teams", "team_members", "locations", "categories",
    "assets", "asset_types", "asset_type_links", "asset_status_history", "ops_projects", "ops_project_members", "milestones",
    "work_orders", "work_order_assignees", "work_order_categories", "work_order_assets", "work_order_watchers",
    "work_order_status_history", "work_order_dependencies", "work_order_counters", "time_entries", "cost_entries", "comments",
    "attachments", "notifications", "saved_filters", "audit_events",
}


def _tables(path):
    con = sqlite3.connect(path)
    try:
        return {r[0] for r in con.execute("select name from sqlite_master where type='table'")}
    finally:
        con.close()


def test_upgrade_from_0002_with_existing_users_backfills_memberships(tmp_path):
    db_path = Path(tmp_path) / "legacy.db"
    app = make_app(database_url=f"sqlite:///{db_path.as_posix()}", ops_enabled=True)
    with app.app_context():
        migrations = str(Path(app.root_path).parent / "migrations")
        upgrade(directory=migrations, revision="0002_launchpad")
        assert "work_orders" not in _tables(db_path)
        from asme.extensions import db
        from asme.models import Project, User

        db.session.add(User(name="Legacy Lead", email="lead@uiowa.edu", username="lead", password_hash=generate_password_hash("x"), role="team_leader", is_active=True))
        db.session.add(Project(slug="rover", title="Rover", summary="s", description="d"))
        db.session.commit()
        upgrade(directory=migrations)
        assert OPS_TABLES <= _tables(db_path)
        from asme.services import bootstrap

        assert bootstrap.current_revision() == "0003_ops_foundation"
        bootstrap.seed_defaults()
        from asme.ops.models import Membership, OpsProject

        membership = Membership.query.join(User, User.id == Membership.user_id).filter(User.email == "lead@uiowa.edu").first()
        assert membership is not None and membership.role.system_key == "team_lead"
        linked = OpsProject.query.filter_by(code="R").first() or OpsProject.query.first()
        assert linked is not None and linked.public_project_id is not None
        db.session.remove()


def test_fresh_database_upgrades_to_head(tmp_path):
    db_path = Path(tmp_path) / "fresh.db"
    app = make_app(database_url=f"sqlite:///{db_path.as_posix()}", ops_enabled=True)
    with app.app_context():
        upgrade(directory=str(Path(app.root_path).parent / "migrations"))
        assert OPS_TABLES <= _tables(db_path)
        from asme.extensions import db

        db.session.remove()
