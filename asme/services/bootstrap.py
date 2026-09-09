"""Schema sync and default data.

* ``sync_schema`` - bring the database to the current Alembic head. A database
  created by the old ``db.create_all()`` code path (no ``alembic_version``) is
  stamped at the baseline revision first, then upgraded.
* ``seed_defaults`` - idempotent seed of projects, the welcome announcement,
  the bootstrap admin, and the Launchpad tracks.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime

from flask import current_app
from sqlalchemy import func, inspect, text
from werkzeug.security import generate_password_hash

from asme.config import settings
from asme.extensions import db
from asme.models import Announcement, Member, Project, User
from asme.services.onboarding.seeds import seed_default_tracks

log = logging.getLogger("asme.bootstrap")

BASELINE_REVISION = "0001_baseline"


def _has_alembic_version() -> bool:
    return "alembic_version" in inspect(db.engine).get_table_names()


def _is_legacy_database() -> bool:
    tables = set(inspect(db.engine).get_table_names())
    return bool(tables) and "alembic_version" not in tables and "users" in tables


def sync_schema() -> str:
    """Upgrade to head. Returns a short description of what happened."""
    from flask_migrate import stamp, upgrade

    migrations_dir = current_app.extensions["migrate"].directory
    if _is_legacy_database():
        log.warning("legacy database detected (no alembic_version); stamping %s", BASELINE_REVISION)
        stamp(directory=migrations_dir, revision=BASELINE_REVISION)
        upgrade(directory=migrations_dir)
        return "stamped-baseline+upgraded"
    upgrade(directory=migrations_dir)
    return "upgraded"


def current_revision() -> str | None:
    if not _has_alembic_version():
        return None
    row = db.session.execute(text("SELECT version_num FROM alembic_version")).first()
    return row[0] if row else None


def seed_defaults(commit=True) -> dict:
    cfg = settings()
    created = {"projects": 0, "announcements": 0, "users": 0}

    if Project.query.count() == 0:
        defaults = [
            {
                "slug": "rover",
                "title": "Rover",
                "project_type": "Rover",
                "summary": "Mobility-focused platform with drivetrain, controls, and testing milestones.",
                "description": (
                    "The Rover team develops a rugged platform for terrain handling, control stability, "
                    "and subsystem validation through iterative build cycles."
                ),
                "status": "Active",
                "timeline": "Concept -> CAD -> Fabrication -> Integration -> Field Test",
                "gallery_json": json.dumps([]),
                "lead_name": "Rover Lead",
            },
            {
                "slug": "arm",
                "title": "Robotic Arm",
                "project_type": "Arm",
                "summary": "Manipulator design integrating structure, actuators, and controls workflows.",
                "description": (
                    "The Arm project focuses on payload handling, repeatability, and manufacturing-ready "
                    "component design for reliable operation."
                ),
                "status": "Prototype",
                "timeline": "Kinematics Study -> Linkage Design -> Controls Tuning -> Validation",
                "gallery_json": json.dumps([]),
                "lead_name": "Controls Lead",
            },
            {
                "slug": "manufacturing",
                "title": "Manufacturing",
                "project_type": "Manufacturing",
                "summary": "CAD-to-fabrication process ownership and quality-first part production.",
                "description": (
                    "The Manufacturing track supports all project teams with machining plans, print strategy, "
                    "and documentation for production consistency."
                ),
                "status": "In Progress",
                "timeline": "Manufacturing Plans -> Material Prep -> Production -> QA",
                "gallery_json": json.dumps([]),
                "lead_name": "Manufacturing Lead",
            },
        ]
        for payload in defaults:
            db.session.add(Project(**payload))
            created["projects"] += 1

    if Announcement.query.count() == 0:
        db.session.add(
            Announcement(
                title="Welcome to ASME @ UIowa",
                body="Spring build season is active. Check project boards and weekly meeting updates.",
                is_published=True,
                show_on_public=True,
                show_on_member=True,
                published_at=datetime.utcnow(),
            )
        )
        created["announcements"] += 1

    if User.query.count() == 0:
        from asme.auth.session import is_admin_member

        members = Member.query.order_by(Member.id.asc()).all()
        for member in members:
            role = "member"
            if "lead" in (member.member_class or "").lower():
                role = "team_leader"
            if is_admin_member(member):
                role = "admin"
            db.session.add(
                User(
                    name=member.name,
                    email=member.email,
                    username=(member.email.split("@", 1)[0] if member.email and "@" in member.email else None),
                    password_hash=generate_password_hash(cfg.default_user_password),
                    role=role,
                    is_active=True,
                    member_id=member.id,
                )
            )
            created["users"] += 1
        if not members:
            db.session.add(_bootstrap_admin(cfg))
            created["users"] += 1
    elif not User.query.filter(func.lower(User.email) == cfg.default_admin_email).first():
        db.session.add(_bootstrap_admin(cfg))
        created["users"] += 1

    seed_default_tracks(commit=False)
    db.session.flush()
    if cfg.ops_enabled:
        from asme.ops.seeds import seed_ops_defaults

        created["ops"] = seed_ops_defaults(commit=False)
    if commit:
        db.session.commit()
    return created


def _bootstrap_admin(cfg) -> User:
    email = cfg.default_admin_email
    return User(
        name="ASME Admin",
        email=email,
        username=(email.split("@", 1)[0] if "@" in email else "admin"),
        password_hash=generate_password_hash(cfg.default_admin_password),
        role="admin",
        is_active=True,
    )
