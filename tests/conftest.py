"""Shared fixtures: an in-memory app, seeded tracks, three users, and a login helper."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from werkzeug.security import generate_password_hash

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("ASME_ENV", "testing")

from asme import create_app  # noqa: E402
from asme import events as event_bus  # noqa: E402
from asme.extensions import db as _db  # noqa: E402
from asme.models import Item, Member, Project, User  # noqa: E402
from asme.services.onboarding.seeds import seed_default_tracks  # noqa: E402

PASSWORD = "correct-horse-battery"


def make_app(**overrides):
    base = dict(
        env="testing",
        secret_key="test-secret",
        database_url="sqlite://",
        auto_migrate=False,
        outbox_worker_enabled=False,
        onboarding_enforce=False,
        session_cookie_secure=False,
        calendar_provider="google",
        login_rate_max_attempts=100,
    )
    base.update(overrides)
    return create_app(**base)


@pytest.fixture
def app():
    application = make_app()
    with application.app_context():
        _db.create_all()
        seed_default_tracks()
        yield application
        _db.session.remove()
        _db.drop_all()


@pytest.fixture
def enforced_app():
    application = make_app(onboarding_enforce=True)
    with application.app_context():
        _db.create_all()
        seed_default_tracks()
        yield application
        _db.session.remove()
        _db.drop_all()


@pytest.fixture
def db(app):
    return _db


def _user(name, email, role, **extra):
    user = User(
        name=name,
        email=email,
        username=email.split("@")[0],
        password_hash=generate_password_hash(PASSWORD),
        role=role,
        is_active=True,
        **extra,
    )
    _db.session.add(user)
    _db.session.flush()
    return user


@pytest.fixture
def users(app):
    admin = _user("Ada Admin", "ada@uiowa.edu", "admin")
    lead = _user("Lee Lead", "lee@uiowa.edu", "team_leader", major="ME", graduation_year=2027)
    member = _user("Mo Member", "mo@uiowa.edu", "member")
    _db.session.commit()
    return {"admin": admin, "lead": lead, "member": member}


@pytest.fixture
def item(app):
    row = Item(name="Digital Calipers", category="Tools", location="Drawer 2", total_qty=3, available_qty=3, active=True)
    _db.session.add(row)
    _db.session.commit()
    return row


@pytest.fixture
def project(app):
    row = Project(slug="rover", title="Rover", summary="s", description="d", status="Active", is_joinable=True)
    _db.session.add(row)
    _db.session.commit()
    return row


@pytest.fixture
def client(app):
    return app.test_client()


def login(client, user, password=PASSWORD):
    return client.post("/login", data={"identifier": user.email, "password": password}, follow_redirects=False)


@pytest.fixture
def login_as(client):
    def _login(user):
        response = login(client, user)
        assert response.status_code in (302, 303), response.data
        return client

    return _login


@pytest.fixture
def captured_events():
    seen = []

    def _catch(name, **payload):
        seen.append((name, payload))

    event_bus.subscribe("*", _catch)
    yield seen
    event_bus.unsubscribe("*", _catch)
