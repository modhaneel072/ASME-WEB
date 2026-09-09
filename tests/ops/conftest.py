"""Fixtures for the ASME Ops domain: seeded organisation, one user per system role,
authorisation contexts and an API client that sends the CSRF header."""

from __future__ import annotations

import pytest
from werkzeug.security import generate_password_hash

from asme.extensions import db as _db
from asme.models import User
from asme.ops import authz
from asme.ops.models import Organization
from asme.ops.seeds import seed_ops_defaults
from asme.ops.tenancy import ensure_membership, role_by_key
from tests.conftest import PASSWORD, make_app

ROLE_KEYS = [
    "chapter_admin", "executive_officer", "project_lead", "team_lead", "full_member", "shop_operator", "requester",
    "inventory_manager", "safety_officer", "treasurer", "faculty_advisor", "sponsor_guest",
]


@pytest.fixture
def app():
    application = make_app(ops_enabled=True, file_url_ttl_seconds=60)
    with application.app_context():
        _db.create_all()
        from asme.services.onboarding.seeds import seed_default_tracks

        seed_default_tracks(commit=False)
        seed_ops_defaults(commit=True)
        yield application
        _db.session.remove()
        _db.drop_all()


@pytest.fixture
def db(app):
    return _db


@pytest.fixture
def org(app) -> Organization:
    return Organization.query.filter_by(slug="asme-uiowa").first()


def make_user(org, name, email, role_key, legacy_role="member"):
    user = User(name=name, email=email, username=email.split("@")[0], password_hash=generate_password_hash(PASSWORD), role=legacy_role, is_active=True, major="ME", graduation_year=2027)
    _db.session.add(user)
    _db.session.flush()
    membership = ensure_membership(user, org)
    membership.role_id = role_by_key(org, role_key).id
    _db.session.commit()
    return user


@pytest.fixture
def people(org):
    """One user per system role, keyed by role key."""
    out = {}
    for key in ROLE_KEYS:
        out[key] = make_user(org, key.replace("_", " ").title(), f"{key}@uiowa.edu", key, legacy_role={"chapter_admin": "admin", "team_lead": "team_leader"}.get(key, "member"))
    return out


@pytest.fixture
def ctx_for(org):
    def _ctx(user):
        membership = ensure_membership(user, org)
        _db.session.commit()
        return authz.build_context(user, org, membership)

    return _ctx


@pytest.fixture
def other_org(app):
    """A second tenant with its own admin, for ID-tampering tests."""
    org2 = Organization(name="Other Chapter", slug="other-chapter")
    _db.session.add(org2)
    _db.session.flush()
    authz.ensure_system_roles(org2)
    from asme.ops.services.directory import default_location

    default_location(org2.id)
    user = User(name="Other Admin", email="other-admin@example.edu", username="otheradmin", password_hash=generate_password_hash(PASSWORD), role="member", is_active=True)
    _db.session.add(user)
    _db.session.flush()
    membership = ensure_membership(user, org2)
    membership.role_id = role_by_key(org2, "chapter_admin").id
    _db.session.commit()
    return {"org": org2, "admin": user, "ctx": authz.build_context(user, org2, membership)}


class ApiClient:
    """Test client wrapper that logs in through the ops session endpoint and sends the CSRF header."""

    def __init__(self, app, user=None):
        self.client = app.test_client()
        self.headers = {"X-Requested-With": "ASME-Ops"}
        if user is not None:
            response = self.client.post("/api/v1/ops/session/login", json={"identifier": user.email, "password": PASSWORD}, headers=self.headers)
            assert response.status_code == 200, response.data
            self.session = response.get_json()["payload"]

    def get(self, path, **kw):
        return self.client.get(path, headers=self.headers, **kw)

    def post(self, path, json=None, **kw):
        return self.client.post(path, json=json, headers=self.headers, **kw)

    def patch(self, path, json=None, **kw):
        return self.client.patch(path, json=json, headers=self.headers, **kw)

    def put(self, path, json=None, **kw):
        return self.client.put(path, json=json, headers=self.headers, **kw)

    def delete(self, path, **kw):
        return self.client.delete(path, headers=self.headers, **kw)


@pytest.fixture
def api(app):
    def _api(user=None):
        return ApiClient(app, user)

    return _api


@pytest.fixture
def lookups(app, people, ctx_for):
    """A project led by the project lead, a team led by the team lead, a location, categories and an asset."""
    from asme.ops.schemas import AssetCreate, LocationCreate, ProjectCreate, TeamCreate, TeamMemberInput
    from asme.ops.services import assets as assets_service
    from asme.ops.services import directory, projects

    admin = ctx_for(people["chapter_admin"])
    project = projects.create_project(admin, ProjectCreate(name="Crater Cruncher Rover", code="CCR", lead_user_id=people["project_lead"].id))
    other_project = projects.create_project(admin, ProjectCreate(name="Showcase", code="SHOW", lead_user_id=people["executive_officer"].id))
    team = directory.create_team(admin, TeamCreate(name="Wheels", project_id=project.id, members=[TeamMemberInput(user_id=people["team_lead"].id, is_lead=True), TeamMemberInput(user_id=people["full_member"].id), TeamMemberInput(user_id=people["shop_operator"].id)]))
    other_team = directory.create_team(admin, TeamCreate(name="Outreach", project_id=other_project.id, members=[TeamMemberInput(user_id=people["executive_officer"].id, is_lead=True)]))
    location = directory.create_location(admin, LocationCreate(name="Robotics Lab"))
    categories = directory.list_categories(admin)
    asset = assets_service.create_asset(admin, AssetCreate(name="Rover", code="ROVER", project_id=project.id, location_id=location.id))
    return {"project": project, "other_project": other_project, "team": team, "other_team": other_team, "location": location, "categories": categories, "asset": asset}
