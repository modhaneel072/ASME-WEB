"""Permission matrix spot checks, scope resolution, and ID tampering across tenants."""

import pytest

from asme.ops import authz
from asme.ops.schemas import ProjectCreate, WorkOrderCreate, WorkOrderUpdate
from asme.ops.services import projects, work_orders
from asme.services.errors import NotFound


def test_system_roles_cover_every_permission_key(org):
    roles = authz.ensure_system_roles(org)
    admin_keys = {rp.permission.key for rp in roles["chapter_admin"].permissions}
    assert admin_keys == set(authz.PERMISSIONS)
    for key, (name, rank, grants) in authz.SYSTEM_ROLES.items():
        for perm, scope in grants:
            assert perm in authz.PERMISSIONS, (key, perm)
            assert scope in authz.SCOPES


def test_legacy_role_mapping_roundtrip():
    assert authz.LEGACY_ROLE_MAP == {"admin": "chapter_admin", "team_leader": "team_lead", "member": "full_member"}
    assert authz.legacy_role_for("chapter_admin") == "admin"
    assert authz.legacy_role_for("team_lead") == "team_leader"
    assert authz.legacy_role_for("full_member") == "member"
    assert authz.legacy_role_for("treasurer") == "member"


def test_membership_created_from_legacy_role(org, db):
    from tests.ops.conftest import make_user

    user = make_user(org, "Legacy Admin", "legacy-admin@uiowa.edu", "full_member", legacy_role="admin")
    from asme.ops.models import Membership

    Membership.query.filter_by(user_id=user.id).delete()
    db.session.commit()
    from asme.ops.tenancy import ensure_membership

    membership = ensure_membership(user, org)
    db.session.commit()
    assert membership.role.system_key == "chapter_admin"


@pytest.mark.parametrize(
    "role,permission,expected",
    [
        ("chapter_admin", "org.manage", True),
        ("executive_officer", "org.manage", False),
        ("executive_officer", "user.manage", True),
        ("requester", "work_order.create", False),
        ("requester", "request.submit", True),
        ("shop_operator", "work_order.read_all", False),
        ("shop_operator", "work_order.read_assigned", True),
        ("treasurer", "purchase.review", True),
        ("treasurer", "work_order.create", False),
        ("faculty_advisor", "audit.read", True),
        ("sponsor_guest", "project.read", False),
        ("safety_officer", "work_order.set_critical", True),
        ("full_member", "work_order.set_critical", False),
        ("inventory_manager", "inventory.manage", True),
        ("executive_officer", "inventory.manage", False),
    ],
)
def test_chapter_level_permissions(people, ctx_for, role, permission, expected):
    assert authz.can(ctx_for(people[role]), permission) is expected


def test_team_lead_scope_is_their_team_only(people, ctx_for, lookups):
    lead = ctx_for(people["team_lead"])
    mine = work_orders.create(ctx_for(people["chapter_admin"]), WorkOrderCreate(title="Ours", team_id=lookups["team"].id, project_id=lookups["project"].id))
    theirs = work_orders.create(ctx_for(people["chapter_admin"]), WorkOrderCreate(title="Theirs", team_id=lookups["other_team"].id, project_id=lookups["other_project"].id))
    assert authz.can(lead, "work_order.edit", record=mine)
    assert not authz.can(lead, "work_order.edit", record=theirs)
    assert authz.can_create_for(lead, "work_order.create", team_id=lookups["team"].id)
    assert not authz.can_create_for(lead, "work_order.create", team_id=lookups["other_team"].id)
    with pytest.raises(authz.Forbidden):
        work_orders.update(lead, theirs, WorkOrderUpdate(title="Hijacked"), {"title"})


def test_project_lead_scope_is_their_project(people, ctx_for, lookups):
    pl = ctx_for(people["project_lead"])
    admin = ctx_for(people["chapter_admin"])
    in_project = work_orders.create(admin, WorkOrderCreate(title="Rover task", project_id=lookups["project"].id))
    outside = work_orders.create(admin, WorkOrderCreate(title="Showcase task", project_id=lookups["other_project"].id))
    assert authz.can(pl, "work_order.assign", record=in_project)
    assert not authz.can(pl, "work_order.assign", record=outside)
    assert authz.can(pl, "project.manage", record=lookups["project"])
    assert not authz.can(pl, "project.manage", record=lookups["other_project"])


def test_member_can_only_edit_own_and_start_assigned(people, ctx_for, lookups):
    admin = ctx_for(people["chapter_admin"])
    member = ctx_for(people["full_member"])
    own = work_orders.create(member, WorkOrderCreate(title="My own task"))
    assigned = work_orders.create(admin, WorkOrderCreate(title="Assigned to me", assignee_user_ids=[people["full_member"].id]))
    unrelated = work_orders.create(admin, WorkOrderCreate(title="Someone else's"))
    assert authz.can(member, "work_order.edit", record=own)
    assert not authz.can(member, "work_order.edit", record=unrelated)
    assert authz.can(member, "work_order.start", record=assigned)
    assert not authz.can(member, "work_order.start", record=unrelated)
    # members may create work but not assign others to it
    from asme.services.errors import Validation

    with pytest.raises(Validation):
        work_orders.create(member, WorkOrderCreate(title="Delegating", assignee_user_ids=[people["team_lead"].id]))


def test_shop_operator_sees_only_assigned_work(people, ctx_for, lookups):
    admin = ctx_for(people["chapter_admin"])
    operator = ctx_for(people["shop_operator"])
    visible = work_orders.create(admin, WorkOrderCreate(title="Print bracket", assignee_user_ids=[people["shop_operator"].id]))
    via_team = work_orders.create(admin, WorkOrderCreate(title="Team job", assignee_team_ids=[lookups["team"].id]))
    hidden = work_orders.create(admin, WorkOrderCreate(title="Exec only"))
    ids = {w.id for w in work_orders.visible_query(operator).all()}
    assert visible.id in ids and via_team.id in ids and hidden.id not in ids
    with pytest.raises(NotFound):
        work_orders.get_visible_or_404(operator, hidden.id)


def test_cross_tenant_ids_are_invisible(people, ctx_for, lookups, other_org):
    """A record id from another organisation looks exactly like a missing record."""
    admin = ctx_for(people["chapter_admin"])
    foreign = projects.create_project(other_org["ctx"], ProjectCreate(name="Foreign", code="FOR"))
    from asme.ops.models import OpsProject
    from asme.ops.tenancy import get_or_404

    with pytest.raises(NotFound):
        get_or_404(OpsProject, foreign.id, admin)
    foreign_wo = work_orders.create(other_org["ctx"], WorkOrderCreate(title="Foreign work"))
    with pytest.raises(NotFound):
        work_orders.get_visible_or_404(admin, foreign_wo.id)
    # even with a chapter-wide grant, a record belonging to another org is refused
    assert not authz.can(admin, "work_order.edit", record=foreign_wo)


def test_api_rejects_cross_tenant_patch(people, api, other_org, lookups):
    foreign = work_orders.create(other_org["ctx"], WorkOrderCreate(title="Foreign"))
    client = api(people["chapter_admin"])
    response = client.patch(f"/api/v1/work-orders/{foreign.id}", json={"title": "Stolen"})
    assert response.status_code == 404
    assert response.get_json()["code"] == "not_found"


def test_api_forbidden_payload_names_the_permission(people, api):
    client = api(people["requester"])
    response = client.post("/api/v1/projects", json={"name": "Nope"})
    assert response.status_code == 403
    body = response.get_json()
    assert body["code"] == "forbidden" and body["permission"] == "project.create"


def test_last_admin_cannot_demote_self(people, api, org):
    client = api(people["chapter_admin"])
    response = client.patch(f"/api/v1/users/{people['chapter_admin'].id}", json={"role_key": "full_member"})
    assert response.status_code == 409
    assert response.get_json()["code"] == "last_admin"


def test_role_change_syncs_legacy_role(people, api, db):
    client = api(people["chapter_admin"])
    response = client.patch(f"/api/v1/users/{people['requester'].id}", json={"role_key": "team_lead"})
    assert response.status_code == 200, response.data
    db.session.refresh(people["requester"])
    assert people["requester"].role == "team_leader"
    assert response.get_json()["payload"]["role_key"] == "team_lead"
