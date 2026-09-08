from asme.constants import ENT_SHOP_ACCESS
from asme.services.onboarding import entitlements
from tests.conftest import PASSWORD, login


def test_login_page_renders(client):
    assert client.get("/login").status_code == 200


def test_login_success_and_session_guard(client, users):
    response = login(client, users["member"])
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/portal")
    routed = client.get("/portal")
    assert routed.status_code == 302 and routed.headers["Location"].endswith("/portal/member")
    page = client.get("/portal/member")
    assert page.status_code == 200
    assert "no-store" in page.headers.get("Cache-Control", "")


def test_login_rejects_bad_password(client, users):
    response = client.post("/login", data={"identifier": users["member"].email, "password": "nope"})
    assert response.status_code == 200
    assert b"Invalid email or password" in response.data


def test_login_rate_limit(app, users):
    app.extensions["asme_login_rate_limiter"].max_attempts = 2
    client = app.test_client()
    for _ in range(2):
        client.post("/login", data={"identifier": users["member"].email, "password": "nope"})
    response = client.post("/login", data={"identifier": users["member"].email, "password": PASSWORD})
    assert b"Too many login attempts" in response.data


def test_member_cannot_open_admin(client, users, login_as):
    login_as(users["member"])
    response = client.get("/portal/admin")
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/portal")


def test_admin_router_goes_to_admin(client, users, login_as):
    login_as(users["admin"])
    response = client.get("/portal")
    assert response.headers["Location"].endswith("/portal/admin")


def test_json_routes_get_json_denials(client):
    response = client.get("/api/v1/me")
    assert response.status_code == 401
    assert response.get_json()["code"] == "login_required"


def test_shadow_mode_lets_checkout_through_and_records_block(client, users, item, login_as):
    login_as(users["member"])
    assert not entitlements.has_entitlement(users["member"], ENT_SHOP_ACCESS)
    response = client.post("/portal/inventory/checkout", data={"item_id": item.id, "qty": 1}, follow_redirects=False)
    assert response.status_code == 302
    assert response.headers["Location"].rstrip("/").endswith(("/portal/member/checkouts", "/portal/member/my-items"))


def test_enforced_mode_blocks_checkout_with_phase_hint(enforced_app, monkeypatch):
    from tests.conftest import _db, _user

    member = _user("Mo Member", "mo@uiowa.edu", "member")
    from asme.models import Item

    row = Item(name="Mill", total_qty=1, available_qty=1, active=True)
    _db.session.add(row)
    _db.session.commit()
    client = enforced_app.test_client()
    login(client, member)
    response = client.post(
        "/api/v1/checkouts",
        json={"item_id": row.id, "qty": 1},
        headers={"Idempotency-Key": "k1"},
    )
    assert response.status_code == 403
    body = response.get_json()
    assert body["code"] == "entitlement_required"
    assert body["entitlement"] == ENT_SHOP_ACCESS
    assert body["phase"] == "shop_ready"


def test_admin_bypasses_entitlements(enforced_app):
    from tests.conftest import _db, _user

    admin = _user("Ada", "ada@uiowa.edu", "admin")
    _db.session.commit()
    assert entitlements.has_entitlement(admin, ENT_SHOP_ACCESS)


def test_logout_clears_session(client, users, login_as):
    login_as(users["member"])
    client.post("/logout")
    assert client.get("/portal/member").status_code == 302
