"""Every page a person can open renders without a server error."""

import pytest

PUBLIC = ["/", "/events", "/gallery", "/who-we-are", "/executive-team", "/projects", "/contact", "/join", "/sponsors", "/login", "/signup", "/forgot-password", "/admin-login", "/kiosk", "/healthz"]
MEMBER = [
    "/portal/member",
    "/portal/member/inventory",
    "/portal/member/inventory?q=cal",
    "/portal/member/my-items",
    "/portal/member/prints",
    "/portal/member/calendar",
    "/portal/member/profile",
    "/portal/member/help",
    "/portal/member/launchpad",
]
LEAD = ["/portal/leader/launchpad", "/portal/member/schedule"]
ADMIN = [
    "/portal/admin",
    "/portal/admin/members",
    "/portal/admin/members?show_inactive=1",
    "/portal/admin/attendance",
    "/portal/admin/attendance?tab=history",
    "/portal/admin/inventory",
    "/portal/admin/prints",
    "/portal/admin/calendar",
    "/portal/admin/settings",
    "/portal/admin/launchpad",
    "/portal/admin/attendance/export.csv",
    "/portal/admin/inventory/treasury-report.csv",
]
LEGACY_REDIRECTS = ["/dashboard", "/attendance", "/inventory", "/prints", "/activity", "/settings", "/scan", "/my-items", "/admin/nfc", "/calendar", "/app"]


@pytest.mark.parametrize("path", PUBLIC)
def test_public_pages(client, users, path):
    response = client.get(path)
    assert response.status_code == 200, (path, response.status_code)


@pytest.mark.parametrize("path", MEMBER)
def test_member_pages(client, users, item, project, login_as, path):
    login_as(users["member"])
    response = client.get(path)
    assert response.status_code == 200, (path, response.status_code, response.data[:300])


@pytest.mark.parametrize("path", LEAD)
def test_lead_pages(client, users, item, login_as, path):
    login_as(users["lead"])
    response = client.get(path)
    assert response.status_code == 200, (path, response.status_code, response.data[:300])


@pytest.mark.parametrize("path", ADMIN)
def test_admin_pages(client, users, item, project, login_as, path):
    login_as(users["admin"])
    response = client.get(path)
    assert response.status_code == 200, (path, response.status_code, response.data[:300])


@pytest.mark.parametrize("path", LEGACY_REDIRECTS)
def test_legacy_ops_redirects_when_disabled(client, users, path):
    response = client.get(path)
    assert response.status_code == 302, path


@pytest.mark.parametrize("path", MEMBER + ADMIN)
def test_portal_requires_login(client, path):
    response = client.get(path)
    assert response.status_code == 302 and "/login" in response.headers["Location"]


def test_admin_can_switch_to_member_view(client, users, item, project, login_as):
    login_as(users["admin"])
    assert client.get("/portal/member/launchpad").status_code == 200


def test_route_count_is_complete(app):
    rules = {rule.rule for rule in app.url_map.iter_rules()}
    for expected in ["/api/v1/checkouts", "/api/checkout", "/transact", "/portal/inventory/checkout", "/portal/admin/launchpad", "/checkin", "/pair/item", "/export"]:
        assert expected in rules, expected
