from asme.constants import ENT_SHOP_ACCESS
from asme.services.onboarding import entitlements


def test_health(client):
    body = client.get("/api/v1/health").get_json()
    assert body["ok"] and body["payload"]["api"] == "v1"


def test_me_includes_launchpad(client, users, login_as):
    login_as(users["member"])
    body = client.get("/api/v1/me").get_json()
    assert body["payload"]["user"]["email"] == "mo@uiowa.edu"
    assert body["payload"]["launchpad"]["current_phase"]["key"] == "signed_up"


def test_onboarding_payload_shape(client, users, login_as):
    login_as(users["member"])
    body = client.get("/api/v1/onboarding").get_json()
    track = body["payload"]["tracks"][0]
    assert track["key"] == "member_onboarding"
    assert [p["key"] for p in track["phases"]] == ["signed_up", "shop_ready", "contributor", "team_lead"]
    assert body["payload"]["percent"] == round(1 / 13 * 100)  # uiowa.edu email already satisfied


def test_checkout_requires_idempotency_key(client, users, item, login_as):
    login_as(users["member"])
    response = client.post("/api/v1/checkouts", json={"item_id": item.id})
    assert response.status_code == 400
    assert response.get_json()["code"] == "idempotency_required"


def test_checkout_and_return_roundtrip(client, users, item, login_as):
    login_as(users["member"])
    response = client.post("/api/v1/checkouts", json={"item_id": item.id, "qty": 1}, headers={"Idempotency-Key": "abc"})
    assert response.status_code == 201, response.data
    loan = response.get_json()["payload"]["checkout"]
    assert loan["state"] == "open"
    replay = client.post("/api/v1/checkouts", json={"item_id": item.id, "qty": 1}, headers={"Idempotency-Key": "abc"})
    assert replay.get_json()["payload"]["checkout"]["id"] == loan["id"]
    listing = client.get("/api/v1/checkouts?state=open").get_json()
    assert len(listing["payload"]["checkouts"]) == 1
    back = client.post(f"/api/v1/checkouts/{loan['id']}/return", json={"condition": "good"})
    assert back.status_code == 200
    assert back.get_json()["payload"]["checkout"]["state"] == "returned"


def test_error_shape_on_conflict(client, users, item, login_as):
    login_as(users["member"])
    client.post("/api/v1/checkouts", json={"item_id": item.id, "qty": 1}, headers={"Idempotency-Key": "a"})
    response = client.post("/api/v1/checkouts", json={"item_id": item.id, "qty": 1}, headers={"Idempotency-Key": "b"})
    assert response.status_code == 409
    body = response.get_json()
    assert body["ok"] is False and body["code"] == "already_checked_out" and "error" in body


def test_items_etag_and_cursor(client, users, item, login_as, db):
    from asme.models import Item

    for i in range(3):
        db.session.add(Item(name=f"Item {i}", total_qty=1, available_qty=1, active=True))
    db.session.commit()
    login_as(users["member"])
    first = client.get("/api/v1/items?limit=2")
    assert first.status_code == 200
    etag = first.headers["ETag"]
    body = first.get_json()
    assert len(body["payload"]["items"]) == 2 and body["payload"]["next_cursor"]
    assert client.get("/api/v1/items?limit=2", headers={"If-None-Match": etag}).status_code == 304
    second = client.get(f"/api/v1/items?limit=2&cursor={body['payload']['next_cursor']}").get_json()
    assert len(second["payload"]["items"]) == 2
    assert {i["id"] for i in body["payload"]["items"]}.isdisjoint({i["id"] for i in second["payload"]["items"]})


def test_bad_cursor_is_400(client, users, login_as):
    login_as(users["member"])
    assert client.get("/api/v1/items?cursor=!!!").status_code == 400


def test_checkin_creates_open_shop_event_when_nothing_scheduled(client, users, login_as):
    login_as(users["member"])
    response = client.post("/api/v1/checkins", json={})
    assert response.status_code == 201, response.data
    payload = response.get_json()["payload"]
    assert payload["event"]["kind"] == "open_shop"
    again = client.post("/api/v1/checkins", json={})
    assert again.status_code == 200 and again.get_json()["payload"]["created"] is False


def test_complete_manual_task_via_api_rejects_rule_driven(client, users, login_as):
    login_as(users["member"])
    response = client.post("/api/v1/onboarding/tasks/complete_profile/complete", json={})
    assert response.status_code == 400
    assert response.get_json()["code"] == "task_not_manual"


def test_lead_can_view_other_member_launchpad(client, users, login_as):
    login_as(users["lead"])
    response = client.get(f"/api/v1/onboarding?user_id={users['member'].id}")
    assert response.status_code == 200
    assert response.get_json()["payload"]["subject"]["id"] == str(users["member"].id)


def test_member_cannot_view_other_member_launchpad(client, users, login_as):
    login_as(users["member"])
    assert client.get(f"/api/v1/onboarding?user_id={users['lead'].id}").status_code == 403


def test_print_request_patch_needs_lead(client, users, login_as, db):
    from asme.models import PrintRequest

    row = PrintRequest(user_id=users["member"].id, printer_type="H2S", file_link="http://x", filament="PLA", status="submitted")
    db.session.add(row)
    db.session.commit()
    login_as(users["member"])
    assert client.patch(f"/api/v1/print-requests/{row.id}", json={"status": "approved"}).status_code == 403
    client.post("/logout")
    login_as(users["admin"])
    response = client.patch(f"/api/v1/print-requests/{row.id}", json={"status": "printing"})
    assert response.status_code == 200
    assert response.get_json()["payload"]["print_request"]["runs"][0]["status"] == "printing"
