"""HTTP-level behaviour: login, session shape, error envelopes, CSRF header, files, OpenAPI, dashboard."""

import io

import pytest

from asme.ops.schemas import WorkOrderCreate
from asme.ops.services import work_orders as wo_service
from tests.conftest import PASSWORD


def test_login_by_email_and_username(app, people):
    client = app.test_client()
    headers = {"X-Requested-With": "ASME-Ops"}
    r = client.post("/api/v1/ops/session/login", json={"identifier": "full_member@uiowa.edu", "password": PASSWORD}, headers=headers)
    assert r.status_code == 200
    payload = r.get_json()["payload"]
    assert payload["role"]["key"] == "full_member" and "work_order.create" in payload["permissions"]
    client.post("/api/v1/ops/session/logout", headers=headers)
    r = client.post("/api/v1/ops/session/login", json={"identifier": "full_member", "password": PASSWORD}, headers=headers)
    assert r.status_code == 200
    r = client.post("/api/v1/ops/session/login", json={"identifier": "full_member", "password": "wrong"}, headers=headers)
    assert r.status_code == 401 and r.get_json()["code"] == "invalid_credentials"


def test_unauthenticated_and_csrf(app, people):
    client = app.test_client()
    assert client.get("/api/v1/work-orders").status_code == 401
    client.post("/api/v1/ops/session/login", json={"identifier": "chapter_admin@uiowa.edu", "password": PASSWORD}, headers={"X-Requested-With": "ASME-Ops"})
    r = client.post("/api/v1/work-orders", json={"title": "No header"})
    assert r.status_code == 403 and r.get_json()["code"] == "csrf"


def test_validation_error_shape(people, api):
    client = api(people["chapter_admin"])
    r = client.post("/api/v1/work-orders", json={"title": "", "priority": "URGENT", "bogus": 1})
    assert r.status_code == 400
    body = r.get_json()
    assert body["code"] == "validation"
    fields = {d["field"] for d in body["details"]}
    assert "title" in fields and "priority" in fields and "bogus" in fields


def test_work_order_roundtrip_over_http(people, api, lookups):
    lead = api(people["team_lead"])
    r = lead.post("/api/v1/work-orders", json={"title": "Inspect wheel hub fasteners", "team_id": lookups["team"].id, "project_id": lookups["project"].id, "priority": "HIGH", "assignee_user_ids": [people["full_member"].id], "category_ids": [lookups["categories"][0].id]})
    assert r.status_code == 201, r.data
    wo = r.get_json()["payload"]
    assert wo["number"] == 1 and wo["permissions"]["edit"] is True and wo["allowed_transitions"] == ["CANCELED", "DONE", "IN_PROGRESS", "ON_HOLD"]

    member = api(people["full_member"])
    listing = member.get(f"/api/v1/work-orders?filter[assignee]=me&filter[priority]=HIGH&sort=due_asc")
    assert listing.status_code == 200
    items = listing.get_json()["payload"]["items"]
    assert [i["id"] for i in items] == [wo["id"]]
    assert listing.get_json()["payload"]["counts"] == {"todo": 1, "done": 0}

    r = member.post(f"/api/v1/work-orders/{wo['id']}/start", json={})
    assert r.status_code == 200 and r.get_json()["payload"]["status"] == "IN_PROGRESS"
    r = member.post(f"/api/v1/work-orders/{wo['id']}/comments", json={"body": "Torque wrench found."})
    assert r.status_code == 201
    r = member.post(f"/api/v1/work-orders/{wo['id']}/time-entries", json={"minutes": 40})
    assert r.status_code == 201 and r.get_json()["payload"]["actual_minutes"] == 40
    r = member.post(f"/api/v1/work-orders/{wo['id']}/complete", json={"completion_note": "All good", "time_minutes": 5})
    assert r.status_code == 200 and r.get_json()["payload"]["status"] == "DONE" and r.get_json()["payload"]["actual_minutes"] == 45
    # member may not edit someone else's closed work; lead gets a clear conflict
    r = lead.patch(f"/api/v1/work-orders/{wo['id']}", json={"title": "x"})
    assert r.status_code == 409 and r.get_json()["code"] == "closed"
    history = lead.get(f"/api/v1/work-orders/{wo['id']}/history").get_json()["payload"]
    assert [h["to_status"] for h in history["status_history"]] == ["OPEN", "IN_PROGRESS", "DONE"]
    assert any(e["event_type"] == "work_order.time_logged" for e in history["events"])


def test_unknown_filter_and_sort_rejected(people, api):
    client = api(people["chapter_admin"])
    assert client.get("/api/v1/work-orders?filter[evil]=1").status_code == 400
    assert client.get("/api/v1/work-orders?sort=drop_table").status_code == 400
    assert client.get("/api/v1/work-orders?page[cursor]=!!!").status_code == 400


def test_requester_cannot_read_work_orders(people, api):
    client = api(people["requester"])
    r = client.get("/api/v1/work-orders")
    assert r.status_code == 403 and r.get_json()["permission"] == "work_order.read_assigned"


def test_file_upload_validation_and_signed_download(app, people, api, ctx_for):
    admin = api(people["chapter_admin"])
    wo = wo_service.create(ctx_for(people["chapter_admin"]), WorkOrderCreate(title="With files"))
    bad = admin.client.post("/api/v1/files", data={"entity_type": "work_order", "entity_id": wo.id, "file": (io.BytesIO(b"MZ..."), "virus.exe")}, headers={"X-Requested-With": "ASME-Ops"}, content_type="multipart/form-data")
    assert bad.status_code == 400 and bad.get_json()["code"] == "file_type"
    good = admin.client.post("/api/v1/files", data={"entity_type": "work_order", "entity_id": wo.id, "file": (io.BytesIO(b"hello"), "notes.txt")}, headers={"X-Requested-With": "ASME-Ops"}, content_type="multipart/form-data")
    assert good.status_code == 201, good.data
    meta = good.get_json()["payload"]
    assert meta["size_bytes"] == 5 and meta["kind"] == "file" and meta["download_url"].startswith(f"/api/v1/files/{meta['id']}/download?token=")
    # signed link works without a session
    anon = app.test_client()
    assert anon.get(meta["download_url"]).status_code == 200
    assert anon.get(meta["download_url"]).data == b"hello"
    assert anon.get(f"/api/v1/files/{meta['id']}/download?token=forged").status_code == 403
    # expired token
    from asme.ops import storage

    with app.app_context():
        token = storage.sign_download(meta["id"])
        assert storage.verify_download(token, max_age=-1) is None
    listing = admin.get(f"/api/v1/files?entity_type=work_order&entity_id={wo.id}").get_json()["payload"]
    assert [f["id"] for f in listing] == [meta["id"]]
    assert admin.delete(f"/api/v1/files/{meta['id']}").status_code == 200
    assert admin.get(f"/api/v1/files?entity_type=work_order&entity_id={wo.id}").get_json()["payload"] == []


def test_operator_cannot_upload_to_unassigned_work(people, api, ctx_for):
    wo = wo_service.create(ctx_for(people["chapter_admin"]), WorkOrderCreate(title="Not yours"))
    operator = api(people["shop_operator"])
    r = operator.client.post("/api/v1/files", data={"entity_type": "work_order", "entity_id": wo.id, "file": (io.BytesIO(b"x"), "a.txt")}, headers={"X-Requested-With": "ASME-Ops"}, content_type="multipart/form-data")
    assert r.status_code == 404  # invisible, not merely forbidden


def test_openapi_document_lists_contracts(people, api):
    client = api(people["chapter_admin"])
    doc = client.get("/api/v1/ops/openapi.json").get_json()
    assert doc["openapi"].startswith("3.1")
    assert "/work-orders" in doc["paths"] and "post" in doc["paths"]["/work-orders"]
    assert doc["paths"]["/work-orders"]["post"]["requestBody"]["content"]["application/json"]["schema"]["$ref"].endswith("WorkOrderCreate")
    assert "WorkOrderCreate" in doc["components"]["schemas"] and "Error" in doc["components"]["schemas"]


def test_dashboard_metrics(people, api, ctx_for, lookups):
    admin_ctx = ctx_for(people["chapter_admin"])
    from datetime import datetime, timedelta

    from asme.ops.schemas import WorkOrderComplete

    wo_service.create(admin_ctx, WorkOrderCreate(title="Open A", priority="HIGH", team_id=lookups["team"].id, due_at=datetime.utcnow() - timedelta(days=1)))
    wo_service.create(admin_ctx, WorkOrderCreate(title="Open B", priority="LOW", team_id=lookups["team"].id))
    done = wo_service.create(admin_ctx, WorkOrderCreate(title="Done C", due_at=datetime.utcnow() + timedelta(days=1)))
    wo_service.start(admin_ctx, done)
    wo_service.complete(admin_ctx, done, WorkOrderComplete(time_minutes=60, costs=[{"type": "part", "amount": "10"}]))
    payload = api(people["chapter_admin"]).get("/api/v1/ops/dashboard/operations?range=30d").get_json()["payload"]
    assert payload["totals"]["open"] == 2 and payload["totals"]["overdue"] == 1 and payload["totals"]["completed"] == 1 and payload["totals"]["created"] == 3
    assert payload["by_priority"]["HIGH"] == 1 and payload["on_time_completion_rate"] == 100
    assert payload["workload_by_team"][0]["open"] == 2 and payload["hours_logged"] == 1.0 and payload["costs"]["parts"] == 10.0
    assert api(people["requester"]).get("/api/v1/ops/dashboard/operations").status_code == 403


def test_session_and_setup_banner(people, api):
    client = api(people["full_member"])
    session = client.get("/api/v1/ops/session").get_json()["payload"]
    assert session["setup_banner_dismissed"] is False and session["organization"]["slug"] == "asme-uiowa"
    assert client.post("/api/v1/ops/setup/banner", json={"dismissed": True}).status_code == 200
    assert client.get("/api/v1/ops/session").get_json()["payload"]["setup_banner_dismissed"] is True
    setup = client.get("/api/v1/ops/setup").get_json()["payload"]
    assert setup["phases"][0]["key"] == "foundation" and setup["required_total"] == 12


def test_lookups_and_saved_filters(people, api, lookups):
    client = api(people["team_lead"])
    lk = client.get("/api/v1/ops/lookups").get_json()["payload"]
    assert {u["id"] for u in lk["users"]} >= {people["team_lead"].id, people["full_member"].id}
    assert any(t["name"] == "Wheels" for t in lk["teams"]) and len(lk["categories"]) == 14
    r = client.post("/api/v1/saved-filters", json={"entity_type": "work_order", "name": "My overdue", "filters": {"due": ["overdue"], "assignee": ["me"]}, "sort": "priority_desc", "visibility": "team", "team_id": lookups["team"].id})
    assert r.status_code == 201, r.data
    member = api(people["full_member"])
    shared = member.get("/api/v1/saved-filters?entity_type=work_order").get_json()["payload"]["shared"]
    assert [s["name"] for s in shared] == ["My overdue"]
    r = member.post("/api/v1/saved-filters", json={"entity_type": "work_order", "name": "Chapter wide", "visibility": "chapter"})
    assert r.status_code == 403  # members cannot share
