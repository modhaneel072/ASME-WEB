from datetime import datetime, timedelta

import pytest

from asme.constants import ENT_EVENT_CHECKIN, ENT_PORTAL_ACCESS, ENT_PRINT_SUBMIT, ENT_SHOP_ACCESS
from asme.models import Event, Item, NFCTag, Phase, PhaseState, Track, TrainingCompletion, TrainingModule, User
from asme.services import attendance, content, inventory
from asme.services.errors import Forbidden, Validation
from asme.services.onboarding import engine, entitlements


def _phase(progress, key):
    return next(p for p in progress["tracks"][0]["phases"] if p["key"] == key)


def _task(phase, key):
    return next(t for t in phase["tasks"] if t["key"] == key)


def test_seed_is_idempotent(app):
    from asme.services.onboarding.seeds import seed_default_tracks

    seed_default_tracks()
    seed_default_tracks()
    assert Track.query.count() == 2
    assert Phase.query.count() == 6
    assert TrainingModule.query.count() == 3


def test_new_member_starts_in_phase_one(app, users):
    progress = engine.evaluate_user(users["member"])
    signed_up = _phase(progress, "signed_up")
    assert signed_up["status"] == "active"
    assert _phase(progress, "shop_ready")["status"] == "locked"
    # a @uiowa.edu address satisfies the email task on day one: 1 of 13 required
    assert _task(signed_up, "verify_email")["status"] == "complete"
    assert progress["percent"] == round(1 / 13 * 100)
    assert progress["steps_left"] == 3


def test_phase_one_completes_and_grants(app, users, project, db):
    member = users["member"]
    # email already @uiowa.edu; fill the rest
    member.major = "ME"
    member.graduation_year = 2028
    db.session.commit()
    progress = engine.evaluate_user(member)
    assert _task(_phase(progress, "signed_up"), "complete_profile")["status"] == "complete"
    assert _task(_phase(progress, "signed_up"), "link_nfc")["status"] == "pending"

    db.session.add(NFCTag(tag_uid="ABC123", user_id=member.id, active=True))
    db.session.commit()
    content.join_project(member, project)  # emits TEAM_JOINED -> engine runs
    progress = engine.progress(engine.subject_for_user(member))
    assert _phase(progress, "signed_up")["status"] == "complete"
    assert _phase(progress, "shop_ready")["status"] == "active"
    granted = entitlements.user_entitlements(member)
    assert {ENT_PORTAL_ACCESS, ENT_EVENT_CHECKIN} <= granted
    assert ENT_SHOP_ACCESS not in granted


def _finish_phase_one(db, member, project):
    member.major = "ME"
    member.graduation_year = 2028
    db.session.add(NFCTag(tag_uid=f"TAG-{member.id}", user_id=member.id, active=True))
    db.session.commit()
    content.join_project(member, project)


def test_shop_ready_requires_training_meeting_and_signed_off_checkout(app, users, project, item, db):
    member, lead = users["member"], users["lead"]
    _finish_phase_one(db, member, project)

    engine.record_training(member, "shop_safety", score=85, actor=lead)
    event = Event(title="GBM", kind="meeting", status="scheduled", start_time=datetime.now(), end_time=datetime.now() + timedelta(hours=1))
    db.session.add(event)
    db.session.commit()
    attendance.check_user_into_event(member, event, method="nfc")
    progress = engine.progress(engine.subject_for_user(member))
    shop = _phase(progress, "shop_ready")
    assert shop["status"] == "active"
    assert _task(shop, "supervised_checkout")["status"] == "pending"

    loan = inventory.checkout(item=item, user=member, qty=1, signed_off_by=lead)
    progress = engine.progress(engine.subject_for_user(member))
    assert _phase(progress, "shop_ready")["status"] == "complete"
    assert {ENT_SHOP_ACCESS, ENT_PRINT_SUBMIT} <= entitlements.user_entitlements(member)
    # optional machine training does not block
    assert _task(_phase(progress, "shop_ready"), "machine_training")["status"] == "pending"


def test_low_quiz_score_does_not_count(app, users, project, db):
    member, lead = users["member"], users["lead"]
    _finish_phase_one(db, member, project)
    engine.record_training(member, "shop_safety", score=60, actor=lead)
    progress = engine.progress(engine.subject_for_user(member))
    assert _task(_phase(progress, "shop_ready"), "shop_safety")["status"] == "pending"


def test_expired_training_regresses_phase_and_revokes(app, users, project, item, db):
    member, lead = users["member"], users["lead"]
    _finish_phase_one(db, member, project)
    engine.record_training(member, "shop_safety", score=90, actor=lead)
    event = Event(title="GBM", status="scheduled", start_time=datetime.now(), end_time=datetime.now() + timedelta(hours=1))
    db.session.add(event)
    db.session.commit()
    attendance.check_user_into_event(member, event)
    inventory.checkout(item=item, user=member, qty=1, signed_off_by=lead)
    assert entitlements.has_entitlement(member, ENT_SHOP_ACCESS)

    completion = TrainingCompletion.query.filter_by(user_id=member.id).first()
    completion.expires_at = datetime.utcnow() - timedelta(days=1)
    db.session.commit()
    progress = engine.evaluate_user(member)
    assert _phase(progress, "shop_ready")["status"] == "active"
    assert _phase(progress, "shop_ready")["steps_left"] == 1
    assert not entitlements.has_entitlement(member, ENT_SHOP_ACCESS)


def test_rule_driven_task_cannot_be_marked_manually(app, users):
    with pytest.raises(Validation) as excinfo:
        engine.complete_task("complete_profile", engine.subject_for_user(users["member"]), users["member"])
    assert excinfo.value.code == "task_not_manual"


def test_signoff_requires_role_and_not_self(app, users):
    member, lead, admin = users["member"], users["lead"], users["admin"]
    subject = engine.subject_for_user(member)
    with pytest.raises(Forbidden):
        engine.complete_task("nominated", subject, member)
    with pytest.raises(Forbidden):
        engine.complete_task("nominated", subject, lead)  # needs admin
    state = engine.complete_task("nominated", subject, admin, note="strong contributor")
    assert state.status == "complete" and state.completed_by_user_id == admin.id


def test_scored_training_cannot_be_self_recorded(app, users):
    with pytest.raises(Forbidden):
        engine.record_training(users["member"], "shop_safety", score=100, actor=users["member"])


def test_percent_counts_required_only(app, users, project, db):
    member = users["member"]
    _finish_phase_one(db, member, project)
    progress = engine.progress(engine.subject_for_user(member))
    track = progress["tracks"][0]
    # 4 of 13 required tasks done (the 2 optional tasks are excluded)
    assert track["required_total"] == 13
    assert track["required_done"] == 4
    assert track["percent"] == round(4 / 13 * 100)


def test_override_grant_survives_reevaluation(app, users):
    member, admin = users["member"], users["admin"]
    entitlements.grant(member, ENT_SHOP_ACCESS, source="override", source_ref="admin", granted_by=admin, reason="visiting", commit=True)
    engine.evaluate_user(member)
    assert entitlements.has_entitlement(member, ENT_SHOP_ACCESS)
    entitlements.revoke(member, ENT_SHOP_ACCESS, revoked_by=admin, commit=True)
    assert not entitlements.has_entitlement(member, ENT_SHOP_ACCESS)


def test_expiring_override(app, users):
    member, admin = users["member"], users["admin"]
    entitlements.grant(member, ENT_SHOP_ACCESS, source="override", granted_by=admin, expires_at=datetime.utcnow() - timedelta(minutes=1), commit=True)
    assert not entitlements.has_entitlement(member, ENT_SHOP_ACCESS)


def test_team_lead_phase_promotes_role(app, users, project, item, db):
    member, lead, admin = users["member"], users["lead"], users["admin"]
    _finish_phase_one(db, member, project)
    subject = engine.subject_for_user(member)
    # brute-force through phases 2 and 3 via overrides is not allowed; complete them properly but briefly
    engine.record_training(member, "shop_safety", score=95, actor=lead)
    for i in range(3):
        event = Event(title=f"GBM {i}", status="scheduled", start_time=datetime.now(), end_time=datetime.now() + timedelta(hours=1))
        db.session.add(event)
        db.session.commit()
        attendance.check_user_into_event(member, event)
    for i in range(3):
        loan = inventory.checkout(item=item, user=member, qty=1, signed_off_by=lead if i == 0 else None)
        inventory.return_loan(loan=loan)
    content.log_hours(member, project, 10)
    progress = engine.progress(subject)
    assert _phase(progress, "contributor")["status"] == "complete"
    assert _phase(progress, "team_lead")["status"] == "active"
    engine.complete_task("nominated", subject, admin)
    engine.record_training(member, "lead_essentials", actor=admin)
    engine.complete_task("build_cycle", subject, lead)
    db.session.refresh(member)
    assert member.role == "team_leader"


def test_chapter_track_counts_rows(app, users, db):
    progress = engine.evaluate_chapter()
    stand_up = _phase(progress, "stand_up")
    assert stand_up["status"] == "active"
    seed_task = _task(stand_up, "seed_inventory")
    assert seed_task["current"] == 0 and seed_task["target"] == 40
    for i in range(40):
        db.session.add(Item(name=f"Tool {i}", total_qty=1, available_qty=1, active=True, location=f"Bin {i % 6}"))
    db.session.commit()
    progress = engine.evaluate_chapter()
    stand_up = _phase(progress, "stand_up")
    assert _task(stand_up, "seed_inventory")["status"] == "complete"
    assert _task(stand_up, "shop_locations")["status"] == "complete"


def test_chapter_manual_task_needs_admin_and_note(app, users):
    subject = engine.subject_for_chapter()
    with pytest.raises(Forbidden):
        engine.complete_task("open_print_queue", subject, users["lead"], note="x")
    with pytest.raises(Validation):
        engine.complete_task("open_print_queue", subject, users["admin"])
    engine.complete_task("open_print_queue", subject, users["admin"], note="printers staffed")


def test_member_overview_lists_members_only(app, users):
    rows = engine.member_overview()
    names = {row["user"].name for row in rows}
    assert "Ada Admin" not in names
    assert {"Lee Lead", "Mo Member"} <= names
