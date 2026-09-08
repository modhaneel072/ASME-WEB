import pytest

from asme.models import Member, NFCTag, User
from asme.services import identity
from asme.services.errors import Conflict, Forbidden, Validation


def test_signup_links_existing_member_by_email(app, db):
    member = Member(name="Mo", email="mo@uiowa.edu", member_class="Member")
    db.session.add(member)
    db.session.commit()
    user = identity.signup("Mo Member", "MO@uiowa.edu", "password123")
    assert user.member_id == member.id
    assert user.role == "member"
    assert user.username == "mo"


def test_signup_validation(app):
    with pytest.raises(Validation):
        identity.signup("M", "mo@uiowa.edu", "password123")
    with pytest.raises(Validation):
        identity.signup("Mo Member", "mo@uiowa.edu", "short")
    identity.signup("Mo Member", "mo@uiowa.edu", "password123")
    with pytest.raises(Conflict):
        identity.signup("Mo Member", "mo@uiowa.edu", "password123")


def test_signup_never_grants_admin(app):
    assert identity.signup("X Y", "x@uiowa.edu", "password123", role="admin").role == "member"


def test_authenticate_by_username_or_email(app, users):
    from tests.conftest import PASSWORD

    assert identity.authenticate("mo", PASSWORD).id == users["member"].id
    assert identity.authenticate("MO@uiowa.edu", PASSWORD).id == users["member"].id
    assert identity.authenticate("mo", "bad") is None


def test_password_reset_roundtrip(app, users):
    token = identity.create_password_reset(users["member"])
    row = identity.find_valid_reset(token)
    assert row is not None
    identity.consume_password_reset(row, "newpassword1", "newpassword1")
    assert identity.find_valid_reset(token) is None
    assert identity.authenticate("mo", "newpassword1") is not None


def test_assign_nfc_moves_tag_and_rejects_duplicates(app, users):
    identity.assign_nfc(users["member"], "TAG-1", None, users["admin"])
    assert users["member"].nfc_uid == "TAG-1"
    with pytest.raises(Conflict):
        identity.assign_nfc(users["lead"], "tag-1", None, users["admin"])
    identity.assign_nfc(users["member"], "TAG-2", None, users["admin"])
    assert NFCTag.query.filter_by(user_id=users["member"].id, active=True).count() == 1
    assert identity.resolve_user_from_tag_uid("TAG-2").id == users["member"].id
    assert identity.resolve_user_from_tag_uid("TAG-1") is None


def test_delete_user_deactivates_when_history_exists(app, users, item):
    from asme.services import inventory

    inventory.checkout(item=item, user=users["member"], qty=1)
    assert identity.admin_delete_user(users["member"], users["admin"]) == "deactivated"
    assert users["member"].is_active is False


def test_delete_user_hard_deletes_without_history(app, users):
    target_id = users["member"].id
    assert identity.admin_delete_user(users["member"], users["admin"]) == "deleted"
    assert User.query.get(target_id) is None


def test_cannot_delete_self(app, users):
    with pytest.raises(Forbidden):
        identity.admin_delete_user(users["admin"], users["admin"])


def test_admin_update_user_changes_role_and_emits(app, users, captured_events):
    identity.admin_update_user(users["member"], {"role": "team_leader", "is_active": "1", "email": "mo@uiowa.edu"}, users["admin"])
    assert users["member"].role == "team_leader"
    from asme import events

    assert any(name == events.USER_UPDATED and payload.get("role_changed") for name, payload in captured_events)


def test_fresh_credentials_assign_unique_nfc(app, users):
    rows = identity.build_fresh_member_credentials()
    assert {row["nfc_uid"] for row in rows} and len({row["nfc_uid"] for row in rows}) == len(rows)
    assert all(row["password"] for row in rows)
