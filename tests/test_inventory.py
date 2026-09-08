from datetime import date, timedelta

import pytest

from asme import events
from asme.models import Loan, StockDiscrepancy, StockLedger
from asme.services import inventory
from asme.services.errors import Conflict, Validation


def test_checkout_decrements_and_writes_ledger(app, users, item, captured_events):
    loan = inventory.checkout(item=item, user=users["member"], qty=2, due_date=date.today() + timedelta(days=7))
    assert loan.status == "OUT"
    assert loan.user_id == users["member"].id
    assert item.available_qty == 1
    ledger = StockLedger.query.filter_by(item_id=item.id).all()
    assert len(ledger) == 1 and ledger[0].delta_available == -2 and ledger[0].reason == "checkout"
    assert any(name == events.LOAN_OPENED for name, _ in captured_events)


def test_checkout_creates_legacy_member_link(app, users, item):
    loan = inventory.checkout(item=item, user=users["member"], qty=1)
    # the member row is optional; when absent the loan is keyed by user only
    assert loan.user_id == users["member"].id


def test_double_checkout_conflicts(app, users, item):
    inventory.checkout(item=item, user=users["member"], qty=1)
    with pytest.raises(Conflict) as excinfo:
        inventory.checkout(item=item, user=users["member"], qty=1)
    assert excinfo.value.code == "already_checked_out"


def test_insufficient_stock_is_atomic(app, users, item):
    with pytest.raises(Conflict) as excinfo:
        inventory.checkout(item=item, user=users["member"], qty=99)
    assert excinfo.value.code == "insufficient_stock"
    assert item.available_qty == 3
    assert Loan.query.count() == 0


def test_idempotency_key_replays_same_loan(app, users, item):
    first = inventory.checkout(item=item, user=users["member"], qty=1, idempotency_key="tap-123")
    second = inventory.checkout(item=item, user=users["member"], qty=1, idempotency_key="tap-123")
    assert first.id == second.id
    assert item.available_qty == 2


def test_return_restores_stock_and_marks_on_time(app, users, item, captured_events):
    loan = inventory.checkout(item=item, user=users["member"], qty=1, due_date=date.today() + timedelta(days=3))
    inventory.return_loan(loan=loan, qty=1, condition="good", user=users["member"])
    assert loan.status == "RETURNED"
    assert loan.returned_on_time
    assert item.available_qty == 3
    assert any(name == events.LOAN_RETURNED and payload["on_time"] for name, payload in captured_events)


def test_return_qty_mismatch(app, users, item):
    loan = inventory.checkout(item=item, user=users["member"], qty=2)
    with pytest.raises(Conflict) as excinfo:
        inventory.return_loan(loan=loan, qty=1)
    assert excinfo.value.code == "qty_mismatch"


def test_return_never_exceeds_total(app, users, item):
    loan = inventory.checkout(item=item, user=users["member"], qty=1)
    inventory.adjust_counts(item, total_qty=3, available_qty=3, note="miscount", actor=users["admin"])
    inventory.return_loan(loan=loan)
    assert item.available_qty == 3


def test_inactive_item_cannot_be_checked_out(app, users, item, db):
    item.active = False
    db.session.commit()
    with pytest.raises(Validation):
        inventory.checkout(item=item, user=users["member"], qty=1)


def test_reconcile_files_discrepancy_not_silent_fix(app, users, item, db):
    inventory.checkout(item=item, user=users["member"], qty=1)
    item.available_qty = 3  # someone hand-edited the counter
    db.session.commit()
    filed = inventory.reconcile_stock(actor=users["admin"])
    assert len(filed) == 1
    row = filed[0]
    assert row.counter_available == 3 and row.derived_available == 2 and row.open_loan_qty == 1
    assert item.available_qty == 3  # untouched
    inventory.resolve_discrepancy(row, users["admin"], note="fixed", apply_derived=True)
    assert item.available_qty == 2
    assert StockDiscrepancy.query.filter_by(status="open").count() == 0


def test_reconcile_is_idempotent(app, users, item, db):
    inventory.checkout(item=item, user=users["member"], qty=1)
    item.available_qty = 3
    db.session.commit()
    assert len(inventory.reconcile_stock()) == 1
    assert len(inventory.reconcile_stock()) == 0


def test_save_item_generates_tracking_ids(app, users):
    row = inventory.save_item({"name": "Torque wrench", "total_qty": "2", "available_qty": "2"}, users["admin"])
    assert row.item_code.startswith("ASME-INV-")
    assert row.treasury_tracker_id.startswith("ASME-TRK-")


def test_bulk_import_creates_and_updates(app, users):
    result = inventory.bulk_import(
        [{"name": "Hex keys", "total_qty": "5"}, {"name": "Hex keys", "total_qty": "6"}, {"name": "", "category": "orphan"}],
        users["admin"],
    )
    assert result["created"] == 1 and result["updated"] == 1 and result["skipped"] == 1


def test_legacy_transact_return_without_open_loan_restocks(app, users, item, db):
    from asme.services.identity import ensure_member_for_user

    member = ensure_member_for_user(users["member"])
    db.session.commit()
    item.available_qty = 1
    db.session.commit()
    inventory.legacy_transact(member=member, item=item, action="return", qty=1)
    assert item.available_qty == 2
