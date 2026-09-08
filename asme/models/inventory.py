from datetime import datetime

from asme.extensions import db


class Item(db.Model):
    __tablename__ = "items"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(160), nullable=False)
    description = db.Column(db.String(300), nullable=True)
    category = db.Column(db.String(80), nullable=True)
    location = db.Column(db.String(120), nullable=True)
    item_condition = db.Column(db.String(120), nullable=True)
    notes = db.Column(db.String(500), nullable=True)
    photo_url = db.Column(db.String(500), nullable=True)
    item_type = db.Column(db.String(20), nullable=False, default="tool")  # tool / consumable
    is_consumable = db.Column(db.Boolean, nullable=False, default=False)
    active = db.Column(db.Boolean, nullable=False, default=True)
    min_stock_threshold = db.Column(db.Integer, nullable=False, default=0)

    # ``available_qty`` is the fast atomic counter; ``stock_ledger`` is the audit trail
    # that lets an admin reconstruct it when the two disagree.
    total_qty = db.Column(db.Integer, nullable=False, default=0)
    available_qty = db.Column(db.Integer, nullable=False, default=0)

    nfc_tag = db.Column(db.String(120), nullable=True, unique=True)
    item_code = db.Column(db.String(40), nullable=True, unique=True, index=True)
    treasury_tracker_id = db.Column(db.String(80), nullable=True, unique=True, index=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    @property
    def checked_out_qty(self) -> int:
        return max((self.total_qty or 0) - (self.available_qty or 0), 0)


class Transaction(db.Model):
    """A loan: one row per checkout, mutated on return.

    The table name and ``action``/``status`` columns are kept for compatibility
    with historical rows; ``state`` in services means ``status`` here.
    """

    __tablename__ = "transactions"
    id = db.Column(db.Integer, primary_key=True)

    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=True, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False, index=True)

    action = db.Column(db.String(20), nullable=False)  # "checkout" or "return"
    qty = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), nullable=True, index=True)  # "OUT" or "RETURNED"

    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    checkout_time = db.Column(db.DateTime, nullable=True)
    return_time = db.Column(db.DateTime, nullable=True)
    due_date = db.Column(db.Date, nullable=True)
    notes = db.Column(db.String(300), nullable=True)
    checkout_notes = db.Column(db.String(300), nullable=True)
    return_condition = db.Column(db.String(120), nullable=True)
    return_notes = db.Column(db.String(300), nullable=True)
    return_photo_path = db.Column(db.String(500), nullable=True)
    idempotency_key = db.Column(db.String(120), nullable=True, unique=True)
    signed_off_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)

    member = db.relationship("Member")
    user = db.relationship("User", foreign_keys=[user_id])
    signed_off_by = db.relationship("User", foreign_keys=[signed_off_by_user_id])
    item = db.relationship("Item")

    @property
    def is_open(self) -> bool:
        return self.status == "OUT"

    @property
    def is_overdue(self) -> bool:
        if not self.is_open or not self.due_date:
            return False
        return self.due_date < datetime.utcnow().date()

    @property
    def returned_on_time(self) -> bool:
        if self.status != "RETURNED" or not self.return_time:
            return False
        if not self.due_date:
            return True
        return self.return_time.date() <= self.due_date

    @property
    def borrower_name(self) -> str:
        if self.user:
            return self.user.name
        if self.member:
            return self.member.name
        return ""


# Alias so services can talk about loans without renaming the historical table.
Loan = Transaction


class ItemTag(db.Model):
    __tablename__ = "item_tags"
    id = db.Column(db.Integer, primary_key=True)

    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False, index=True)
    tag_value = db.Column(db.String(160), nullable=False, unique=True, index=True)
    source = db.Column(db.String(40), nullable=False, default="uid")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    item = db.relationship("Item")


class StockLedger(db.Model):
    """Append-only record of every change to ``items.available_qty`` / ``total_qty``."""

    __tablename__ = "stock_ledger"
    id = db.Column(db.Integer, primary_key=True)

    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False, index=True)
    delta_available = db.Column(db.Integer, nullable=False, default=0)
    delta_total = db.Column(db.Integer, nullable=False, default=0)
    available_after = db.Column(db.Integer, nullable=True)
    total_after = db.Column(db.Integer, nullable=True)
    reason = db.Column(db.String(40), nullable=False)  # checkout / return / adjust / import / bootstrap
    transaction_id = db.Column(db.Integer, db.ForeignKey("transactions.id"), nullable=True, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    note = db.Column(db.String(300), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)

    item = db.relationship("Item")
    transaction = db.relationship("Transaction")
    user = db.relationship("User")


class StockDiscrepancy(db.Model):
    """Filed by the nightly reconciliation when the counter and the loans disagree."""

    __tablename__ = "stock_discrepancies"
    id = db.Column(db.Integer, primary_key=True)

    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False, index=True)
    counter_available = db.Column(db.Integer, nullable=False)
    derived_available = db.Column(db.Integer, nullable=False)
    open_loan_qty = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), nullable=False, default="open")  # open / resolved
    resolved_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    resolved_at = db.Column(db.DateTime, nullable=True)
    resolution_note = db.Column(db.String(300), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    item = db.relationship("Item")
    resolved_by = db.relationship("User")
