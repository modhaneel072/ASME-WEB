"""Inventory: catalog, loans (checkout/return), stock adjustments, tags, imports.

Every change to ``items.available_qty`` goes through this module and writes a
``StockLedger`` row in the same transaction, so the counter can always be
reconciled against the loans.
"""

from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime

from sqlalchemy import case, func
from werkzeug.utils import secure_filename

from asme import events
from asme.constants import (
    BULK_INVENTORY_ALLOWED_EXTENSIONS,
    INVENTORY_ITEM_CODE_PREFIX,
    ITEM_TYPES,
    LOAN_STATE_OUT,
    LOAN_STATE_RETURNED,
    TAG_PREFIXES,
    TREASURY_TRACKER_PREFIX,
)
from asme.extensions import db
from asme.models import Item, ItemTag, Loan, Member, NFCTag, StockDiscrepancy, StockLedger, Transaction, User
from asme.services import audit
from asme.services.errors import Conflict, NotFound, Validation
from asme.utils import clean_tag_value, normalize_text_key, parse_bool_flag, parse_non_negative_int

# --------------------------------------------------------------------------- tags & ids


def parse_item_id_from_tag(tag):
    raw = clean_tag_value(tag).lower()
    for prefix in TAG_PREFIXES:
        if raw.startswith(prefix):
            try:
                return int(raw.split(":", 1)[1].strip())
            except Exception:
                return None
    return None


def find_item_by_tag(tag):
    cleaned_tag = clean_tag_value(tag)
    if not cleaned_tag:
        return None, "empty"
    tag_item_id = parse_item_id_from_tag(cleaned_tag)
    if tag_item_id:
        item = db.session.get(Item, tag_item_id)
        if item:
            return item, "payload_item_id"
    mapped = ItemTag.query.filter(func.lower(ItemTag.tag_value) == cleaned_tag.lower()).first()
    if mapped and mapped.item:
        return mapped.item, "item_tags"
    legacy_item = Item.query.filter(func.lower(Item.nfc_tag) == cleaned_tag.lower()).first()
    if legacy_item:
        return legacy_item, "legacy_item_nfc_tag"
    return None, "not_found"


def resolve_item(item_tag, item_id):
    item_tag = clean_tag_value(item_tag)
    item_id = str(item_id or "").strip()
    if item_tag:
        item, _via = find_item_by_tag(item_tag)
        if item:
            return item
    if item_id:
        try:
            return db.session.get(Item, int(item_id))
        except Exception:
            return None
    return None


def normalize_inventory_item_code(raw_value):
    cleaned = clean_tag_value(raw_value).upper()
    if not cleaned:
        return ""
    cleaned = re.sub(r"[^A-Z0-9\-]+", "-", cleaned)
    cleaned = re.sub(r"-{2,}", "-", cleaned).strip("-")
    return cleaned[:40]


def normalize_treasury_tracker_id(raw_value):
    cleaned = clean_tag_value(raw_value).upper()
    if not cleaned:
        return ""
    cleaned = re.sub(r"[^A-Z0-9\-]+", "-", cleaned)
    cleaned = re.sub(r"-{2,}", "-", cleaned).strip("-")
    return cleaned[:80]


def generate_inventory_item_code(item_id):
    try:
        numeric_id = int(item_id)
    except Exception:
        numeric_id = 0
    return f"{INVENTORY_ITEM_CODE_PREFIX}{max(numeric_id, 0):05d}"


def generate_treasury_tracker_id(item_code, item_id=None):
    normalized_code = normalize_inventory_item_code(item_code) or generate_inventory_item_code(item_id or 0)
    suffix = normalized_code
    if suffix.startswith(INVENTORY_ITEM_CODE_PREFIX):
        suffix = suffix[len(INVENTORY_ITEM_CODE_PREFIX) :]
    return normalize_treasury_tracker_id(f"{TREASURY_TRACKER_PREFIX}{suffix}")


def ensure_item_tracking_ids(item):
    if not item:
        return
    if not item.id:
        db.session.flush()
    item_code = normalize_inventory_item_code(item.item_code) or generate_inventory_item_code(item.id)
    base_code = item_code
    suffix = 2
    while (
        Item.query.filter(func.lower(func.coalesce(Item.item_code, "")) == item_code.lower(), Item.id != item.id)
        .order_by(Item.id.asc())
        .first()
    ):
        item_code = normalize_inventory_item_code(f"{base_code}-{suffix}")
        suffix += 1
    item.item_code = item_code

    tracker_id = normalize_treasury_tracker_id(item.treasury_tracker_id) or generate_treasury_tracker_id(item.item_code, item.id)
    base_tracker = tracker_id
    suffix = 2
    while (
        Item.query.filter(
            func.lower(func.coalesce(Item.treasury_tracker_id, "")) == tracker_id.lower(), Item.id != item.id
        )
        .order_by(Item.id.asc())
        .first()
    ):
        tracker_id = normalize_treasury_tracker_id(f"{base_tracker}-{suffix}")
        suffix += 1
    item.treasury_tracker_id = tracker_id


def get_item_primary_tag_map():
    mapping = {}
    for row in ItemTag.query.order_by(ItemTag.id.asc()).all():
        cleaned = clean_tag_value(row.tag_value)
        if not cleaned or row.item_id in mapping:
            continue
        mapping[row.item_id] = cleaned
    return mapping


def get_item_primary_nfc(item, tag_map=None):
    if not item:
        return ""
    direct_tag = clean_tag_value(item.nfc_tag)
    if direct_tag:
        return direct_tag
    return clean_tag_value((tag_map or {}).get(item.id))


def validate_item_tag_uid(tag_uid, current_item_id=None):
    """Return an error message if ``tag_uid`` is already used anywhere, else ``""``."""
    cleaned = clean_tag_value(tag_uid)
    if not cleaned:
        return ""
    if Member.query.filter(func.lower(Member.nfc_tag) == cleaned.lower()).first():
        return "That NFC UID is already assigned to a member profile."
    if User.query.filter(func.lower(User.nfc_uid) == cleaned.lower()).first():
        return "That NFC UID is already assigned to a user account."
    if NFCTag.query.filter(func.lower(NFCTag.tag_uid) == cleaned.lower(), NFCTag.active.is_(True)).first():
        return "That NFC UID is already assigned to a member login tag."
    if Item.query.filter(
        func.lower(func.coalesce(Item.nfc_tag, "")) == cleaned.lower(), Item.id != (current_item_id or 0)
    ).first():
        return "That NFC UID is already assigned to another inventory item."
    if ItemTag.query.filter(
        func.lower(ItemTag.tag_value) == cleaned.lower(), ItemTag.item_id != (current_item_id or 0)
    ).first():
        return "That NFC UID is already assigned in the item tag map."
    return ""


def set_item_primary_nfc(item, tag_uid, source="inventory_admin"):
    cleaned = clean_tag_value(tag_uid)
    item.nfc_tag = cleaned or None
    if not cleaned:
        return
    existing = ItemTag.query.filter(func.lower(ItemTag.tag_value) == cleaned.lower()).first()
    if existing:
        existing.item_id = item.id
        if not existing.source:
            existing.source = source[:40]
        return
    db.session.add(ItemTag(item_id=item.id, tag_value=cleaned, source=source[:40]))


def register_item_tag(item, tag_value, source="manual"):
    """Legacy NFC admin: map an additional tag to an item."""
    tag_value = clean_tag_value(tag_value)
    if not tag_value:
        raise Validation("Tag value is required.", field="tag_value")
    existing = ItemTag.query.filter(func.lower(ItemTag.tag_value) == tag_value.lower()).first()
    if existing and existing.item_id != item.id:
        raise Conflict(f"That tag is already assigned to {existing.item.name}.", code="tag_taken")
    if existing:
        existing.source = source
    else:
        db.session.add(ItemTag(item_id=item.id, tag_value=tag_value, source=source))
    if not clean_tag_value(item.nfc_tag):
        item.nfc_tag = tag_value
    db.session.commit()


# --------------------------------------------------------------------------- ledger


def _ledger(item, *, delta_available=0, delta_total=0, reason, loan=None, user=None, note=None):
    db.session.add(
        StockLedger(
            item_id=item.id,
            delta_available=delta_available,
            delta_total=delta_total,
            available_after=item.available_qty,
            total_after=item.total_qty,
            reason=reason,
            transaction_id=loan.id if loan else None,
            user_id=user.id if user else None,
            note=(note or "")[:300] or None,
        )
    )


# --------------------------------------------------------------------------- loans


def get_open_loan(item_id, user=None, member=None):
    """Most recent OUT loan for this item by this person (user or legacy member)."""
    query = Loan.query.filter(Loan.item_id == item_id, Loan.status == LOAN_STATE_OUT)
    if user and member:
        query = query.filter((Loan.user_id == user.id) | (Loan.member_id == member.id))
    elif user:
        query = query.filter(Loan.user_id == user.id)
    elif member:
        query = query.filter(Loan.member_id == member.id)
    else:
        return None
    return query.order_by(Loan.checkout_time.desc(), Loan.id.desc()).first()


def open_loans_for(user=None, member=None):
    query = Loan.query.filter(Loan.status == LOAN_STATE_OUT)
    if user and member:
        query = query.filter((Loan.user_id == user.id) | (Loan.member_id == member.id))
    elif user:
        query = query.filter(Loan.user_id == user.id)
    elif member:
        query = query.filter(Loan.member_id == member.id)
    else:
        return []
    return query.order_by(Loan.checkout_time.desc(), Loan.id.desc()).all()


def checkout(
    *,
    item,
    user=None,
    member=None,
    qty=1,
    notes=None,
    due_date=None,
    idempotency_key=None,
    signed_off_by=None,
    source="portal",
) -> Loan:
    """Open a loan. The stock decrement is a single conditional UPDATE, so two
    concurrent checkouts cannot both succeed on the last unit."""
    if user is None and member is None:
        raise Validation("A borrower is required.", code="borrower_required")
    if member is None and user is not None:
        from asme.services.identity import member_for_user

        member = member_for_user(user)
    if user is None and member is not None:
        from asme.services.identity import user_for_member

        user = user_for_member(member)
    qty = int(qty or 0)
    if qty <= 0:
        raise Validation("Quantity must be at least 1.", field="qty")
    if not item.active:
        raise Validation(f"{item.name} is not available for checkout.", code="item_inactive")

    if idempotency_key:
        existing = Loan.query.filter_by(idempotency_key=idempotency_key).first()
        if existing:
            return existing

    if get_open_loan(item.id, user=user, member=member):
        raise Conflict(
            "You already have this item checked out. Return it before checking out again.",
            code="already_checked_out",
        )

    updated = (
        Item.query.filter(Item.id == item.id, Item.available_qty >= qty)
        .update({Item.available_qty: Item.available_qty - qty}, synchronize_session=False)
    )
    if updated != 1:
        db.session.rollback()
        db.session.refresh(item)
        raise Conflict(
            f"Not enough stock. {item.name} has {item.available_qty} available.",
            code="insufficient_stock",
            available=item.available_qty,
        )

    now = datetime.utcnow()
    loan = Loan(
        member_id=member.id if member else None,
        user_id=user.id if user else None,
        item_id=item.id,
        action="checkout",
        qty=qty,
        status=LOAN_STATE_OUT,
        timestamp=now,
        checkout_time=now,
        due_date=due_date,
        notes=notes,
        checkout_notes=notes,
        idempotency_key=idempotency_key,
        signed_off_by_user_id=signed_off_by.id if signed_off_by else None,
    )
    db.session.add(loan)
    db.session.flush()
    db.session.refresh(item)
    _ledger(item, delta_available=-qty, reason="checkout", loan=loan, user=user, note=f"source={source}")
    db.session.commit()
    events.emit(
        events.LOAN_OPENED,
        loan_id=loan.id,
        user_id=user.id if user else None,
        member_id=member.id if member else None,
        item_id=item.id,
        signed_off=bool(signed_off_by),
    )
    return loan


def return_loan(*, loan, qty=None, condition=None, notes=None, photo_path=None, user=None, source="portal") -> Loan:
    if loan.status != LOAN_STATE_OUT:
        raise Conflict("That loan is already closed.", code="loan_closed")
    if qty is not None and int(qty) != loan.qty:
        raise Conflict(
            f"Return quantity must match the checked-out quantity ({loan.qty}).",
            code="qty_mismatch",
            expected=loan.qty,
        )
    qty = loan.qty
    item = loan.item
    capped = case((Item.available_qty + qty > Item.total_qty, Item.total_qty), else_=Item.available_qty + qty)
    Item.query.filter(Item.id == item.id).update({Item.available_qty: capped}, synchronize_session=False)

    now = datetime.utcnow()
    loan.status = LOAN_STATE_RETURNED
    if user and not loan.user_id:
        loan.user_id = user.id
    loan.return_time = now
    loan.return_condition = (condition or "good")[:120]
    loan.return_notes = notes
    loan.return_photo_path = photo_path
    db.session.flush()
    db.session.refresh(item)
    _ledger(item, delta_available=qty, reason="return", loan=loan, user=user or loan.user, note=f"source={source}")
    db.session.commit()
    events.emit(
        events.LOAN_RETURNED,
        loan_id=loan.id,
        user_id=loan.user_id,
        member_id=loan.member_id,
        item_id=item.id,
        on_time=loan.returned_on_time,
    )
    return loan


def sign_off_loan(loan, leader) -> Loan:
    """A team lead vouches for a supervised checkout (Launchpad phase 2)."""
    loan.signed_off_by_user_id = leader.id
    db.session.commit()
    events.emit(events.LOAN_OPENED, loan_id=loan.id, user_id=loan.user_id, member_id=loan.member_id, item_id=loan.item_id, signed_off=True)
    return loan


def legacy_transact(*, member, item, action, qty, notes=None, due_date=None, auth_user=None):
    """Legacy kiosk endpoint semantics: checkout or a free-form return."""
    if action == "checkout":
        return checkout(item=item, member=member, user=auth_user, qty=qty, notes=notes, due_date=due_date, source="kiosk")
    if action != "return":
        raise Validation("Invalid inventory action.")
    open_loan = get_open_loan(item.id, member=member)
    if open_loan and open_loan.qty == qty:
        return return_loan(loan=open_loan, condition="manual-return", notes=notes, user=auth_user, source="kiosk")
    # No matching open loan: record a manual restock so the counter still moves.
    capped = case((Item.available_qty + qty > Item.total_qty, Item.total_qty), else_=Item.available_qty + qty)
    Item.query.filter(Item.id == item.id).update({Item.available_qty: capped}, synchronize_session=False)
    now = datetime.utcnow()
    tx = Transaction(
        member_id=member.id,
        user_id=auth_user.id if auth_user else None,
        item_id=item.id,
        action="return",
        qty=qty,
        status=LOAN_STATE_RETURNED,
        timestamp=now,
        return_time=now,
        notes=notes,
        return_condition="manual-return",
        return_notes=notes,
    )
    db.session.add(tx)
    db.session.flush()
    db.session.refresh(item)
    _ledger(item, delta_available=qty, reason="return", loan=tx, user=auth_user, note="source=kiosk manual-return")
    db.session.commit()
    return tx


# --------------------------------------------------------------------------- admin: items


def save_item(form, actor) -> Item:
    from asme.utils import parse_positive_int

    item_id = parse_positive_int(form.get("item_id"), default=0)
    name = (form.get("name") or "").strip()
    if not name:
        raise Validation("Item name is required.", field="name")
    item = db.session.get(Item, item_id) if item_id else None
    if not item:
        item = Item(name=name[:160], total_qty=0, available_qty=0)
        db.session.add(item)
        db.session.flush()

    nfc_tag = clean_tag_value(form.get("nfc_tag"))
    tag_error = validate_item_tag_uid(nfc_tag, current_item_id=item.id)
    if tag_error:
        db.session.rollback()
        raise Conflict(tag_error, code="nfc_taken")

    submitted_item_code = normalize_inventory_item_code(form.get("item_code"))
    if submitted_item_code and Item.query.filter(
        func.lower(func.coalesce(Item.item_code, "")) == submitted_item_code.lower(), Item.id != item.id
    ).first():
        db.session.rollback()
        raise Conflict("Item ID is already in use. Choose a different Item ID.", code="item_code_taken")
    submitted_tracker_id = normalize_treasury_tracker_id(form.get("treasury_tracker_id"))
    if submitted_tracker_id and Item.query.filter(
        func.lower(func.coalesce(Item.treasury_tracker_id, "")) == submitted_tracker_id.lower(), Item.id != item.id
    ).first():
        db.session.rollback()
        raise Conflict("Tracker ID is already in use. Choose a different tracker ID.", code="tracker_taken")

    previous_total = item.total_qty or 0
    previous_available = item.available_qty or 0
    total_qty = parse_non_negative_int(form.get("total_qty"), default=max(previous_total, 0))
    available_qty = parse_non_negative_int(form.get("available_qty"), default=min(previous_available, total_qty))
    item_type = (form.get("item_type") or "").strip().lower() or "tool"
    if item_type not in ITEM_TYPES:
        item_type = "tool"
    item.name = name[:160]
    item.description = (form.get("description") or "").strip()[:300] or None
    item.category = (form.get("category") or "").strip()[:80] or None
    item.location = (form.get("location") or "").strip()[:120] or None
    item.item_condition = (form.get("item_condition") or "").strip()[:120] or None
    item.notes = (form.get("notes") or "").strip()[:500] or None
    item.photo_url = (form.get("photo_url") or "").strip()[:500] or None
    item.item_type = item_type
    item.is_consumable = item_type == "consumable"
    item.active = (form.get("active") or "1").strip() in {"1", "true", "yes", "on"}
    item.min_stock_threshold = parse_non_negative_int(form.get("min_stock_threshold"), default=0)
    item.total_qty = max(total_qty, 0)
    item.available_qty = max(0, min(available_qty, item.total_qty))
    if submitted_item_code:
        item.item_code = submitted_item_code
    if submitted_tracker_id:
        item.treasury_tracker_id = submitted_tracker_id
    ensure_item_tracking_ids(item)
    set_item_primary_nfc(item, nfc_tag, source="inventory_admin")
    if item.total_qty != previous_total or item.available_qty != previous_available:
        _ledger(
            item,
            delta_available=item.available_qty - previous_available,
            delta_total=item.total_qty - previous_total,
            reason="edit",
            user=actor,
        )
    audit.record(
        "save_inventory_item",
        f"item_id={item.id} name={item.name} item_code={item.item_code} tracker_id={item.treasury_tracker_id}",
        actor=actor,
    )
    db.session.commit()
    events.emit(events.CHAPTER_CHANGED, reason="item_saved")
    return item


def adjust_counts(item, total_qty, available_qty, note, actor) -> Item:
    previous_total = item.total_qty or 0
    previous_available = item.available_qty or 0
    item.total_qty = max(0, int(total_qty))
    item.available_qty = max(0, min(int(available_qty), item.total_qty))
    note = (note or "").strip() or "Admin quantity correction"
    # Legacy: a zero-qty correction row keeps the activity feed honest.
    from asme.services.identity import member_for_user

    admin_member = member_for_user(actor) if actor else None
    correction = Transaction(
        member_id=admin_member.id if admin_member else None,
        user_id=actor.id if actor else None,
        item_id=item.id,
        action="return",
        qty=0,
        status=LOAN_STATE_RETURNED,
        timestamp=datetime.utcnow(),
        return_time=datetime.utcnow(),
        return_condition="admin-correction",
        return_notes=note[:300],
        notes=note[:300],
    )
    db.session.add(correction)
    db.session.flush()
    _ledger(
        item,
        delta_available=item.available_qty - previous_available,
        delta_total=item.total_qty - previous_total,
        reason="adjust",
        loan=correction,
        user=actor,
        note=note,
    )
    audit.record("adjust_inventory", f"item_id={item.id} total={item.total_qty} available={item.available_qty}", actor=actor)
    db.session.commit()
    return item


def bootstrap_counts(starter_qty, actor) -> int:
    starter_qty = max(1, int(starter_qty or 1))
    updated = 0
    touched = []
    for item in Item.query.order_by(Item.id.asc()).all():
        prev_total, prev_avail = item.total_qty or 0, item.available_qty or 0
        changed = False
        if prev_total <= 0:
            item.total_qty = starter_qty
            changed = True
        if prev_avail <= 0:
            item.available_qty = min(item.total_qty or starter_qty, starter_qty)
            changed = True
        if item.available_qty > item.total_qty:
            item.available_qty = item.total_qty
            changed = True
        if changed:
            updated += 1
            touched.append(f"{item.id}:{item.name}")
            _ledger(
                item,
                delta_available=item.available_qty - prev_avail,
                delta_total=item.total_qty - prev_total,
                reason="bootstrap",
                user=actor,
            )
    audit.record(
        "bootstrap_inventory_counts",
        f"starter_qty={starter_qty} updated_items={updated} ids={','.join(touched[:40])}",
        actor=actor,
    )
    db.session.commit()
    return updated


def reset_inventory(actor) -> dict:
    ledger_deleted = StockLedger.query.delete(synchronize_session=False)
    StockDiscrepancy.query.delete(synchronize_session=False)
    tx_deleted = Transaction.query.delete(synchronize_session=False)
    tags_deleted = ItemTag.query.delete(synchronize_session=False)
    items_deleted = Item.query.delete(synchronize_session=False)
    audit.record(
        "reset_inventory",
        f"deleted_items={items_deleted} deleted_tags={tags_deleted} deleted_transactions={tx_deleted} deleted_ledger={ledger_deleted}",
        actor=actor,
    )
    db.session.commit()
    events.emit(events.CHAPTER_CHANGED, reason="inventory_reset")
    return {"items": items_deleted, "tags": tags_deleted, "transactions": tx_deleted}


# --------------------------------------------------------------------------- bulk import


def bulk_row_value(row, *aliases):
    normalized = {normalize_text_key(key): value for key, value in (row or {}).items()}
    for alias in aliases:
        candidate = normalized.get(normalize_text_key(alias))
        if candidate is None:
            continue
        text_value = str(candidate).strip()
        if text_value:
            return text_value
    return ""


def parse_bulk_inventory_rows(upload_file):
    if not upload_file or not upload_file.filename:
        raise Validation("Upload a CSV or Excel file first.")
    filename = secure_filename(upload_file.filename)
    if "." not in filename:
        raise Validation("Unsupported file type. Use CSV or Excel.")
    extension = filename.rsplit(".", 1)[1].lower()
    if extension not in BULK_INVENTORY_ALLOWED_EXTENSIONS:
        raise Validation("Unsupported file type. Use CSV, XLSX, or XLS.")

    if extension == "csv":
        raw_bytes = upload_file.read()
        if not raw_bytes:
            raise Validation("The uploaded file is empty.")
        try:
            text_blob = raw_bytes.decode("utf-8-sig")
        except Exception:
            text_blob = raw_bytes.decode("latin-1", errors="ignore")
        reader = csv.DictReader(io.StringIO(text_blob))
        if not reader.fieldnames:
            raise Validation("CSV file must include a header row.")
        return [dict(row) for row in reader]

    try:
        import pandas as pd
    except Exception as exc:  # pragma: no cover
        raise Validation("Excel import requires pandas and openpyxl in requirements.") from exc
    upload_file.stream.seek(0)
    frame = pd.read_excel(upload_file)
    if frame.empty:
        raise Validation("The uploaded file is empty.")
    return frame.fillna("").to_dict(orient="records")


def bulk_import(rows, actor) -> dict:
    created_count = updated_count = skipped_count = 0
    errors = []

    for index, row in enumerate(rows, start=2):
        if not any(str(value).strip() for value in (row or {}).values()):
            continue
        name = bulk_row_value(row, "name", "item", "item_name")
        if not name:
            skipped_count += 1
            errors.append(f"Row {index}: missing item name.")
            continue

        item_code_input = normalize_inventory_item_code(bulk_row_value(row, "item_code", "item id", "inventory_id", "asset_id"))
        tracker_input = normalize_treasury_tracker_id(
            bulk_row_value(row, "treasury_tracker_id", "tracker_id", "tracker", "treasury_id")
        )
        nfc_tag_input = clean_tag_value(bulk_row_value(row, "nfc_tag", "nfc", "nfc_uid", "tag_uid", "tag"))

        item = None
        if item_code_input:
            item = Item.query.filter(func.lower(func.coalesce(Item.item_code, "")) == item_code_input.lower()).first()
        if not item and nfc_tag_input:
            item = Item.query.filter(func.lower(func.coalesce(Item.nfc_tag, "")) == nfc_tag_input.lower()).first()
            if not item:
                mapped_tag = ItemTag.query.filter(func.lower(ItemTag.tag_value) == nfc_tag_input.lower()).first()
                if mapped_tag:
                    item = mapped_tag.item
        if not item:
            item = Item.query.filter(func.lower(Item.name) == name.lower()).order_by(Item.id.asc()).first()

        is_new = item is None
        if is_new:
            item = Item(name=name[:160], total_qty=0, available_qty=0)

        if item_code_input and Item.query.filter(
            func.lower(func.coalesce(Item.item_code, "")) == item_code_input.lower(), Item.id != (item.id or 0)
        ).first():
            skipped_count += 1
            errors.append(f"Row {index}: item code '{item_code_input}' is already used.")
            continue
        if tracker_input and Item.query.filter(
            func.lower(func.coalesce(Item.treasury_tracker_id, "")) == tracker_input.lower(), Item.id != (item.id or 0)
        ).first():
            skipped_count += 1
            errors.append(f"Row {index}: tracker ID '{tracker_input}' is already used.")
            continue
        tag_error = validate_item_tag_uid(nfc_tag_input, current_item_id=item.id if item.id else None)
        if tag_error:
            skipped_count += 1
            errors.append(f"Row {index}: {tag_error}")
            continue

        total_raw = bulk_row_value(
            row, "total_qty", "total", "quantity_total", "qty_total", "qty", "quantity", "count", "amount",
            "item_amount", "item amount", "item amount?", "item_count", "amount_on_hand", "on_hand",
        )
        total_default = 1 if is_new else max(item.total_qty, 0)
        total_qty = parse_non_negative_int(total_raw, default=total_default)
        available_raw = bulk_row_value(
            row, "available_qty", "available", "quantity_available", "qty_available", "available_count", "in_stock", "stock"
        )
        if available_raw:
            available_qty = parse_non_negative_int(available_raw, default=total_qty)
        else:
            available_qty = total_qty if is_new else min(max(item.available_qty, 0), total_qty)

        item_type_raw = (bulk_row_value(row, "item_type", "type", "inventory_type") or "").strip().lower()
        if item_type_raw not in ITEM_TYPES:
            item_type_raw = item.item_type if item.item_type in ITEM_TYPES else "tool"
        prev_total, prev_avail = (item.total_qty or 0), (item.available_qty or 0)

        item.name = name[:160]
        item.category = bulk_row_value(row, "category")[:80] or None
        item.location = bulk_row_value(row, "location", "bin", "shelf")[:120] or None
        item.description = bulk_row_value(row, "description")[:300] or None
        item.item_condition = bulk_row_value(row, "item_condition", "condition")[:120] or None
        item.notes = bulk_row_value(row, "notes")[:500] or None
        item.photo_url = bulk_row_value(row, "photo_url", "photo", "image_url")[:500] or None
        item.item_type = item_type_raw
        item.is_consumable = item_type_raw == "consumable"
        item.active = parse_bool_flag(bulk_row_value(row, "active", "is_active", "enabled"), default=True if is_new else bool(item.active))
        item.min_stock_threshold = parse_non_negative_int(
            bulk_row_value(row, "min_stock_threshold", "min_stock", "reorder_threshold"),
            default=max(item.min_stock_threshold or 0, 0),
        )
        item.total_qty = max(total_qty, 0)
        item.available_qty = max(0, min(available_qty, item.total_qty))
        if item_code_input:
            item.item_code = item_code_input
        if tracker_input:
            item.treasury_tracker_id = tracker_input
        if is_new:
            db.session.add(item)
            db.session.flush()
        ensure_item_tracking_ids(item)
        set_item_primary_nfc(item, nfc_tag_input, source="bulk_inventory_import")
        if is_new or item.total_qty != prev_total or item.available_qty != prev_avail:
            _ledger(
                item,
                delta_available=item.available_qty - prev_avail,
                delta_total=item.total_qty - prev_total,
                reason="import",
                user=actor,
            )
        if is_new:
            created_count += 1
        else:
            updated_count += 1

    if not created_count and not updated_count:
        db.session.rollback()
        raise Validation("Bulk import made no changes. Fix file issues and try again.", errors=errors[:8])

    audit.record(
        "bulk_inventory_import",
        f"rows={len(rows)} created={created_count} updated={updated_count} skipped={skipped_count}",
        actor=actor,
    )
    db.session.commit()
    events.emit(events.CHAPTER_CHANGED, reason="inventory_import")
    return {"created": created_count, "updated": updated_count, "skipped": skipped_count, "errors": errors}


# --------------------------------------------------------------------------- reports


def treasury_report_rows():
    tag_map = get_item_primary_tag_map()
    rows = []
    for item in Item.query.order_by(Item.name.asc(), Item.id.asc()).all():
        rows.append(
            [
                item.item_code or "",
                item.treasury_tracker_id or "",
                item.id,
                item.name or "",
                get_item_primary_nfc(item, tag_map) or "",
                item.category or "",
                item.location or "",
                item.total_qty or 0,
                item.available_qty or 0,
                item.checked_out_qty,
                item.min_stock_threshold or 0,
                "yes" if item.active else "no",
            ]
        )
    return rows


TREASURY_REPORT_HEADERS = [
    "item_code", "treasury_tracker_id", "db_item_id", "item_name", "nfc_id", "category", "location",
    "total_qty", "available_qty", "checked_out_qty", "min_stock_threshold", "active",
]


def item_nfc_pdf_bytes() -> bytes:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    items = Item.query.order_by(Item.name.asc(), Item.id.asc()).all()
    tag_map = get_item_primary_tag_map()
    now = datetime.now()
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    page_width, page_height = letter
    margin_left = 36
    y = page_height - 44

    def draw_header():
        nonlocal y
        pdf.setFont("Helvetica-Bold", 13)
        pdf.drawString(margin_left, y, "ASME Inventory NFC / Tracker List")
        pdf.setFont("Helvetica", 9)
        pdf.drawString(margin_left, y - 13, f"Generated: {now.strftime('%Y-%m-%d %H:%M')}")
        y -= 30
        pdf.setFont("Helvetica-Bold", 8.5)
        for offset, label in ((0, "Item Code"), (76, "Item Name"), (292, "NFC ID"), (430, "Tracker ID")):
            pdf.drawString(margin_left + offset, y, label)
        y -= 8
        pdf.line(margin_left, y, page_width - margin_left, y)
        y -= 12

    draw_header()
    pdf.setFont("Helvetica", 8.5)
    for item in items:
        if y < 36:
            pdf.showPage()
            y = page_height - 44
            draw_header()
            pdf.setFont("Helvetica", 8.5)
        pdf.drawString(margin_left, y, (item.item_code or "")[:18])
        pdf.drawString(margin_left + 76, y, (item.name or "")[:42])
        pdf.drawString(margin_left + 292, y, (get_item_primary_nfc(item, tag_map) or "-")[:25])
        pdf.drawString(margin_left + 430, y, (item.treasury_tracker_id or "")[:26])
        y -= 13
    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()


# --------------------------------------------------------------------------- reconciliation


def reconcile_stock(actor=None) -> list[StockDiscrepancy]:
    """Compare ``available_qty`` with ``total_qty - sum(open loans)`` and file
    discrepancies for admin review. Never corrects silently."""
    open_by_item = dict(
        db.session.query(Loan.item_id, func.coalesce(func.sum(Loan.qty), 0))
        .filter(Loan.status == LOAN_STATE_OUT)
        .group_by(Loan.item_id)
        .all()
    )
    filed = []
    for item in Item.query.order_by(Item.id.asc()).all():
        open_qty = int(open_by_item.get(item.id, 0) or 0)
        derived = max((item.total_qty or 0) - open_qty, 0)
        if derived == (item.available_qty or 0):
            continue
        already_open = StockDiscrepancy.query.filter_by(item_id=item.id, status="open").first()
        if already_open:
            already_open.counter_available = item.available_qty or 0
            already_open.derived_available = derived
            already_open.open_loan_qty = open_qty
            continue
        row = StockDiscrepancy(
            item_id=item.id,
            counter_available=item.available_qty or 0,
            derived_available=derived,
            open_loan_qty=open_qty,
        )
        db.session.add(row)
        filed.append(row)
    if filed:
        audit.record("stock_reconciliation", f"discrepancies_filed={len(filed)}", actor=actor)
    db.session.commit()
    return filed


def resolve_discrepancy(row, actor, note=None, apply_derived=False):
    if apply_derived:
        item = row.item
        prev = item.available_qty or 0
        item.available_qty = max(0, min(row.derived_available, item.total_qty or 0))
        _ledger(item, delta_available=item.available_qty - prev, reason="reconcile", user=actor, note=note)
    row.status = "resolved"
    row.resolved_by_user_id = actor.id if actor else None
    row.resolved_at = datetime.utcnow()
    row.resolution_note = (note or "")[:300] or None
    audit.record("resolve_stock_discrepancy", f"discrepancy_id={row.id} apply_derived={apply_derived}", actor=actor)
    db.session.commit()
    return row


def overdue_loans():
    return (
        Loan.query.filter(Loan.status == LOAN_STATE_OUT, Loan.due_date.isnot(None), Loan.due_date < date.today())
        .order_by(Loan.due_date.asc(), Loan.id.asc())
        .all()
    )


def low_stock_items():
    return (
        Item.query.filter(Item.active.is_(True), Item.available_qty <= Item.min_stock_threshold)
        .order_by(Item.available_qty.asc(), Item.name.asc(), Item.id.asc())
        .all()
    )
