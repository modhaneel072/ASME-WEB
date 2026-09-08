"""JSON shapes for the API blueprints."""

from __future__ import annotations

from flask import url_for

from asme.auth.session import normalize_role
from asme.utils import iso_or_none


def serialize_member(member):
    if not member:
        return None
    return {
        "id": member.id,
        "name": member.name,
        "email": member.email,
        "member_class": member.member_class,
        "nfc_tag": member.nfc_tag,
        "created_at": iso_or_none(member.created_at),
    }


def serialize_user(user, entitlements=None):
    if not user:
        return None
    payload = {
        "id": user.id,
        "name": user.name,
        "email": user.email,
        "username": user.username or "",
        "role": normalize_role(user.role),
        "is_active": bool(user.is_active),
        "member_id": user.member_id,
        "major": user.major,
        "graduation_year": user.graduation_year,
        "nfc_linked": bool(user.nfc_uid),
        "created_at": iso_or_none(user.created_at),
        "last_login_at": iso_or_none(user.last_login_at),
    }
    if entitlements is not None:
        payload["entitlements"] = sorted(entitlements)
    return payload


def serialize_item(item):
    return {
        "id": item.id,
        "name": item.name,
        "description": item.description,
        "category": item.category,
        "location": item.location,
        "item_code": item.item_code,
        "item_type": item.item_type,
        "active": bool(item.active),
        "total_qty": item.total_qty,
        "available_qty": item.available_qty,
        "min_stock_threshold": item.min_stock_threshold,
        "nfc_tag": item.nfc_tag,
        "low_stock": item.available_qty <= max(item.min_stock_threshold or 0, 2),
        "out_of_stock": item.available_qty <= 0,
        "created_at": iso_or_none(item.created_at),
    }


def serialize_loan(tx):
    return {
        "id": tx.id,
        "timestamp": iso_or_none(tx.timestamp),
        "member_id": tx.member_id,
        "member_name": tx.borrower_name,
        "user_id": tx.user_id,
        "item_id": tx.item_id,
        "item_name": tx.item.name if tx.item else "",
        "action": tx.action,
        "status": tx.status,
        "state": "open" if tx.status == "OUT" else "returned",
        "qty": tx.qty,
        "checkout_time": iso_or_none(tx.checkout_time),
        "return_time": iso_or_none(tx.return_time),
        "due_date": iso_or_none(tx.due_date),
        "overdue": tx.is_overdue,
        "notes": tx.notes,
        "checkout_notes": tx.checkout_notes,
        "return_condition": tx.return_condition,
        "return_notes": tx.return_notes,
        "return_photo_path": tx.return_photo_path,
        "signed_off_by_user_id": tx.signed_off_by_user_id,
    }


serialize_transaction = serialize_loan


def serialize_open_checkout(tx):
    checkout_at = tx.checkout_time or tx.timestamp
    return {
        "transaction_id": tx.id,
        "item_id": tx.item_id,
        "item_name": tx.item.name if tx.item else "",
        "item_category": tx.item.category if tx.item else None,
        "item_location": tx.item.location if tx.item else None,
        "qty": tx.qty,
        "checkout_time": iso_or_none(checkout_at),
        "checkout_notes": tx.checkout_notes or tx.notes,
        "due_date": iso_or_none(tx.due_date),
        "status": tx.status,
    }


def serialize_attendance_scan(scan):
    return {
        "id": scan.id,
        "member_id": scan.member_id,
        "member_name": scan.member.name if scan.member else "",
        "uid": scan.scanned_uid,
        "attendance_date": iso_or_none(scan.attendance_date),
        "scanned_at": iso_or_none(scan.scanned_at),
    }


def serialize_print_job(job):
    return {
        "id": job.id,
        "member_id": job.member_id,
        "member_name": job.member.name if job.member else "",
        "printer_type": job.printer_type,
        "file_name": job.file_name,
        "status": job.status,
        "notes": job.notes,
        "submitted_at": iso_or_none(job.submitted_at),
        "started_at": iso_or_none(job.started_at),
        "completed_at": iso_or_none(job.completed_at),
        "open_url": url_for("legacy_ops.open_print_job", job_id=job.id),
        "download_url": url_for("legacy_ops.download_print_job", job_id=job.id),
    }


def serialize_print_request(row):
    return {
        "id": row.id,
        "user_id": row.user_id,
        "printer_type": row.printer_type,
        "file_link": row.file_link,
        "has_file": bool(row.file_path),
        "filament": row.filament,
        "material": row.material,
        "color": row.color,
        "infill_percent": row.infill_percent,
        "priority": row.priority,
        "deadline": iso_or_none(row.deadline),
        "notes": row.notes,
        "admin_notes": row.admin_notes,
        "status": row.status,
        "reviewed_at": iso_or_none(row.reviewed_at),
        "created_at": iso_or_none(row.created_at),
        "runs": [
            {"id": run.id, "status": run.status, "started_at": iso_or_none(run.started_at), "completed_at": iso_or_none(run.completed_at)}
            for run in row.runs
        ],
    }


def serialize_event(event, sync=None):
    return {
        "id": event.id,
        "title": event.title,
        "kind": event.kind,
        "location": event.location,
        "status": event.status,
        "start_time": iso_or_none(event.start_time),
        "end_time": iso_or_none(event.end_time),
        "calendar_link": event.calendar_event_link or (sync.link if sync else None),
        "sync_status": sync.status if sync else None,
    }


def serialize_attendance_record(row):
    return {
        "id": row.id,
        "event_id": row.event_id,
        "event_title": row.event.title if row.event else None,
        "user_id": row.user_id,
        "method": row.checkin_method,
        "checkin_time": iso_or_none(row.checkin_time),
    }
