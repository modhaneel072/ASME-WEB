"""3D printing: portal print requests and the legacy shell-command job queue."""

from __future__ import annotations

import os
import subprocess
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from flask import current_app
from werkzeug.utils import secure_filename

from asme import events
from asme.config import settings
from asme.constants import ALLOWED_GCODE_EXTENSIONS, PRINT_REQUEST_PRINTERS, PRINT_REQUEST_STATUSES, PRINTER_TYPES
from asme.extensions import db
from asme.models import PrintJob, PrintRequest, PrintRun
from asme.services import audit
from asme.services.errors import Conflict, NotFound, Validation
from asme.utils import parse_due_date, parse_positive_int


def upload_dir() -> Path:
    path = Path(current_app.instance_path) / "gcode_uploads"
    path.mkdir(parents=True, exist_ok=True)
    return path


def allowed_gcode(filename):
    if "." not in (filename or ""):
        return False
    return filename.rsplit(".", 1)[1].lower() in ALLOWED_GCODE_EXTENSIONS


def normalize_portal_printer_type(raw_value):
    cleaned = (raw_value or "").strip().upper().replace("-", "_")
    return cleaned if cleaned in PRINT_REQUEST_PRINTERS else ""


def _store_upload(file_upload, prefix="printreq"):
    safe_name = secure_filename(file_upload.filename)
    if not safe_name or not allowed_gcode(safe_name):
        raise Validation("Upload .gcode, .gco, or .3mf files only.", field="print_file")
    try:
        file_upload.stream.seek(0, os.SEEK_END)
        file_size = file_upload.stream.tell()
        file_upload.stream.seek(0)
    except Exception:
        file_size = 0
    max_bytes = settings().print_max_upload_bytes
    if file_size > max_bytes:
        raise Validation(f"File is too large. Max upload size is {max(1, max_bytes // (1024 * 1024))} MB.", field="print_file")
    stored_name = f"{prefix}_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{safe_name}"
    output_path = upload_dir() / stored_name
    file_upload.save(output_path)
    return str(output_path), safe_name


# --------------------------------------------------------------------------- portal requests


def submit_print_request(user, form, file_upload=None, member=None) -> PrintRequest:
    printer_type = normalize_portal_printer_type(form.get("printer_type"))
    filament = (form.get("filament") or "").strip()
    if printer_type not in PRINT_REQUEST_PRINTERS:
        raise Validation("Choose a valid printer queue (P1S-1..P1S-4 or H2S).", field="printer_type")
    if not filament:
        raise Validation("Filament type/color is required.", field="filament")

    file_link = (form.get("file_link") or "").strip() or None
    file_path = None
    if file_upload and file_upload.filename:
        file_path, _name = _store_upload(file_upload)
    if not file_path and not file_link:
        raise Validation("Upload a file or provide a link.", field="print_file")

    row = PrintRequest(
        user_id=user.id,
        member_id=member.id if member else None,
        printer_type=printer_type,
        file_path=file_path,
        file_link=file_link,
        filament=filament[:160],
        material=(form.get("material") or "").strip() or None,
        color=(form.get("color") or "").strip() or None,
        infill_percent=min(parse_positive_int(form.get("infill_percent"), default=20), 100),
        priority=((form.get("priority") or "normal").strip().lower() or "normal")[:40],
        deadline=parse_due_date(form.get("deadline")),
        notes=(form.get("notes") or "").strip() or None,
        status="submitted",
    )
    db.session.add(row)
    db.session.commit()
    events.emit(events.PRINT_SUBMITTED, request_id=row.id, user_id=user.id)
    return row


def update_print_request(row, status, printer_type, admin_notes, actor) -> PrintRequest:
    status = (status or "").strip().lower()
    if status not in PRINT_REQUEST_STATUSES:
        raise Validation("Invalid status value.", field="status")
    printer_type = normalize_portal_printer_type(printer_type) or row.printer_type
    if printer_type not in PRINT_REQUEST_PRINTERS:
        printer_type = row.printer_type
    previous = row.status
    row.status = status
    row.printer_type = printer_type
    row.admin_notes = (admin_notes or "").strip() or None
    row.reviewed_by_user_id = actor.id if actor else None
    row.reviewed_at = datetime.utcnow()

    # Keep an execution record per attempt so a failed print can be re-run.
    if status == "printing" and previous != "printing":
        db.session.add(
            PrintRun(request_id=row.id, printer_type=printer_type, status="printing", started_by_user_id=actor.id if actor else None, started_at=datetime.utcnow())
        )
    elif status in {"completed", "rejected"}:
        run = PrintRun.query.filter_by(request_id=row.id).order_by(PrintRun.id.desc()).first()
        if run and run.status == "printing":
            run.status = "done" if status == "completed" else "failed"
            run.completed_at = datetime.utcnow()
    audit.record("update_print_request", f"request_id={row.id} status={status}", actor=actor)
    db.session.commit()
    events.emit(events.PRINT_UPDATED, request_id=row.id, user_id=row.user_id, status=status)
    return row


def pending_count():
    return PrintRequest.query.filter(PrintRequest.status.in_(["submitted", "approved", "printing"])).count()


# --------------------------------------------------------------------------- legacy job queue


def launch_print_command(job):
    cfg = settings()
    cmd_template = cfg.h2s_print_cmd if job.printer_type == "H2S" else cfg.p1s_print_cmd
    env_name = f"ASME_{job.printer_type}_PRINT_CMD"
    if not cmd_template:
        return f"{env_name} is not configured. Add it to instance/print_commands.env and restart the app."
    command = cmd_template.format(file=job.file_path, filename=job.file_name, job_id=job.id)
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    if result.returncode == 0:
        return None
    error_text = (result.stderr or result.stdout or "print command failed").strip()
    return f"{env_name} failed: {error_text[:300]}"


def dispatch_next_job(printer_type):
    if PrintJob.query.filter_by(printer_type=printer_type, status="printing").first():
        return None
    next_job = (
        PrintJob.query.filter_by(printer_type=printer_type, status="queued")
        .order_by(PrintJob.submitted_at.asc(), PrintJob.id.asc())
        .first()
    )
    if not next_job:
        return None
    next_job.status = "printing"
    next_job.started_at = datetime.utcnow()
    db.session.commit()

    dispatch_error = launch_print_command(next_job)
    if dispatch_error:
        next_job.status = "failed"
        next_job.completed_at = datetime.utcnow()
        next_job.notes = f"{next_job.notes} | {dispatch_error}" if next_job.notes else dispatch_error
        db.session.commit()
        return dispatch_next_job(printer_type)
    return next_job


def submit_print_job(member, printer_type, file_upload, notes=None) -> tuple[PrintJob, bool]:
    printer_type = (printer_type or "").strip().upper()
    if not member:
        raise Validation("Could not find member for print job. Scan/select member first.")
    if printer_type not in PRINTER_TYPES:
        raise Validation("Invalid printer queue selected.", field="printer_type")
    if not file_upload or not file_upload.filename:
        raise Validation("Please upload a print file.", field="gcode_file")
    if not allowed_gcode(file_upload.filename):
        raise Validation("Invalid file type. Use .gcode, .gco, or .3mf.", field="gcode_file")

    original_name = secure_filename(file_upload.filename)
    stored_name = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}_{original_name}"
    file_path = upload_dir() / stored_name
    file_upload.save(file_path)

    job = PrintJob(
        member_id=member.id,
        printer_type=printer_type,
        file_name=original_name,
        file_path=str(file_path),
        notes=notes,
        status="queued",
    )
    db.session.add(job)
    db.session.commit()
    started = dispatch_next_job(printer_type)
    return job, bool(started and started.id == job.id)


def finish_print_job(job, status) -> PrintJob:
    if status not in {"done", "failed"}:
        raise Validation("Invalid job status.")
    job.status = status
    job.completed_at = datetime.utcnow()
    db.session.commit()
    dispatch_next_job(job.printer_type)
    return job


def delete_print_job_with_file(job):
    if job.status == "printing":
        raise Conflict("Cannot delete an active printing job. Mark it done or failed first.", code="job_active")
    file_path, file_name, printer_type = job.file_path, job.file_name, job.printer_type
    file_removed, file_error = False, None
    if file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
            file_removed = True
        except Exception as exc:
            file_error = str(exc)
    db.session.delete(job)
    db.session.commit()
    dispatch_next_job(printer_type)
    return {"file_name": file_name, "file_removed": file_removed, "file_error": file_error}


def queue_snapshot():
    payload = {}
    for printer in PRINTER_TYPES:
        active = (
            PrintJob.query.filter_by(printer_type=printer, status="printing")
            .order_by(PrintJob.started_at.asc(), PrintJob.id.asc())
            .first()
        )
        queued = (
            PrintJob.query.filter_by(printer_type=printer, status="queued")
            .order_by(PrintJob.submitted_at.asc(), PrintJob.id.asc())
            .all()
        )
        finished = (
            PrintJob.query.filter(PrintJob.printer_type == printer, PrintJob.status.in_(["done", "failed"]))
            .order_by(PrintJob.completed_at.desc(), PrintJob.id.desc())
            .limit(8)
            .all()
        )
        payload[printer] = {"active": active, "queued": queued, "recent_finished": finished}
    return payload


def get_job_or_404(job_id) -> PrintJob:
    job = db.session.get(PrintJob, job_id)
    if not job:
        raise NotFound("Print job not found.")
    return job
