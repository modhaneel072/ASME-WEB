"""Transactional outbox + in-process worker.

``enqueue`` adds an ``OutboxJob`` row to the *current* session without
committing, so it lands in the same transaction as the domain write that needs
it. The worker thread (or ``process_pending`` in tests) picks jobs up, retries
with exponential backoff, and marks them failed after ``max_attempts``.

At ~200 members on free-tier hosting a table and a thread is the proportionate
answer; the handler registry is the seam for moving to RQ later.
"""

from __future__ import annotations

import json
import logging
import threading
import time
from datetime import datetime, timedelta
from typing import Callable

from asme.extensions import db
from asme.models import OutboxJob

log = logging.getLogger("asme.jobs")

_handlers: dict[str, Callable[[dict], None]] = {}
_worker_threads: dict[int, threading.Thread] = {}
_stop_events: dict[int, threading.Event] = {}


def handler(kind: str):
    def decorator(fn):
        _handlers[kind] = fn
        return fn

    return decorator


def registered_kinds():
    return sorted(_handlers)


def enqueue(kind: str, payload: dict | None = None, run_at: datetime | None = None, max_attempts: int = 5) -> OutboxJob:
    job = OutboxJob(
        kind=kind,
        payload_json=json.dumps(payload or {}, default=str),
        status="pending",
        run_at=run_at or datetime.utcnow(),
        max_attempts=max_attempts,
    )
    db.session.add(job)
    return job


def _claim(job: OutboxJob) -> bool:
    updated = (
        OutboxJob.query.filter(OutboxJob.id == job.id, OutboxJob.status == "pending")
        .update({OutboxJob.status: "running", OutboxJob.locked_at: datetime.utcnow()}, synchronize_session=False)
    )
    db.session.commit()
    return updated == 1


def _run_one(job: OutboxJob) -> bool:
    fn = _handlers.get(job.kind)
    payload = json.loads(job.payload_json or "{}")
    try:
        if fn is None:
            raise RuntimeError(f"no handler registered for job kind '{job.kind}'")
        fn(payload)
        db.session.refresh(job)
        job.status = "done"
        job.completed_at = datetime.utcnow()
        job.last_error = None
        db.session.commit()
        return True
    except Exception as exc:
        db.session.rollback()
        job = db.session.get(OutboxJob, job.id)
        job.attempts = (job.attempts or 0) + 1
        job.last_error = str(exc)[:1000]
        job.locked_at = None
        if job.attempts >= (job.max_attempts or 1):
            job.status = "failed"
            log.error("outbox job failed permanently id=%s kind=%s error=%s", job.id, job.kind, exc)
        else:
            job.status = "pending"
            job.run_at = datetime.utcnow() + timedelta(seconds=min(2 ** job.attempts * 15, 3600))
            log.warning("outbox job retry id=%s kind=%s attempt=%s error=%s", job.id, job.kind, job.attempts, exc)
        db.session.commit()
        return False


def process_pending(limit: int = 20) -> int:
    """Run due jobs once. Returns how many succeeded. Safe to call from anywhere
    inside an app context (tests call it directly)."""
    now = datetime.utcnow()
    # Release jobs stuck in "running" for more than 10 minutes (crashed worker).
    stale = now - timedelta(minutes=10)
    OutboxJob.query.filter(OutboxJob.status == "running", OutboxJob.locked_at < stale).update(
        {OutboxJob.status: "pending", OutboxJob.locked_at: None}, synchronize_session=False
    )
    db.session.commit()

    jobs = (
        OutboxJob.query.filter(OutboxJob.status == "pending", OutboxJob.run_at <= now)
        .order_by(OutboxJob.run_at.asc(), OutboxJob.id.asc())
        .limit(limit)
        .all()
    )
    succeeded = 0
    for job in jobs:
        if not _claim(job):
            continue
        if _run_one(job):
            succeeded += 1
    return succeeded


def ensure_recurring(kind: str, every: timedelta, payload: dict | None = None) -> bool:
    """Make sure a recurring job of ``kind`` exists in the next ``every`` window."""
    horizon = datetime.utcnow() - every
    recent = (
        OutboxJob.query.filter(OutboxJob.kind == kind, OutboxJob.status.in_(["pending", "running", "done"]))
        .filter((OutboxJob.completed_at >= horizon) | (OutboxJob.status.in_(["pending", "running"])))
        .first()
    )
    if recent:
        return False
    enqueue(kind, payload or {}, run_at=datetime.utcnow(), max_attempts=3)
    db.session.commit()
    return True


def stats():
    rows = db.session.query(OutboxJob.status, db.func.count(OutboxJob.id)).group_by(OutboxJob.status).all()
    return {status: int(count) for status, count in rows}


def recent_failures(limit=25):
    return OutboxJob.query.filter(OutboxJob.status == "failed").order_by(OutboxJob.id.desc()).limit(limit).all()


def retry(job: OutboxJob):
    job.status = "pending"
    job.attempts = 0
    job.run_at = datetime.utcnow()
    job.last_error = None
    db.session.commit()


# --------------------------------------------------------------------------- worker thread


def _loop(app, stop_event: threading.Event):
    cfg = app.config["SETTINGS"]
    with app.app_context():
        while not stop_event.is_set():
            try:
                ensure_recurring("stock.reconcile", timedelta(hours=24))
                process_pending()
            except Exception:  # pragma: no cover - defensive
                log.exception("outbox worker iteration failed")
                try:
                    db.session.rollback()
                except Exception:
                    pass
            finally:
                db.session.remove()
            stop_event.wait(cfg.outbox_poll_seconds)


def start_worker(app) -> threading.Thread | None:
    key = id(app)
    if key in _worker_threads and _worker_threads[key].is_alive():
        return _worker_threads[key]
    stop_event = threading.Event()
    thread = threading.Thread(target=_loop, args=(app, stop_event), name="asme-outbox", daemon=True)
    thread.start()
    _worker_threads[key] = thread
    _stop_events[key] = stop_event
    return thread


def stop_worker(app):
    key = id(app)
    event = _stop_events.pop(key, None)
    if event:
        event.set()
    thread = _worker_threads.pop(key, None)
    if thread:
        thread.join(timeout=2)
