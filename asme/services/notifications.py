"""Outbound notifications. Anything that talks to SMTP goes through here."""

from __future__ import annotations

from asme.config import settings
from asme.integrations import mail


def send_meeting_cancel_confirmation_email(meeting, confirm_url, reject_url) -> str | None:
    """Legacy synchronous send. Returns an error string or ``None``."""
    cfg = settings()
    notify_to = cfg.cancel_notify_to or cfg.smtp_user
    if not mail.is_configured(cfg):
        return "SMTP is not configured. Set ASME_SMTP_USER and ASME_SMTP_PASS."
    if not notify_to:
        return "ASME_CANCEL_NOTIFY_TO is not configured."
    body = "\n".join(
        [
            "A meeting cancellation was requested from the website.",
            "",
            f"Team: {meeting.team_name}",
            f"Room: {meeting.room}",
            f"Date: {meeting.meeting_date.isoformat()}",
            f"Time: {meeting.start_time.strftime('%H:%M')} - {meeting.end_time.strftime('%H:%M')}",
            f"Requested by: {meeting.requester_email or 'Not provided'}",
            f"Notes: {meeting.notes or ''}",
            "",
            f"Confirm cancellation: {confirm_url}",
            f"Reject cancellation:  {reject_url}",
        ]
    )
    try:
        mail.send_email(
            cfg,
            notify_to,
            f"ASME Meeting Cancellation Request: {meeting.team_name} ({meeting.room})",
            body,
        )
    except Exception as exc:
        return f"Failed to send cancellation email: {str(exc)[:250]}"
    return None


def queue_email(to, subject, body):
    """Durable send through the outbox (retried by the worker)."""
    from asme.jobs.outbox import enqueue

    return enqueue("mail.send", {"to": to, "subject": subject, "body": body})
