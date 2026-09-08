"""SMTP adapter. Callers never build ``EmailMessage`` objects themselves."""

from __future__ import annotations

import smtplib
from email.message import EmailMessage

from asme.config import Settings


class MailNotConfigured(RuntimeError):
    pass


def is_configured(cfg: Settings) -> bool:
    return bool(cfg.smtp_user and cfg.smtp_pass)


def send_email(cfg: Settings, to: str, subject: str, body: str) -> None:
    """Send a plain-text email. Raises ``MailNotConfigured`` or ``smtplib`` errors."""
    if not is_configured(cfg):
        raise MailNotConfigured("SMTP is not configured. Set ASME_SMTP_USER and ASME_SMTP_PASS.")
    recipient = (to or "").strip()
    if not recipient:
        raise ValueError("Recipient is required.")

    message = EmailMessage()
    message["From"] = cfg.smtp_user
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)

    with smtplib.SMTP(cfg.smtp_host, cfg.smtp_port, timeout=30) as smtp:
        smtp.starttls()
        smtp.login(cfg.smtp_user, cfg.smtp_pass)
        smtp.send_message(message)
