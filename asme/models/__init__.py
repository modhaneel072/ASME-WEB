"""All ORM models, split by domain. Table names are unchanged from the original
single ``models.py`` so existing databases keep working."""

from asme.extensions import db
from asme.models.audit import AuditLog, OutboxJob
from asme.models.content import Announcement, ContactMessage, Project, ProjectMembership, WorkLog
from asme.models.events import AttendanceRecord, AttendanceScan, CalendarSync, Event, Meeting
from asme.models.fabrication import PrintJob, PrintRequest, PrintRun
from asme.models.identity import Member, NFCTag, PasswordResetToken, User
from asme.models.inventory import Item, ItemTag, Loan, StockDiscrepancy, StockLedger, Transaction
from asme.models.onboarding import (
    Entitlement,
    Phase,
    PhaseState,
    Task,
    TaskState,
    Track,
    TrainingCompletion,
    TrainingModule,
)

__all__ = [
    "db",
    "Announcement",
    "AttendanceRecord",
    "AttendanceScan",
    "AuditLog",
    "CalendarSync",
    "ContactMessage",
    "Entitlement",
    "Event",
    "Item",
    "ItemTag",
    "Loan",
    "Meeting",
    "Member",
    "NFCTag",
    "OutboxJob",
    "PasswordResetToken",
    "Phase",
    "PhaseState",
    "PrintJob",
    "PrintRequest",
    "PrintRun",
    "Project",
    "ProjectMembership",
    "StockDiscrepancy",
    "StockLedger",
    "Task",
    "TaskState",
    "Track",
    "TrainingCompletion",
    "TrainingModule",
    "Transaction",
    "User",
    "WorkLog",
]
