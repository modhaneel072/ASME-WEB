"""Domain constants shared across services and blueprints."""

import re

ROLE_ORDER = {"guest": 0, "member": 1, "team_leader": 2, "admin": 3}
ASSIGNABLE_ROLES = ("member", "team_leader", "admin")

PRINTER_TYPES = ("H2S", "P1S")
PRINT_REQUEST_PRINTERS = ("P1S_1", "P1S_2", "P1S_3", "P1S_4", "H2S")
PRINT_REQUEST_STATUSES = ("submitted", "approved", "printing", "completed", "rejected")
PRINT_RUN_STATUSES = ("queued", "printing", "done", "failed")

MEETING_ROOMS = ("Robotics Room", "Fluids Lab")
EVENT_STATUSES = ("requested", "scheduled", "cancelled")
EVENT_KINDS = ("meeting", "workshop", "build_session", "social", "open_shop")

ALLOWED_GCODE_EXTENSIONS = {"gcode", "gco", "3mf"}
ALLOWED_RETURN_PHOTO_EXTENSIONS = {"jpg", "jpeg", "png", "webp"}
BULK_INVENTORY_ALLOWED_EXTENSIONS = {"csv", "xlsx", "xls"}

TAG_PREFIXES = ("item_id:", "item:")
ITEM_TYPES = ("tool", "consumable")
INVENTORY_ITEM_CODE_PREFIX = "ASME-INV-"
TREASURY_TRACKER_PREFIX = "ASME-TRK-"

LOAN_STATE_OUT = "OUT"
LOAN_STATE_RETURNED = "RETURNED"

CONTACT_KINDS = ("contact", "join", "sponsor", "help")
CONTACT_STATUSES = ("new", "in_progress", "resolved")

EMAIL_RE = re.compile(r"^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$", re.IGNORECASE)
ROSTER_BOOL_WORDS = {"yes", "no"}

KNOWN_NON_MS_MAIL_DOMAINS = {
    "gmail.com",
    "yahoo.com",
    "icloud.com",
    "aol.com",
    "proton.me",
    "protonmail.com",
    "gmx.com",
    "mail.com",
}
POSSIBLY_CONSUMER_MS_DOMAINS = {"outlook.com", "hotmail.com", "live.com", "msn.com"}

# Entitlement keys granted by the onboarding engine (see services/onboarding/seeds.py).
ENT_PORTAL_ACCESS = "portal_access"
ENT_EVENT_CHECKIN = "event_checkin"
ENT_SHOP_ACCESS = "shop_access"
ENT_PRINT_SUBMIT = "print_submit"
ENT_ROOM_BOOKING = "room_booking"
ENT_EXTENDED_LOAN = "extended_loan_14d"
ENT_PRINT_APPROVE = "print_approve"
ENT_CHECKOUT_SIGNOFF = "checkout_signoff"
ALL_ENTITLEMENTS = (
    ENT_PORTAL_ACCESS,
    ENT_EVENT_CHECKIN,
    ENT_SHOP_ACCESS,
    ENT_PRINT_SUBMIT,
    ENT_ROOM_BOOKING,
    ENT_EXTENDED_LOAN,
    ENT_PRINT_APPROVE,
    ENT_CHECKOUT_SIGNOFF,
)

SUBJECT_USER = "user"
SUBJECT_CHAPTER = "chapter"
