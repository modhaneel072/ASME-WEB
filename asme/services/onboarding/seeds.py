"""Default Launchpad tracks. Idempotent: re-running updates names/rules by key
and never touches existing progress rows."""

from __future__ import annotations

from asme.constants import (
    ENT_CHECKOUT_SIGNOFF,
    ENT_EVENT_CHECKIN,
    ENT_EXTENDED_LOAN,
    ENT_PORTAL_ACCESS,
    ENT_PRINT_APPROVE,
    ENT_PRINT_SUBMIT,
    ENT_ROOM_BOOKING,
    ENT_SHOP_ACCESS,
    SUBJECT_CHAPTER,
    SUBJECT_USER,
)
from asme.extensions import db
from asme.models import Phase, Task, Track, TrainingModule

TRAINING_MODULES = [
    {
        "key": "shop_safety",
        "title": "Shop safety",
        "provider": "ASME @ UIowa",
        "duration_min": 75,
        "is_optional": False,
        "min_score": 80,
        "validity_days": 365,
        "description": "Required before any tool leaves the crib. Covers PPE, machine lockout, and incident reporting.",
    },
    {
        "key": "machine_basics",
        "title": "Machine-specific training: lathe, mill, welding",
        "provider": "Shop manager",
        "duration_min": 60,
        "is_optional": True,
        "min_score": None,
        "validity_days": 730,
        "description": "Hands-on sign-off per machine. Optional for shop access, required to run the machine unsupervised.",
    },
    {
        "key": "lead_essentials",
        "title": "Lead essentials: approvals, budget, incident reporting",
        "provider": "Exec board",
        "duration_min": 45,
        "is_optional": False,
        "min_score": None,
        "validity_days": None,
        "description": "What a team lead is responsible for: print approvals, checkout sign-off, budget requests.",
    },
]

MEMBER_TRACK = {
    "key": "member_onboarding",
    "name": "Member Launchpad",
    "audience": SUBJECT_USER,
    "description": "From signup to shop-trusted. Finish a phase to unlock what it grants.",
    "phases": [
        {
            "key": "signed_up",
            "name": "Signed up",
            "tagline": "Get on the roster",
            "description": "Tell us who you are and where to find you.",
            "order": 1,
            "est_minutes": 10,
            "grants": [ENT_PORTAL_ACCESS, ENT_EVENT_CHECKIN],
            "tasks": [
                {
                    "key": "verify_email",
                    "name": "Use your university email",
                    "description": "Your account email must be an @uiowa.edu address.",
                    "is_required": True,
                    "est_minutes": 2,
                    "rule_type": "email_domain",
                    "rule_config": {"domains": ["uiowa.edu"]},
                    "cta_route": "/portal/member/profile",
                    "cta_label": "Update email",
                    "order": 1,
                },
                {
                    "key": "complete_profile",
                    "name": "Complete your profile",
                    "description": "Major and graduation year, so leads know who is on the floor.",
                    "is_required": True,
                    "est_minutes": 3,
                    "rule_type": "existence",
                    "rule_config": {"fields": ["major", "graduation_year"]},
                    "cta_route": "/portal/member/profile",
                    "cta_label": "Edit profile",
                    "order": 2,
                },
                {
                    "key": "link_nfc",
                    "name": "Link your NFC card",
                    "description": "Ask an admin at the kiosk to assign your card. It is how you check in and check tools out.",
                    "is_required": True,
                    "est_minutes": 2,
                    "rule_type": "existence",
                    "rule_config": {"fields": ["nfc"]},
                    "cta_route": "/portal/member/profile",
                    "cta_label": "See card status",
                    "order": 3,
                },
                {
                    "key": "join_team",
                    "name": "Join a project team",
                    "description": "Pick at least one team. You can change later.",
                    "is_required": True,
                    "est_minutes": 3,
                    "rule_type": "existence",
                    "rule_config": {"fields": ["team"]},
                    "cta_route": "/portal/member/launchpad#teams",
                    "cta_label": "Choose a team",
                    "order": 4,
                },
            ],
        },
        {
            "key": "shop_ready",
            "name": "Shop ready",
            "tagline": "Earn shop access",
            "description": "Safety first. Finishing this phase is what lets you take a tool off the wall.",
            "order": 2,
            "est_minutes": 95,
            "grants": [ENT_SHOP_ACCESS, ENT_PRINT_SUBMIT],
            "tasks": [
                {
                    "key": "shop_safety",
                    "name": "Shop safety module",
                    "description": "75 minutes, quiz score of 80% or better. Recorded by a lead.",
                    "is_required": True,
                    "est_minutes": 75,
                    "rule_type": "training",
                    "rule_config": {"module": "shop_safety", "min_score": 80},
                    "cta_route": "/portal/member/launchpad#training",
                    "cta_label": "Training details",
                    "order": 1,
                },
                {
                    "key": "first_meeting",
                    "name": "Attend a general meeting",
                    "description": "Tap in at any general meeting with your card.",
                    "is_required": True,
                    "est_minutes": 60,
                    "rule_type": "count_threshold",
                    "rule_config": {"model": "AttendanceRecord", "min": 1},
                    "cta_route": "/portal/member/calendar",
                    "cta_label": "Upcoming meetings",
                    "order": 2,
                },
                {
                    "key": "supervised_checkout",
                    "name": "Supervised first checkout",
                    "description": "Check a tool out with a team lead present; the lead signs it off.",
                    "is_required": True,
                    "est_minutes": 10,
                    "rule_type": "loan_signed_off",
                    "rule_config": {"min": 1},
                    "cta_route": "/portal/member/inventory",
                    "cta_label": "Browse tools",
                    "order": 3,
                },
                {
                    "key": "machine_training",
                    "name": "Machine-specific training",
                    "description": "Lathe, mill, welding - optional now, required before running one unsupervised.",
                    "is_required": False,
                    "est_minutes": 60,
                    "rule_type": "training",
                    "rule_config": {"module": "machine_basics"},
                    "cta_route": "/portal/member/launchpad#training",
                    "cta_label": "Training details",
                    "order": 4,
                },
            ],
        },
        {
            "key": "contributor",
            "name": "Contributor",
            "tagline": "Show up and ship",
            "description": "Consistent attendance and clean returns earn longer loans and room booking.",
            "order": 3,
            "est_minutes": None,
            "grants": [ENT_EXTENDED_LOAN, ENT_ROOM_BOOKING],
            "tasks": [
                {
                    "key": "three_meetings",
                    "name": "Attend 3 meetings this semester",
                    "description": None,
                    "is_required": True,
                    "est_minutes": None,
                    "rule_type": "count_threshold",
                    "rule_config": {"model": "AttendanceRecord", "min": 3, "window": "semester"},
                    "cta_route": "/portal/member/calendar",
                    "cta_label": "Upcoming meetings",
                    "order": 1,
                },
                {
                    "key": "clean_returns",
                    "name": "Return 3 checkouts on time, nothing overdue",
                    "description": "An overdue tool pauses this phase until it comes back.",
                    "is_required": True,
                    "est_minutes": None,
                    "rule_type": "clean_streak",
                    "rule_config": {"min": 3},
                    "cta_route": "/portal/member/my-items",
                    "cta_label": "My items",
                    "order": 2,
                },
                {
                    "key": "ten_hours",
                    "name": "Log 10 hours on a project",
                    "description": None,
                    "is_required": True,
                    "est_minutes": None,
                    "rule_type": "hours_threshold",
                    "rule_config": {"min": 10, "window": "semester"},
                    "cta_route": "/portal/member/launchpad#hours",
                    "cta_label": "Log hours",
                    "order": 3,
                },
                {
                    "key": "first_print",
                    "name": "Submit and collect a print",
                    "description": None,
                    "is_required": False,
                    "est_minutes": None,
                    "rule_type": "count_threshold",
                    "rule_config": {"model": "PrintRequest", "min": 1, "where": {"status": "completed"}},
                    "cta_route": "/portal/member/prints",
                    "cta_label": "3D printing",
                    "order": 4,
                },
            ],
        },
        {
            "key": "team_lead",
            "name": "Team lead",
            "tagline": "Nominated",
            "description": "Leads approve prints, sign off checkouts and run subteams.",
            "order": 4,
            "est_minutes": None,
            "grants": ["role:team_leader", ENT_PRINT_APPROVE, ENT_CHECKOUT_SIGNOFF],
            "tasks": [
                {
                    "key": "nominated",
                    "name": "Nominated by an exec member",
                    "description": "An admin records the nomination.",
                    "is_required": True,
                    "est_minutes": None,
                    "rule_type": "signoff",
                    "rule_config": {"by_role": "admin"},
                    "cta_route": None,
                    "cta_label": None,
                    "order": 1,
                },
                {
                    "key": "lead_training",
                    "name": "Lead essentials training",
                    "description": "Approvals, budget requests, incident reporting.",
                    "is_required": True,
                    "est_minutes": 45,
                    "rule_type": "training",
                    "rule_config": {"module": "lead_essentials"},
                    "cta_route": "/portal/member/launchpad#training",
                    "cta_label": "Training details",
                    "order": 2,
                },
                {
                    "key": "build_cycle",
                    "name": "Complete a full build cycle as a contributor",
                    "description": "Signed off by your current team lead.",
                    "is_required": True,
                    "est_minutes": None,
                    "rule_type": "signoff",
                    "rule_config": {"by_role": "team_leader"},
                    "cta_route": None,
                    "cta_label": None,
                    "order": 3,
                },
            ],
        },
    ],
}

CHAPTER_TRACK = {
    "key": "chapter_setup",
    "name": "Chapter setup",
    "audience": SUBJECT_CHAPTER,
    "description": "Runs once per academic year. The August checklist, kept where it cannot graduate.",
    "phases": [
        {
            "key": "stand_up",
            "name": "Stand up the chapter",
            "tagline": "Roster, teams, crib, rooms",
            "description": "The MaintainX trio - locations, assets, team members - in ASME terms.",
            "order": 1,
            "est_minutes": 120,
            "grants": [],
            "tasks": [
                {
                    "key": "import_roster",
                    "name": "Import the roster (20+ accounts)",
                    "description": "Admin -> Members -> Bulk Import Roster.",
                    "is_required": True,
                    "est_minutes": 15,
                    "rule_type": "chapter_count",
                    "rule_config": {"model": "User", "min": 20, "where": {"role_in": ["member", "team_leader"]}},
                    "cta_route": "/portal/admin/members",
                    "cta_label": "Members",
                    "order": 1,
                },
                {
                    "key": "define_teams",
                    "name": "Define 3+ project teams",
                    "description": "Joinable projects are what members pick from in their first phase.",
                    "is_required": True,
                    "est_minutes": 15,
                    "rule_type": "chapter_count",
                    "rule_config": {"model": "Project", "min": 3, "where": {"joinable": True}},
                    "cta_route": "/portal/admin",
                    "cta_label": "Projects",
                    "order": 2,
                },
                {
                    "key": "seed_inventory",
                    "name": "Seed the tool crib (40+ items)",
                    "description": "Bulk import from the treasury spreadsheet.",
                    "is_required": True,
                    "est_minutes": 30,
                    "rule_type": "chapter_count",
                    "rule_config": {"model": "Item", "min": 40},
                    "cta_route": "/portal/admin/inventory",
                    "cta_label": "Inventory",
                    "order": 3,
                },
                {
                    "key": "shop_locations",
                    "name": "Register 5+ shop locations",
                    "description": "Every item has a bin, drawer or shelf.",
                    "is_required": True,
                    "est_minutes": 10,
                    "rule_type": "chapter_count",
                    "rule_config": {"model": "ItemLocation", "min": 5},
                    "cta_route": "/portal/admin/inventory",
                    "cta_label": "Inventory",
                    "order": 4,
                },
                {
                    "key": "connect_calendar",
                    "name": "Connect the room calendar",
                    "description": "Google service account or Outlook app registration.",
                    "is_required": True,
                    "est_minutes": 20,
                    "rule_type": "config_set",
                    "rule_config": {"setting": "calendar"},
                    "cta_route": "/portal/admin/calendar",
                    "cta_label": "Calendar status",
                    "order": 5,
                },
                {
                    "key": "provision_kiosk_tag",
                    "name": "Provision the first NFC card",
                    "description": "Assign at least one card so the kiosk flow is proven.",
                    "is_required": True,
                    "est_minutes": 5,
                    "rule_type": "chapter_count",
                    "rule_config": {"model": "NFCTag", "min": 1},
                    "cta_route": "/portal/admin/members",
                    "cta_label": "NFC tags",
                    "order": 6,
                },
            ],
        },
        {
            "key": "run_semester",
            "name": "Run the semester",
            "tagline": "Schedule, queue, coverage",
            "description": "Operate. The last task is the one that matters: most of the roster safety-trained.",
            "order": 2,
            "est_minutes": 60,
            "grants": [],
            "tasks": [
                {
                    "key": "publish_schedule",
                    "name": "Publish 3+ upcoming meetings",
                    "description": None,
                    "is_required": True,
                    "est_minutes": 15,
                    "rule_type": "chapter_count",
                    "rule_config": {"model": "Event", "min": 3, "where": {"future": True}},
                    "cta_route": "/portal/admin/attendance",
                    "cta_label": "Events",
                    "order": 1,
                },
                {
                    "key": "open_print_queue",
                    "name": "Open the print queue",
                    "description": "Confirm printers are staffed and filament is stocked.",
                    "is_required": True,
                    "est_minutes": 10,
                    "rule_type": "manual",
                    "rule_config": {"note_required": True},
                    "cta_route": "/portal/admin/prints",
                    "cta_label": "Prints",
                    "order": 2,
                },
                {
                    "key": "set_loan_periods",
                    "name": "Set loan periods",
                    "description": "Default 7 days; contributors get 14.",
                    "is_required": True,
                    "est_minutes": 5,
                    "rule_type": "manual",
                    "rule_config": {},
                    "cta_route": "/portal/admin/inventory",
                    "cta_label": "Inventory",
                    "order": 3,
                },
                {
                    "key": "safety_coverage",
                    "name": "60% of members shop-ready",
                    "description": "Fraction of active members who have completed the Shop ready phase.",
                    "is_required": True,
                    "est_minutes": None,
                    "rule_type": "member_phase_ratio",
                    "rule_config": {"phase": "shop_ready", "min_ratio": 0.6},
                    "cta_route": "/portal/admin/launchpad",
                    "cta_label": "Member progress",
                    "order": 4,
                },
            ],
        },
    ],
}


def _upsert_track(spec) -> Track:
    track = Track.query.filter_by(key=spec["key"]).first()
    if track is None:
        track = Track(key=spec["key"], name=spec["name"], audience=spec["audience"])
        db.session.add(track)
        db.session.flush()
    track.name = spec["name"]
    track.audience = spec["audience"]
    track.description = spec.get("description")
    track.is_active = True

    seen_phase_ids = []
    for phase_spec in spec["phases"]:
        phase = Phase.query.filter_by(track_id=track.id, key=phase_spec["key"]).first()
        if phase is None:
            phase = Phase(track_id=track.id, key=phase_spec["key"], name=phase_spec["name"])
            db.session.add(phase)
            db.session.flush()
        phase.name = phase_spec["name"]
        phase.tagline = phase_spec.get("tagline")
        phase.description = phase_spec.get("description")
        phase.order = phase_spec["order"]
        phase.est_minutes = phase_spec.get("est_minutes")
        phase.grants = phase_spec.get("grants") or []
        seen_phase_ids.append(phase.id)

        seen_task_ids = []
        for task_spec in phase_spec["tasks"]:
            task = Task.query.filter_by(phase_id=phase.id, key=task_spec["key"]).first()
            if task is None:
                task = Task(phase_id=phase.id, key=task_spec["key"], name=task_spec["name"])
                db.session.add(task)
                db.session.flush()
            task.name = task_spec["name"]
            task.description = task_spec.get("description")
            task.is_required = bool(task_spec.get("is_required", True))
            task.est_minutes = task_spec.get("est_minutes")
            task.rule_type = task_spec["rule_type"]
            task.rule_config = task_spec.get("rule_config") or {}
            task.cta_route = task_spec.get("cta_route")
            task.cta_label = task_spec.get("cta_label")
            task.help_url = task_spec.get("help_url")
            task.order = task_spec["order"]
            seen_task_ids.append(task.id)
    return track


def seed_training_modules():
    for spec in TRAINING_MODULES:
        module = TrainingModule.query.filter_by(key=spec["key"]).first()
        if module is None:
            module = TrainingModule(key=spec["key"], title=spec["title"])
            db.session.add(module)
        for field_name, value in spec.items():
            setattr(module, field_name, value)


def seed_default_tracks(commit=True):
    seed_training_modules()
    _upsert_track(MEMBER_TRACK)
    _upsert_track(CHAPTER_TRACK)
    if commit:
        db.session.commit()
