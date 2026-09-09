"""Setup Center: phase/task completion computed from persisted configuration.

Nothing here is a stored checkbox. Each task has a ``check`` that counts real rows;
tasks whose module is not built yet report ``available: False`` and are shown
without a call to action so the product never claims a feature it lacks.
"""

from __future__ import annotations

from sqlalchemy import func

from asme.extensions import db
from asme.models import User
from asme.ops.models import Asset, Category, Location, Membership, OpsProject, Team

# (key, title, description, minutes, required, route, check(ctx) -> (current, target), available)
PHASES = [
    {
        "key": "foundation",
        "title": "Build the Foundation",
        "description": "Tell ASME Ops who you are, where your spaces are, what you own and who is on the team.",
        "tasks": [
            {
                "key": "chapter_profile",
                "title": "Chapter profile",
                "description": "Name, timezone and academic year so dates and reports line up.",
                "minutes": 2,
                "required": True,
                "route": "/app/settings/chapter",
                "available": True,
            },
            {
                "key": "locations",
                "title": "Add locations",
                "description": "At least one location besides the default so assets and work have a home.",
                "minutes": 5,
                "required": True,
                "route": "/app/locations",
                "available": True,
            },
            {
                "key": "assets",
                "title": "Add assets",
                "description": "Register five or more assets: the rover, printers, stations, test equipment.",
                "minutes": 10,
                "required": True,
                "route": "/app/assets",
                "available": True,
            },
            {
                "key": "teams_users",
                "title": "Teams and users",
                "description": "One team and three members so work can be assigned.",
                "minutes": 5,
                "required": True,
                "route": "/app/teams-users",
                "available": True,
            },
            {
                "key": "officer_guide",
                "title": "Officer onboarding guide",
                "description": "Read how projects, work orders and requests fit together.",
                "minutes": 8,
                "required": False,
                "route": "/app/setup#guide",
                "available": True,
            },
        ],
    },
    {
        "key": "project_work",
        "title": "Organize Project Work",
        "description": "Turn the chapter's programmes into projects with categories and parts behind them.",
        "tasks": [
            {"key": "first_project", "title": "Create your first project", "description": "A programme with a lead, a target date and milestones.", "minutes": 5, "required": True, "route": "/app/projects", "available": True},
            {"key": "categories", "title": "Create categories", "description": "Mechanical, Electrical, Safety… used to route and report on work.", "minutes": 3, "required": True, "route": "/app/categories", "available": True},
            {"key": "parts", "title": "Add parts inventory", "description": "Fasteners, motors, filament - with minimum stock levels.", "minutes": 15, "required": True, "route": "/app/parts", "available": False},
            {"key": "procedure", "title": "Publish your first procedure", "description": "A pre-drive checklist or shop safety inspection.", "minutes": 10, "required": True, "route": "/app/library/procedures", "available": False},
        ],
    },
    {
        "key": "standardize",
        "title": "Standardize Operations",
        "description": "Let the system generate and route work for you.",
        "tasks": [
            {"key": "maintenance_plan", "title": "Create a maintenance plan", "description": "Monthly printer inspection, semester station check.", "minutes": 5, "required": True, "route": "/app/maintenance-plans", "available": False},
            {"key": "request_portal", "title": "Configure a request portal", "description": "A QR code members scan to report a broken tool.", "minutes": 5, "required": True, "route": "/app/requests", "available": False},
            {"key": "automation", "title": "Add an automation", "description": "Notify the safety officer when a critical work order is created.", "minutes": 5, "required": True, "route": "/app/automations", "available": False},
            {"key": "dashboard", "title": "Create a dashboard", "description": "Pin the report cards your exec board looks at every week.", "minutes": 5, "required": True, "route": "/app/reporting/dashboards", "available": False},
        ],
    },
]


def _check(ctx, key: str) -> tuple[int, int]:
    org_id = ctx.organization.id
    if key == "chapter_profile":
        org = ctx.organization
        settings = org.settings
        done = int(bool(org.name and org.timezone and settings.get("profile_confirmed")))
        return done, 1
    if key == "locations":
        n = Location.query.filter(Location.organization_id == org_id, Location.is_default.is_(False), Location.archived_at.is_(None)).count()
        return min(n, 1), 1
    if key == "assets":
        n = Asset.query.filter(Asset.organization_id == org_id, Asset.archived_at.is_(None)).count()
        return min(n, 5), 5
    if key == "teams_users":
        teams = Team.query.filter(Team.organization_id == org_id, Team.archived_at.is_(None)).count()
        users = db.session.query(func.count(Membership.id)).join(User, User.id == Membership.user_id).filter(Membership.organization_id == org_id, Membership.member_status == "active", User.is_active.is_(True)).scalar() or 0
        return min(teams, 1) + min(int(users), 3), 4
    if key == "officer_guide":
        return int(bool(ctx.membership.settings.get("officer_guide_read"))), 1
    if key == "first_project":
        n = OpsProject.query.filter(OpsProject.organization_id == org_id, OpsProject.archived_at.is_(None)).count()
        return min(n, 1), 1
    if key == "categories":
        n = Category.query.filter(Category.organization_id == org_id, Category.archived_at.is_(None)).count()
        return min(n, 1), 1
    return 0, 1


def evaluate(ctx) -> dict:
    phases = []
    required_total = required_done = 0
    next_step = None
    for phase in PHASES:
        tasks = []
        phase_required = phase_done = 0
        for task in phase["tasks"]:
            current, target = _check(ctx, task["key"]) if task["available"] else (0, 1)
            complete = current >= target
            if task["required"]:
                phase_required += 1
                phase_done += 1 if complete else 0
            payload = {**task, "current": current, "target": target, "complete": complete}
            tasks.append(payload)
            if next_step is None and task["required"] and not complete and task["available"]:
                next_step = {"phase": phase["key"], "task": task["key"], "title": task["title"], "route": task["route"]}
        required_total += phase_required
        required_done += phase_done
        phases.append(
            {
                "key": phase["key"],
                "title": phase["title"],
                "description": phase["description"],
                "tasks": tasks,
                "required_total": phase_required,
                "required_done": phase_done,
                "complete": phase_required > 0 and phase_done == phase_required,
                "estimated_minutes": sum(t["minutes"] for t in phase["tasks"] if t["required"]),
                "available_tasks": sum(1 for t in phase["tasks"] if t["available"]),
            }
        )
    percent = int(round(required_done / required_total * 100)) if required_total else 0
    return {
        "phases": phases,
        "percent": percent,
        "required_total": required_total,
        "required_done": required_done,
        "steps_left": required_total - required_done,
        "next_step": next_step,
        "banner_dismissed": bool(ctx.membership.settings.get("setup_banner_dismissed")),
        "complete": required_total > 0 and required_done == required_total,
    }


def mark_guide_read(ctx):
    settings = ctx.membership.settings
    settings["officer_guide_read"] = True
    ctx.membership.settings = settings
    db.session.commit()


def confirm_profile(ctx):
    settings = ctx.organization.settings
    settings["profile_confirmed"] = True
    ctx.organization.settings = settings
    db.session.commit()
