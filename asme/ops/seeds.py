"""Ops seed data.

* ``seed_ops_defaults`` - idempotent, safe for production: organisation, permissions,
  system roles, memberships for existing users, default location, categories, asset
  types, ops projects linked to existing public project pages.
* ``seed_ops_demo`` - development only: the Crater Cruncher Rover programme with
  teams, people, assets, milestones and work in every state. Seeded logins use the
  documented development password.
"""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from decimal import Decimal

from sqlalchemy import func
from werkzeug.security import generate_password_hash

from asme.extensions import db
from asme.models import Project as PublicProject
from asme.models import User
from asme.ops import authz
from asme.ops.models import Asset, Category, Location, Membership, Milestone, OpsProject, OpsProjectMember, Team, TeamMember, WorkOrder, WorkOrderCounter
from asme.ops.schemas import AssetCreate, CostInput, ProjectCreate, TeamCreate, TeamMemberInput, WorkOrderComplete, WorkOrderCreate
from asme.ops.services import assets as assets_service
from asme.ops.services import directory
from asme.ops.services import projects as projects_service
from asme.ops.services import work_orders as wo_service
from asme.ops.services.organizations import ensure_all_memberships
from asme.ops.tenancy import default_organization, ensure_membership, role_by_key

DEV_PASSWORD = "ChangeMe123!"


def seed_ops_defaults(commit: bool = True) -> dict:
    org = default_organization()
    authz.ensure_system_roles(org)
    created_memberships = ensure_all_memberships(org)
    directory.default_location(org.id)
    categories_created = directory.ensure_default_categories(org.id)
    types_created = assets_service.ensure_default_asset_types(org.id)
    if db.session.get(WorkOrderCounter, org.id) is None:
        db.session.add(WorkOrderCounter(organization_id=org.id, next_number=1))
    linked = 0
    for public in PublicProject.query.order_by(PublicProject.id.asc()).all():
        if OpsProject.query.filter_by(public_project_id=public.id).first():
            continue
        code = "".join(w[0] for w in re.findall(r"[A-Za-z0-9]+", public.title))[:6].upper() or f"P{public.id}"
        base, n = code, 2
        while OpsProject.query.filter(OpsProject.organization_id == org.id, func.lower(OpsProject.code) == code.lower()).first():
            code = f"{base}{n}"
            n += 1
        db.session.add(
            OpsProject(
                organization_id=org.id, name=public.title, code=code, description=public.summary, status="active" if (public.status or "").lower() != "archived" else "archived",
                public_project_id=public.id, visibility="public",
            )
        )
        linked += 1
    if commit:
        db.session.commit()
    return {"memberships": created_memberships, "categories": categories_created, "asset_types": types_created, "projects_linked": linked}


# --------------------------------------------------------------------------- demo


def _user(name, email, role_key, org, *, major="Mechanical Engineering", grad=2027, legacy_role=None):
    email = email.lower()
    user = User.query.filter(func.lower(User.email) == email).first()
    role = role_by_key(org, role_key)
    if user is None:
        user = User(
            name=name, email=email, username=email.split("@")[0], password_hash=generate_password_hash(DEV_PASSWORD),
            role=legacy_role or authz.legacy_role_for(role_key), is_active=True, major=major, graduation_year=grad,
        )
        db.session.add(user)
        db.session.flush()
    membership = ensure_membership(user, org)
    membership.role_id = role.id
    membership.member_status = "active"
    db.session.flush()
    return user


def _ctx(user, org):
    membership = ensure_membership(user, org)
    return authz.build_context(user, org, membership)


def _team(ctx, name, members, parent=None, project=None, color=None, description=None):
    team = Team.query.filter_by(organization_id=ctx.organization.id, name=name).first()
    if team is None:
        team = directory.create_team(ctx, TeamCreate(name=name, description=description, parent_team_id=parent.id if parent else None, project_id=project.id if project else None, color=color, members=[TeamMemberInput(user_id=u.id, is_lead=lead) for u, lead in members]))
    return team


def _location(ctx, name, parent=None, building=None, room=None):
    row = Location.query.filter_by(organization_id=ctx.organization.id, name=name).first()
    if row is None:
        from asme.ops.schemas import LocationCreate

        row = directory.create_location(ctx, LocationCreate(name=name, parent_location_id=parent.id if parent else None, building=building, room=room))
    return row


def _asset(ctx, name, code, **kwargs):
    row = Asset.query.filter_by(organization_id=ctx.organization.id, code=code).first()
    if row is None:
        row = assets_service.create_asset(ctx, AssetCreate(name=name, code=code, **kwargs))
    return row


def seed_ops_demo(commit: bool = True) -> dict:
    """Idempotent by code/email/title; re-running adds nothing new."""
    seed_ops_defaults(commit=False)
    org = default_organization()
    admin = User.query.filter(User.role == "admin").order_by(User.id.asc()).first()
    if admin is None:
        admin = _user("Ada Admin", "admin@uiowa.edu", "chapter_admin", org)
    ensure_membership(admin, org).role_id = role_by_key(org, "chapter_admin").id
    db.session.flush()
    actx = _ctx(admin, org)

    priya = _user("Priya Natarajan", "pnatarajan@uiowa.edu", "project_lead", org)
    marcus = _user("Marcus Bell", "mbell@uiowa.edu", "team_lead", org)
    elena = _user("Elena Ortiz", "eortiz@uiowa.edu", "team_lead", org, major="Mechanical Engineering")
    jordan = _user("Jordan Kim", "jkim@uiowa.edu", "team_lead", org, major="Electrical Engineering")
    sam = _user("Sam Rivera", "srivera@uiowa.edu", "team_lead", org, major="Computer Science and Engineering")
    avery = _user("Avery Johnson", "avery@uiowa.edu", "full_member", org)
    taylor = _user("Taylor Kim", "taylor@uiowa.edu", "full_member", org)
    noah = _user("Noah Patel", "npatel@uiowa.edu", "full_member", org, grad=2028)
    grace = _user("Grace Lin", "glin@uiowa.edu", "full_member", org, grad=2028, major="Electrical Engineering")
    omar = _user("Omar Haddad", "ohaddad@uiowa.edu", "full_member", org, major="Computer Science and Engineering")
    riley = _user("Riley Shaw", "rshaw@uiowa.edu", "safety_officer", org)
    dev = _user("Dev Anand", "danand@uiowa.edu", "inventory_manager", org)
    maya = _user("Maya Chen", "mchen@uiowa.edu", "treasurer", org)
    carver = _user("Helen Carver", "hcarver@uiowa.edu", "faculty_advisor", org, grad=None)
    chris = _user("Chris Doyle", "cdoyle@uiowa.edu", "shop_operator", org, grad=2029)
    exec_officer = _user("Brayden Nagra", "brayden-nagra@uiowa.edu", "executive_officer", org)

    # locations
    esc = _location(actx, "Engineering Student Center", building="Seamans Center")
    robotics = _location(actx, "Robotics Lab", parent=esc, building="Seamans Center", room="G440")
    shop = _location(actx, "Machine Shop", parent=esc, building="Seamans Center", room="1245")
    bench = _location(actx, "Electronics Bench", parent=robotics, building="Seamans Center", room="G440-B")
    storage = _location(actx, "ASME Storage", building="Seamans Center", room="B012")
    field = _location(actx, "Test Field", building="Hawkeye Recreation Fields")

    # projects
    def _project(code, name, **kw):
        row = OpsProject.query.filter_by(organization_id=org.id, code=code).first()
        if row is None:
            public = PublicProject.query.filter_by(slug=kw.pop("public_slug", None) or "").first()
            linked = OpsProject.query.filter_by(public_project_id=public.id).first() if public else None
            if linked is not None:
                # seed_ops_defaults() already mirrored the public project; adopt it as the demo project.
                linked.code, linked.name = code, name
                for key, value in kw.items():
                    setattr(linked, key, value)
                db.session.flush()
                return linked
            row = projects_service.create_project(actx, ProjectCreate(name=name, code=code, public_project_id=public.id if public else None, **kw))
        return row

    ccr = _project(
        "CCR", "Crater Cruncher Rover", description="Additive-manufacturing-heavy rover that excavates and deposits regolith simulant in a crater arena.",
        lead_user_id=priya.id, faculty_advisor_user_id=carver.id, competition="Lunar Crater Challenge 2027", academic_year="2026-27", risk_level="medium",
        start_date=date(2026, 8, 24), target_date=date(2027, 5, 15), budget_amount=Decimal("12000"), budget_code="ASME-CCR-27",
        repository_url="https://github.com/modhaneel072/crater-cruncher", cad_url="https://uiowa.sharepoint.com/asme/ccr/cad", public_slug="rover",
    )
    ops = _project("OPS", "General Chapter Operations", description="Shop upkeep, equipment maintenance, inventory and safety work that is not tied to a single build.", lead_user_id=exec_officer.id, academic_year="2026-27", risk_level="low", status="active")
    showcase = _project("SHOW", "Fall Engineering Showcase", description="Chapter booth, rover demo and recruiting at the College of Engineering showcase.", lead_user_id=exec_officer.id, academic_year="2026-27", risk_level="low", start_date=date(2026, 9, 1), target_date=date(2026, 10, 24), budget_amount=Decimal("800"))
    if ccr.public_project_id is None:
        public = PublicProject.query.filter_by(slug="rover").first()
        if public and OpsProject.query.filter_by(public_project_id=public.id).first() is None:
            ccr.public_project_id = public.id
    projects_service.set_members(
        actx, ccr,
        [
            __import__("asme.ops.schemas", fromlist=["ProjectMemberInput"]).ProjectMemberInput(user_id=u.id, project_role=r)
            for u, r in ((priya, "lead"), (carver, "advisor"), (marcus, "member"), (elena, "member"), (jordan, "member"), (sam, "member"), (avery, "member"), (taylor, "member"), (noah, "member"), (grace, "member"), (omar, "member"), (riley, "observer"))
        ],
    ) if len(ccr.members) < 5 else None

    # teams
    board = _team(actx, "Executive Board", [(exec_officer, True), (priya, False), (maya, False), (riley, False), (dev, False)], color="#ffcd00")
    rover_team = _team(actx, "Crater Cruncher Rover", [(priya, True), (marcus, False), (elena, False), (jordan, False), (sam, False)], project=ccr, color="#0878d1", description="Parent team for all rover subteams.")
    wheels = _team(actx, "Wheels and Mobility", [(marcus, True), (avery, False), (noah, False)], parent=rover_team, project=ccr, color="#0891b2")
    arm = _team(actx, "Robotic Arm", [(elena, True), (taylor, False)], parent=rover_team, project=ccr, color="#7c5ce7")
    electrical = _team(actx, "Electrical", [(jordan, True), (grace, False)], parent=rover_team, project=ccr, color="#e58a00")
    software = _team(actx, "Software and Autonomy", [(sam, True), (omar, False)], parent=rover_team, project=ccr, color="#00a878")
    fabrication = _team(actx, "Fabrication", [(marcus, True), (chris, False), (avery, False)], project=ops, color="#475569")
    outreach = _team(actx, "Events and Outreach", [(exec_officer, True), (taylor, False)], project=showcase, color="#db2777")
    procurement = _team(actx, "Inventory and Procurement", [(dev, True), (maya, False)], project=ops, color="#0f766e")
    safety = _team(actx, "Safety", [(riley, True), (chris, False)], project=ops, color="#d84a4a")

    # assets
    types = {t.name: t for t in assets_service.list_asset_types(actx)}
    rover = _asset(actx, "Crater Cruncher Rover", "CCR-ROVER", project_id=ccr.id, location_id=robotics.id, responsible_team_id=rover_team.id, criticality="critical", asset_type_ids=[types["Vehicle / Rover"].id], owner_user_id=priya.id, description="Competition rover, 2026-27 build.")
    chassis = _asset(actx, "Chassis", "CCR-CHASSIS", parent_asset_id=rover.id, project_id=ccr.id, location_id=robotics.id, responsible_team_id=wheels.id, criticality="high", asset_type_ids=[types["Subsystem"].id], manufacturer="ASME @ UIowa", model="Welded 6061 frame")
    mobility = _asset(actx, "Mobility System", "CCR-MOB", parent_asset_id=rover.id, project_id=ccr.id, location_id=robotics.id, responsible_team_id=wheels.id, criticality="high", asset_type_ids=[types["Subsystem"].id])
    for pos, code in (("Front Left", "FL"), ("Front Right", "FR"), ("Rear Left", "RL"), ("Rear Right", "RR")):
        _asset(actx, f"{pos} Wheel Module", f"CCR-WHL-{code}", parent_asset_id=mobility.id, project_id=ccr.id, location_id=robotics.id, responsible_team_id=wheels.id, criticality="medium", asset_type_ids=[types["Subsystem"].id])
    arm_asset = _asset(actx, "Rover Robotic Arm", "CCR-ARM", parent_asset_id=rover.id, project_id=ccr.id, location_id=robotics.id, responsible_team_id=arm.id, criticality="high", asset_type_ids=[types["Subsystem"].id])
    for name, code in (("Shoulder Assembly", "SHO"), ("Elbow Assembly", "ELB"), ("End Effector", "EE")):
        _asset(actx, name, f"CCR-ARM-{code}", parent_asset_id=arm_asset.id, project_id=ccr.id, location_id=robotics.id, responsible_team_id=arm.id, criticality="medium", asset_type_ids=[types["Subsystem"].id])
    enclosure = _asset(actx, "Rover Electrical Enclosure", "CCR-ELEC", parent_asset_id=rover.id, project_id=ccr.id, location_id=bench.id, responsible_team_id=electrical.id, criticality="high", asset_type_ids=[types["Subsystem"].id])
    compute = _asset(actx, "Compute and Control", "CCR-CTRL", parent_asset_id=rover.id, project_id=ccr.id, location_id=bench.id, responsible_team_id=software.id, criticality="high", asset_type_ids=[types["Subsystem"].id], model="Raspberry Pi 5 + ESP32-S3")
    printer1 = _asset(actx, "3D Printer 01", "PRN-01", location_id=robotics.id, responsible_team_id=fabrication.id, criticality="medium", asset_type_ids=[types["3D Printer"].id], manufacturer="Bambu Lab", model="H2S", serial_number="H2S-2025-0417", purchase_date=date(2025, 8, 12), purchase_cost=Decimal("1899"))
    printer2 = _asset(actx, "3D Printer 02", "PRN-02", location_id=robotics.id, responsible_team_id=fabrication.id, criticality="medium", asset_type_ids=[types["3D Printer"].id], manufacturer="Bambu Lab", model="P1S", serial_number="P1S-2024-1188", purchase_date=date(2024, 9, 3), purchase_cost=Decimal("699"))
    station = _asset(actx, "Soldering Station 01", "SOL-01", location_id=bench.id, responsible_team_id=electrical.id, criticality="low", asset_type_ids=[types["Soldering Station"].id], manufacturer="Hakko", model="FX-951")
    charger = _asset(actx, "Battery Charger 01", "CHG-01", location_id=bench.id, responsible_team_id=electrical.id, criticality="medium", asset_type_ids=[types["Battery / Charger"].id], manufacturer="ISDT", model="Q8 Max")
    scope = _asset(actx, "Oscilloscope 01", "OSC-01", location_id=bench.id, responsible_team_id=electrical.id, criticality="low", asset_type_ids=[types["Test Equipment"].id], manufacturer="Rigol", model="DS1054Z")
    psu = _asset(actx, "Bench Power Supply 01", "PSU-01", location_id=bench.id, responsible_team_id=electrical.id, criticality="low", asset_type_ids=[types["Test Equipment"].id], manufacturer="Rigol", model="DP832")

    # milestones
    if not ccr.milestones:
        from asme.ops.schemas import MilestoneCreate

        for name, due, status, weight, owner in (
            ("Critical design review", date(2026, 10, 30), "complete", 2, priya),
            ("Chassis welded and squared", date(2026, 11, 20), "in_progress", 2, marcus),
            ("First drive test", date(2027, 1, 30), "planned", 3, marcus),
            ("Arm integration", date(2027, 3, 6), "planned", 3, elena),
            ("Autonomy field test", date(2027, 4, 10), "planned", 2, sam),
            ("Competition packing", date(2027, 5, 8), "planned", 1, priya),
        ):
            projects_service.create_milestone(actx, ccr, MilestoneCreate(name=name, due_date=due, status=status, weight=weight, owner_user_id=owner.id))
        projects_service.create_milestone(actx, showcase, MilestoneCreate(name="Booth reserved", due_date=date(2026, 9, 30), status="complete", owner_user_id=exec_officer.id))
        projects_service.create_milestone(actx, showcase, MilestoneCreate(name="Showcase day", due_date=date(2026, 10, 24), status="planned", weight=3, owner_user_id=exec_officer.id))

    cats = {c.name: c for c in directory.list_categories(actx)}
    now = datetime.utcnow()

    def _wo(title, actor, *, status="OPEN", created_days_ago=0, complete_note=None, time_minutes=None, costs=(), comments=(), assignee_users=(), **kw):
        existing = WorkOrder.query.filter_by(organization_id=org.id, title=title).first()
        if existing:
            return existing
        c = _ctx(actor, org)
        wo = wo_service.create(c, WorkOrderCreate(title=title, assignee_user_ids=[u.id for u in assignee_users], status="DRAFT" if status == "DRAFT" else "OPEN", **kw))
        if created_days_ago:
            wo.created_at = now - timedelta(days=created_days_ago)
            wo.last_activity_at = wo.created_at
            for h in wo.status_history:
                h.changed_at = wo.created_at
            db.session.commit()
        for author, text in comments:
            wo_service.add_comment(_ctx(author, org), wo, text)
        if status in {"IN_PROGRESS", "ON_HOLD", "DONE"}:
            wo_service.start(c, wo)
        if status == "ON_HOLD":
            wo_service.hold(c, wo, "Waiting on parts")
        if status == "DONE":
            wo_service.complete(c, wo, WorkOrderComplete(completion_note=complete_note, time_minutes=time_minutes, costs=[CostInput(**x) for x in costs]))
            wo.completed_at = now - timedelta(days=max(created_days_ago - 2, 0), hours=3)
            db.session.commit()
        if status == "CANCELED":
            wo_service.cancel(c, wo, "Superseded by harness redesign")
        return wo

    _wo("Inspect wheel hub fasteners", marcus, project_id=ccr.id, team_id=wheels.id, primary_asset_id=mobility.id, location_id=robotics.id, priority="HIGH", work_type="INSPECTION", category_ids=[cats["Mechanical"].id, cats["Inspection"].id], due_at=now + timedelta(days=2), estimated_minutes=45, assignee_users=(avery, noah), created_days_ago=3, description="Torque-check all M6 hub fasteners after the last field test; replace any nylock that has been reused twice.", comments=((avery, "Torque wrench is in Tool Drawer 3, calibrated last month."),))
    _wo("Validate robotic arm current limits", elena, status="IN_PROGRESS", project_id=ccr.id, team_id=arm.id, primary_asset_id=arm_asset.id, location_id=bench.id, priority="MEDIUM", work_type="PROJECT", category_ids=[cats["Electrical"].id, cats["Embedded Systems"].id], due_at=now + timedelta(days=6), estimated_minutes=180, assignee_users=(taylor, grace), created_days_ago=5, description="Set per-joint current limits on the shoulder and elbow drivers; log stall current at 12 V and 16 V.", comments=((grace, "Stall test at 12 V done, 14.2 A on the shoulder. 16 V tomorrow."),))
    _wo("Update telemetry packet parser", sam, status="IN_PROGRESS", project_id=ccr.id, team_id=software.id, primary_asset_id=compute.id, priority="MEDIUM", work_type="PROJECT", category_ids=[cats["Software"].id], due_at=now + timedelta(days=9), estimated_minutes=240, assignee_users=(omar,), created_days_ago=4, description="Parser drops packets when the IMU frame arrives before the GPS frame. Add sequence numbers and a reorder buffer.")
    _wo("Print spare sensor mount", marcus, project_id=ccr.id, team_id=fabrication.id, primary_asset_id=printer1.id, location_id=robotics.id, priority="LOW", work_type="PROJECT", category_ids=[cats["Fabrication"].id], due_at=now + timedelta(days=12), estimated_minutes=30, assignee_users=(chris,), created_days_ago=2, description="Two copies of sensor_mount_v3.3mf in PETG, 40 % infill.")
    _wo("Run pre-drive safety inspection", riley, project_id=ccr.id, team_id=safety.id, primary_asset_id=rover.id, location_id=field.id, priority="CRITICAL", work_type="SAFETY", category_ids=[cats["Safety"].id, cats["Inspection"].id], due_at=now - timedelta(days=1, hours=4), estimated_minutes=40, assignee_users=(riley, marcus), created_days_ago=6, description="E-stop, battery strap, wheel nut witness marks, arm stow latch. No drive test without a signed inspection.")
    _wo("Inventory M4 fasteners", dev, project_id=ops.id, team_id=procurement.id, location_id=storage.id, priority="LOW", work_type="PROCUREMENT", category_ids=[cats["Procurement"].id], due_at=now + timedelta(days=20), estimated_minutes=60, assignee_users=(dev,), created_days_ago=1, description="Count M4x10/16/20 socket head and nylocks; reorder anything under 100.")
    _wo("Replace 3D printer 01 nozzle", marcus, status="DONE", project_id=ops.id, team_id=fabrication.id, primary_asset_id=printer1.id, location_id=robotics.id, priority="MEDIUM", work_type="REACTIVE", category_ids=[cats["Damage"].id, cats["Fabrication"].id], due_at=now - timedelta(days=8), estimated_minutes=30, assignee_users=(chris,), created_days_ago=12, complete_note="Swapped 0.4 mm hardened nozzle; first layer test clean.", time_minutes=35, costs=({"type": "part", "amount": "24.99", "description": "Hardened steel nozzle 0.4 mm"},))
    _wo("Calibrate 3D printer 02 bed", marcus, status="DONE", project_id=ops.id, team_id=fabrication.id, primary_asset_id=printer2.id, location_id=robotics.id, priority="LOW", work_type="PREVENTIVE", category_ids=[cats["Preventive"].id], due_at=now - timedelta(days=15), estimated_minutes=20, assignee_users=(avery,), created_days_ago=18, complete_note="Auto-level plus manual tram; 0.05 mm across the plate.", time_minutes=25)
    _wo("Battery pack health check", jordan, status="ON_HOLD", project_id=ccr.id, team_id=electrical.id, primary_asset_id=charger.id, location_id=bench.id, priority="HIGH", work_type="INSPECTION", category_ids=[cats["Electrical"].id, cats["Inspection"].id], due_at=now + timedelta(days=4), estimated_minutes=90, assignee_users=(grace,), created_days_ago=7, description="Internal-resistance sweep on both 6S packs. Blocked until the replacement balance leads arrive.", comments=((jordan, "Balance leads ordered from the vendor, ETA Thursday."),))
    _wo("Chassis weld inspection", riley, status="DONE", project_id=ccr.id, team_id=safety.id, primary_asset_id=chassis.id, location_id=shop.id, priority="HIGH", work_type="INSPECTION", category_ids=[cats["Inspection"].id, cats["Mechanical"].id], due_at=now - timedelta(days=20), estimated_minutes=60, assignee_users=(riley,), created_days_ago=24, complete_note="Dye-penetrant on all eight corner joints, no indications.", time_minutes=70, costs=({"type": "other", "amount": "18.50", "description": "Dye penetrant kit"},))
    _wo("Showcase: reserve demo space and power", exec_officer, project_id=showcase.id, team_id=outreach.id, location_id=esc.id, priority="MEDIUM", work_type="EVENT", category_ids=[cats["Event"].id], due_at=now + timedelta(days=14), estimated_minutes=30, assignee_users=(taylor,), created_days_ago=9)
    _wo("Showcase: transport rover and stands", exec_officer, project_id=showcase.id, team_id=outreach.id, primary_asset_id=rover.id, priority="MEDIUM", work_type="EVENT", category_ids=[cats["Event"].id], due_at=now + timedelta(days=30), estimated_minutes=120, assignee_users=(marcus, taylor), created_days_ago=9)
    _wo("Monthly 3D printer inspection", dev, project_id=ops.id, team_id=fabrication.id, primary_asset_id=printer1.id, location_id=robotics.id, priority="LOW", work_type="PREVENTIVE", category_ids=[cats["Preventive"].id], due_at=now + timedelta(days=5), estimated_minutes=25, assignee_users=(chris,), created_days_ago=3, recurrence={"frequency": "monthly", "interval": 1, "mode": "fixed"}, description="Belts, rods, nozzle wear, filament path; wipe the bed.")
    _wo("Semester soldering station inspection", riley, status="DONE", project_id=ops.id, team_id=safety.id, primary_asset_id=station.id, location_id=bench.id, priority="MEDIUM", work_type="PREVENTIVE", category_ids=[cats["Preventive"].id, cats["Safety"].id], due_at=now - timedelta(days=30), estimated_minutes=20, assignee_users=(riley,), created_days_ago=33, complete_note="Tip replaced, fume extractor filter changed.", time_minutes=20, costs=({"type": "part", "amount": "31.20", "description": "Extractor filter + T12 tip"},))
    hub = _wo("Wheel hub CAD revision (v4)", marcus, project_id=ccr.id, team_id=wheels.id, primary_asset_id=mobility.id, priority="MEDIUM", work_type="PROJECT", category_ids=[cats["Mechanical"].id, cats["Documentation"].id], due_at=now + timedelta(days=16), estimated_minutes=300, assignee_users=(noah,), created_days_ago=8, parent_completion_policy="auto", description="Thicker hub flange and captured nut pockets. Parent of the CAD / print / fit-test chain.")
    if not hub.children:
        from asme.ops.schemas import SubWorkOrderCreate

        mc = _ctx(marcus, org)
        wo_service.create_sub(mc, hub, SubWorkOrderCreate(title="Update hub CAD and drawings", assignee_user_ids=[noah.id], due_at=now + timedelta(days=6), priority="MEDIUM", estimated_minutes=120))
        wo_service.create_sub(mc, hub, SubWorkOrderCreate(title="Print v4 hub prototype", assignee_user_ids=[chris.id], due_at=now + timedelta(days=10), priority="MEDIUM", estimated_minutes=60))
        wo_service.create_sub(mc, hub, SubWorkOrderCreate(title="Fit test on rear-left module", assignee_user_ids=[avery.id], due_at=now + timedelta(days=15), priority="MEDIUM", estimated_minutes=90))
    _wo("Rework arm wire harness", jordan, status="CANCELED", project_id=ccr.id, team_id=electrical.id, primary_asset_id=arm_asset.id, priority="MEDIUM", work_type="PROJECT", category_ids=[cats["Electrical"].id], created_days_ago=14, assignee_users=(grace,))
    _wo("Draft competition packing checklist", priya, status="DRAFT", project_id=ccr.id, team_id=rover_team.id, priority="NONE", work_type="DOCUMENTATION", category_ids=[cats["Documentation"].id], due_at=now + timedelta(days=60), created_days_ago=1)
    _wo("Replace oscilloscope probe", jordan, project_id=ops.id, team_id=electrical.id, primary_asset_id=scope.id, location_id=bench.id, priority="LOW", work_type="REACTIVE", category_ids=[cats["Damage"].id], estimated_minutes=10, created_days_ago=2, description="Channel 2 probe ground clip is broken; swap with the spare in the drawer and order a replacement.")

    if commit:
        db.session.commit()
    return {
        "users": User.query.count(),
        "teams": Team.query.filter_by(organization_id=org.id).count(),
        "locations": Location.query.filter_by(organization_id=org.id).count(),
        "assets": Asset.query.filter_by(organization_id=org.id).count(),
        "projects": OpsProject.query.filter_by(organization_id=org.id).count(),
        "work_orders": WorkOrder.query.filter_by(organization_id=org.id).count(),
        "password": DEV_PASSWORD,
    }
