from datetime import datetime, timedelta

from flask import Flask
from werkzeug.security import generate_password_hash

from models import (
    Announcement,
    AttendanceRecord,
    Event,
    Item,
    ItemTag,
    Member,
    NFCTag,
    PrintRequest,
    Project,
    Transaction,
    User,
    db,
)


def create_app():
    app = Flask(__name__)
    app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///inventory.db"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    db.init_app(app)
    return app


def seed_members():
    if Member.query.count() > 0:
        return
    rows = [
        Member(name="Avery Johnson", email="avery@uiowa.edu", member_class="ME Junior", nfc_tag="ASME-1001"),
        Member(name="Morgan Lee", email="morgan@uiowa.edu", member_class="Team Lead", nfc_tag="ASME-1002"),
        Member(name="Taylor Kim", email="taylor@uiowa.edu", member_class="Senior", nfc_tag="ASME-1003"),
        Member(name="Jordan Patel", email="jordan@uiowa.edu", member_class="Admin", nfc_tag="ASME-1004"),
    ]
    db.session.add_all(rows)


def seed_items():
    if Item.query.count() > 0:
        return
    rows = [
        Item(
            name="Arduino Uno Kit",
            description="Controller board + cable kit for controls prototyping.",
            category="Electronics",
            location="Electronics Bin A",
            item_condition="good",
            total_qty=8,
            available_qty=8,
        ),
        Item(
            name="Digital Calipers",
            description="150mm digital calipers.",
            category="Tools",
            location="Tool Drawer 2",
            item_condition="good",
            total_qty=4,
            available_qty=4,
        ),
        Item(
            name="Safety Goggles",
            description="Shop safety eyewear.",
            category="Safety",
            location="PPE Cabinet",
            item_condition="new",
            total_qty=20,
            available_qty=20,
        ),
    ]
    db.session.add_all(rows)


def seed_users():
    if User.query.count() > 0:
        return
    default_password_hash = generate_password_hash("ChangeMe123!")
    members = Member.query.order_by(Member.id.asc()).all()
    for member in members:
        role = "member"
        role_hint = (member.member_class or "").lower()
        if "lead" in role_hint:
            role = "team_leader"
        if "admin" in role_hint:
            role = "admin"
        db.session.add(
            User(
                name=member.name,
                email=member.email,
                password_hash=default_password_hash,
                role=role,
                member_id=member.id,
                is_active=True,
            )
        )


def seed_projects():
    if Project.query.count() > 0:
        return
    db.session.add_all(
        [
            Project(
                slug="rover",
                title="Rover",
                project_type="Rover",
                summary="Mobility platform development with drivetrain and controls integration.",
                description="Rover subsystem owners run design, build, and test cycles across each semester.",
                status="Active",
                timeline="Concept, CAD, Fabrication, Integration, Testing",
                gallery_json='["https://example.com/rover-1.jpg"]',
                lead_name="Rover Lead",
            ),
            Project(
                slug="arm",
                title="Robotic Arm",
                project_type="Arm",
                summary="Manipulator mechanism and controls reliability improvements.",
                description="Arm team is focused on repeatability, payload handling, and packaging constraints.",
                status="Prototype",
                timeline="Design, Build, Controls Tuning, Validation",
                gallery_json='["https://example.com/arm-1.jpg"]',
                lead_name="Arm Lead",
            ),
            Project(
                slug="manufacturing",
                title="Manufacturing",
                project_type="Manufacturing",
                summary="Process quality and fabrication readiness across all teams.",
                description="Manufacturing members support parts, fixtures, and process documentation.",
                status="In Progress",
                timeline="Planning, Material Prep, Production, QA",
                gallery_json='["https://example.com/manufacturing-1.jpg"]',
                lead_name="Manufacturing Lead",
            ),
        ]
    )


def seed_announcements():
    if Announcement.query.count() > 0:
        return
    db.session.add(
        Announcement(
            title="Kickoff Complete",
            body="Spring kickoff is complete. Subteam onboarding and project planning are live.",
            is_published=True,
            show_on_public=True,
            show_on_member=True,
            published_at=datetime.utcnow(),
        )
    )


def seed_nfc_tags():
    if NFCTag.query.count() > 0:
        return
    users = User.query.order_by(User.id.asc()).limit(2).all()
    for idx, user in enumerate(users, start=1):
        db.session.add(
            NFCTag(
                tag_uid=f"USER-NFC-{idx:04d}",
                user_id=user.id,
                active=True,
                assigned_at=datetime.utcnow(),
            )
        )


def seed_item_tags():
    if ItemTag.query.count() > 0:
        return
    items = Item.query.order_by(Item.id.asc()).limit(3).all()
    for item in items:
        db.session.add(ItemTag(item_id=item.id, tag_value=f"item_id:{item.id}", source="seed"))


def seed_sample_transactions():
    if Transaction.query.count() > 0:
        return
    user = User.query.filter_by(role="member").first() or User.query.first()
    member = user.member if user else None
    item = Item.query.first()
    if not user or not member or not item:
        return
    item.available_qty = max(0, item.available_qty - 1)
    now = datetime.utcnow() - timedelta(days=1)
    db.session.add(
        Transaction(
            member_id=member.id,
            user_id=user.id,
            item_id=item.id,
            action="checkout",
            qty=1,
            status="OUT",
            timestamp=now,
            checkout_time=now,
            due_date=(datetime.utcnow() + timedelta(days=6)).date(),
            notes="Seed checkout",
            checkout_notes="Seed checkout",
        )
    )


def seed_events_and_attendance():
    if Event.query.count() == 0:
        admin = User.query.filter_by(role="admin").first() or User.query.first()
        now = datetime.utcnow()
        db.session.add_all(
            [
                Event(
                    title="Weekly Build Meeting",
                    description="Subsystem updates and blocker review.",
                    location="Robotics Room",
                    status="scheduled",
                    start_time=now + timedelta(days=1),
                    end_time=now + timedelta(days=1, hours=1),
                    created_by_user_id=admin.id if admin else None,
                ),
                Event(
                    title="Print Queue Review",
                    description="Priority print approvals for this week.",
                    location="Fluids Lab",
                    status="requested",
                    start_time=now + timedelta(days=2),
                    end_time=now + timedelta(days=2, hours=1),
                    created_by_user_id=admin.id if admin else None,
                ),
            ]
        )
        db.session.flush()

    if AttendanceRecord.query.count() == 0:
        event = Event.query.order_by(Event.id.asc()).first()
        user = User.query.order_by(User.id.asc()).first()
        member = user.member if user else None
        if event and user:
            db.session.add(
                AttendanceRecord(
                    event_id=event.id,
                    user_id=user.id,
                    member_id=member.id if member else None,
                    checkin_method="manual",
                    checkin_time=datetime.utcnow(),
                )
            )


def seed_print_requests():
    if PrintRequest.query.count() > 0:
        return
    user = User.query.filter_by(role="member").first() or User.query.first()
    member = user.member if user else None
    if not user:
        return
    db.session.add(
        PrintRequest(
            user_id=user.id,
            member_id=member.id if member else None,
            printer_type="P1S",
            file_link="https://example.com/demo-part.stl",
            material="PLA",
            color="Black",
            infill_percent=25,
            priority="normal",
            status="submitted",
            notes="Demo queue seed request.",
        )
    )


def seed_data():
    seed_members()
    seed_items()
    db.session.flush()
    seed_users()
    db.session.flush()
    seed_projects()
    seed_announcements()
    seed_nfc_tags()
    seed_item_tags()
    seed_sample_transactions()
    seed_events_and_attendance()
    seed_print_requests()
    db.session.commit()


if __name__ == "__main__":
    app = create_app()
    with app.app_context():
        db.create_all()
        seed_data()
        print("Database created and seeded: inventory.db")
