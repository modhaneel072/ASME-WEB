"""Local development seed: sample members, items, a project and a couple of events.

    python db_init.py

Runs migrations + default seeds first (same as ``python manage.py upgrade``),
then adds example rows only if the tables are empty. Never run against production.
"""

from __future__ import annotations

from datetime import datetime, timedelta

from werkzeug.security import generate_password_hash

from asme import create_app
from asme.extensions import db
from asme.models import Event, Item, Member, Project, User
from asme.services import bootstrap
from asme.services.onboarding import engine

SAMPLE_PASSWORD = "ChangeMe123!"


def seed_members():
    if Member.query.count() > 0:
        return
    rows = [
        Member(name="Avery Johnson", email="avery@uiowa.edu", member_class="ME Junior", nfc_tag="ASME-1001"),
        Member(name="Morgan Lee", email="morgan@uiowa.edu", member_class="Team Lead", nfc_tag="ASME-1002"),
        Member(name="Taylor Kim", email="taylor@uiowa.edu", member_class="Senior", nfc_tag="ASME-1003"),
    ]
    db.session.add_all(rows)
    db.session.flush()
    for member in rows:
        if User.query.filter_by(email=member.email).first():
            continue
        role = "team_leader" if "lead" in member.member_class.lower() else "member"
        db.session.add(
            User(
                name=member.name,
                email=member.email,
                username=member.email.split("@")[0],
                password_hash=generate_password_hash(SAMPLE_PASSWORD),
                role=role,
                is_active=True,
                member_id=member.id,
                nfc_uid=member.nfc_tag,
                major="Mechanical Engineering",
                graduation_year=2027,
            )
        )


def seed_items():
    if Item.query.count() > 0:
        return
    rows = [
        Item(name="Arduino Uno Kit", description="Controller board + cable kit.", category="Electronics", location="Electronics Bin A", total_qty=8, available_qty=8),
        Item(name="Digital Calipers", description="150mm digital calipers.", category="Tools", location="Tool Drawer 2", total_qty=4, available_qty=4),
        Item(name="Cordless Drill", description="18V drill with two batteries.", category="Power Tools", location="Cabinet 1", total_qty=2, available_qty=2),
        Item(name="PLA Filament 1kg", description="Black PLA, 1.75mm.", category="Consumables", location="Printer Shelf", total_qty=10, available_qty=10, item_type="consumable", is_consumable=True, min_stock_threshold=2),
        Item(name="Torque Wrench", description="1/2 in drive, 20-150 ft-lb.", category="Tools", location="Tool Drawer 3", total_qty=1, available_qty=1),
    ]
    db.session.add_all(rows)


def seed_events():
    if Event.query.count() > 0:
        return
    start = datetime.now().replace(minute=0, second=0, microsecond=0) + timedelta(days=2, hours=2)
    db.session.add_all(
        [
            Event(title="General Body Meeting", kind="meeting", location="Robotics Room", status="scheduled", start_time=start, end_time=start + timedelta(hours=1)),
            Event(title="Rover build session", kind="build_session", location="Fluids Lab", status="scheduled", start_time=start + timedelta(days=2), end_time=start + timedelta(days=2, hours=3)),
        ]
    )


if __name__ == "__main__":
    app = create_app(auto_migrate=False, outbox_worker_enabled=False)
    with app.app_context():
        print("schema:", bootstrap.sync_schema())
        print("defaults:", bootstrap.seed_defaults())
        seed_members()
        seed_items()
        seed_events()
        db.session.commit()
        for user in User.query.filter(User.is_active.is_(True)).all():
            engine.evaluate_user(user)
        engine.evaluate_chapter()
        print(f"members={Member.query.count()} users={User.query.count()} items={Item.query.count()} projects={Project.query.count()} events={Event.query.count()}")
        print(f"sample logins use password: {SAMPLE_PASSWORD}")
