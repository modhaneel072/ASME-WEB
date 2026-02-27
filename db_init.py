from flask import Flask
from models import AttendanceScan, Item, Member, PrintJob, Transaction, db

def create_app():
    app = Flask(__name__)
    app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///inventory.db"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    db.init_app(app)
    return app

def seed_data():
    # Only add if tables are empty
    if Member.query.count() == 0:
        db.session.add_all([
            Member(name="Test Member", email="test@uiowa.edu", member_class="ME Junior", nfc_tag="ASME-0001"),
            Member(name="Officer Example", email="officer@uiowa.edu", member_class="Senior", nfc_tag="ASME-0002"),
        ])
    if Item.query.count() == 0:
        db.session.add_all([
            Item(name="Arduino Uno", category="Electronics", location="Bin A", total_qty=10, available_qty=10, nfc_tag=None),
            Item(name="Calipers", category="Tools", location="Drawer 2", total_qty=3, available_qty=3, nfc_tag=None),
        ])
    db.session.commit()

if __name__ == "__main__":
    app = create_app()
    with app.app_context():
        db.create_all()
        seed_data()
        print("✅ Database created: inventory.db (with sample data)")
