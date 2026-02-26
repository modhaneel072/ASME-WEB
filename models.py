from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date

db = SQLAlchemy()

class Member(db.Model):
    __tablename__ = "members"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(160), nullable=False, unique=True)
    member_class = db.Column(db.String(80), nullable=False)  # e.g., "Freshman", "Sophomore", "ME Junior", etc.

    # Optional NFC tag for member card
    nfc_tag = db.Column(db.String(120), nullable=True, unique=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class Item(db.Model):
    __tablename__ = "items"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(160), nullable=False)
    category = db.Column(db.String(80), nullable=True)
    location = db.Column(db.String(120), nullable=True)

    total_qty = db.Column(db.Integer, nullable=False, default=0)
    available_qty = db.Column(db.Integer, nullable=False, default=0)

    # NFC tag on a bin/tool group
    nfc_tag = db.Column(db.String(120), nullable=True, unique=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class Transaction(db.Model):
    __tablename__ = "transactions"
    id = db.Column(db.Integer, primary_key=True)

    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False)
    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False)

    action = db.Column(db.String(20), nullable=False)  # "checkout" or "return"
    qty = db.Column(db.Integer, nullable=False)

    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    due_date = db.Column(db.Date, nullable=True)
    notes = db.Column(db.String(300), nullable=True)

    member = db.relationship("Member")
    item = db.relationship("Item")
