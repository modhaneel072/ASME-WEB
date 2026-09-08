from datetime import datetime

from asme.extensions import db


class Member(db.Model):
    """Legacy roster profile.

    ``User`` is the canonical identity. This table is kept for one release as a
    compatibility layer for old kiosk flows and historical rows that reference
    ``member_id``; new code links everything to ``users.id``.
    """

    __tablename__ = "members"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(160), nullable=False, unique=True)
    member_class = db.Column(db.String(80), nullable=False)  # e.g., "Freshman", "ME Junior"

    nfc_tag = db.Column(db.String(120), nullable=True, unique=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(160), nullable=False)
    email = db.Column(db.String(160), nullable=False, unique=True, index=True)
    username = db.Column(db.String(80), nullable=True, unique=True, index=True)
    password_hash = db.Column(db.String(260), nullable=False)
    role = db.Column(db.String(30), nullable=False, default="member")
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    nfc_uid = db.Column(db.String(160), nullable=True, unique=True, index=True)
    major = db.Column(db.String(120), nullable=True)
    graduation_year = db.Column(db.Integer, nullable=True)
    phone = db.Column(db.String(40), nullable=True)
    exec_title = db.Column(db.String(160), nullable=True)
    exec_message = db.Column(db.String(500), nullable=True)
    headshot_url = db.Column(db.String(500), nullable=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=True, unique=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    last_login_at = db.Column(db.DateTime, nullable=True)

    member = db.relationship("Member", foreign_keys=[member_id])

    @property
    def display_name(self) -> str:
        return (self.name or self.email or f"user-{self.id}").strip()


class NFCTag(db.Model):
    __tablename__ = "nfc_tags"
    id = db.Column(db.Integer, primary_key=True)

    tag_uid = db.Column(db.String(160), nullable=False, unique=True, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    active = db.Column(db.Boolean, nullable=False, default=True)
    assigned_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    unassigned_at = db.Column(db.DateTime, nullable=True)
    assigned_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    notes = db.Column(db.String(260), nullable=True)

    user = db.relationship("User", foreign_keys=[user_id])
    assigned_by = db.relationship("User", foreign_keys=[assigned_by_user_id])


class PasswordResetToken(db.Model):
    __tablename__ = "password_reset_tokens"
    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    token = db.Column(db.String(200), nullable=False, unique=True, index=True)
    expires_at = db.Column(db.DateTime, nullable=False)
    used_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    user = db.relationship("User", foreign_keys=[user_id])
