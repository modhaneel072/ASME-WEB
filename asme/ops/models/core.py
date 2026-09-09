"""Organisation, identity & access, teams, locations, categories, and cross-cutting
records (comments, attachments, notifications, saved filters, audit events)."""

from __future__ import annotations

import json
from datetime import datetime

from asme.extensions import db
from asme.ops.models.base import OpsBase, new_uuid, utcnow


class Organization(db.Model):
    __tablename__ = "organizations"
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    name = db.Column(db.String(200), nullable=False)
    slug = db.Column(db.String(80), nullable=False, unique=True, index=True)
    logo_url = db.Column(db.String(500), nullable=True)
    timezone = db.Column(db.String(60), nullable=False, default="America/Chicago")
    academic_year_start_month = db.Column(db.Integer, nullable=False, default=8)
    settings_json = db.Column(db.Text, nullable=False, default="{}")
    created_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=utcnow, onupdate=utcnow, nullable=False)

    @property
    def settings(self) -> dict:
        try:
            value = json.loads(self.settings_json or "{}")
        except Exception:
            return {}
        return value if isinstance(value, dict) else {}

    @settings.setter
    def settings(self, value):
        self.settings_json = json.dumps(dict(value or {}), default=str)


class Permission(db.Model):
    __tablename__ = "permissions"
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    key = db.Column(db.String(80), nullable=False, unique=True, index=True)
    description = db.Column(db.String(300), nullable=True)


class Role(OpsBase, db.Model):
    __tablename__ = "roles"
    __table_args__ = (db.UniqueConstraint("organization_id", "name", name="uq_roles_org_name"),)
    name = db.Column(db.String(120), nullable=False)
    system_key = db.Column(db.String(60), nullable=True, index=True)
    description = db.Column(db.String(300), nullable=True)
    is_custom = db.Column(db.Boolean, nullable=False, default=False)
    rank = db.Column(db.Integer, nullable=False, default=0)  # display / sort order only

    permissions = db.relationship("RolePermission", back_populates="role", cascade="all, delete-orphan")


class RolePermission(db.Model):
    __tablename__ = "role_permissions"
    __table_args__ = (db.UniqueConstraint("role_id", "permission_id", name="uq_role_permission"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    role_id = db.Column(db.String(36), db.ForeignKey("roles.id"), nullable=False, index=True)
    permission_id = db.Column(db.String(36), db.ForeignKey("permissions.id"), nullable=False, index=True)
    scope_type = db.Column(db.String(20), nullable=False, default="chapter")  # chapter/project/team/assigned/own

    role = db.relationship("Role", back_populates="permissions")
    permission = db.relationship("Permission")


class Membership(OpsBase, db.Model):
    __tablename__ = "memberships"
    __table_args__ = (db.UniqueConstraint("organization_id", "user_id", name="uq_membership_org_user"),)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    role_id = db.Column(db.String(36), db.ForeignKey("roles.id"), nullable=False, index=True)
    member_status = db.Column(db.String(20), nullable=False, default="active")  # active/invited/suspended
    joined_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    title = db.Column(db.String(120), nullable=True)
    settings_json = db.Column(db.Text, nullable=False, default="{}")

    user = db.relationship("User", foreign_keys=[user_id])
    role = db.relationship("Role", foreign_keys=[role_id])
    organization = db.relationship("Organization")

    @property
    def settings(self) -> dict:
        try:
            value = json.loads(self.settings_json or "{}")
        except Exception:
            return {}
        return value if isinstance(value, dict) else {}

    @settings.setter
    def settings(self, value):
        self.settings_json = json.dumps(dict(value or {}), default=str)


class Team(OpsBase, db.Model):
    __tablename__ = "teams"
    __table_args__ = (db.UniqueConstraint("organization_id", "name", name="uq_teams_org_name"),)
    name = db.Column(db.String(160), nullable=False)
    description = db.Column(db.Text, nullable=True)
    parent_team_id = db.Column(db.String(36), db.ForeignKey("teams.id"), nullable=True, index=True)
    project_id = db.Column(db.String(36), db.ForeignKey("ops_projects.id"), nullable=True, index=True)
    color = db.Column(db.String(20), nullable=True)
    escalation_note = db.Column(db.String(500), nullable=True)
    archived_at = db.Column(db.DateTime, nullable=True)

    members = db.relationship("TeamMember", back_populates="team", cascade="all, delete-orphan")
    parent = db.relationship("Team", remote_side="Team.id", foreign_keys=[parent_team_id])
    project = db.relationship("OpsProject", foreign_keys=[project_id])

    @property
    def lead_user_ids(self) -> set[int]:
        return {m.user_id for m in self.members if m.is_lead}

    @property
    def member_user_ids(self) -> set[int]:
        return {m.user_id for m in self.members}


class TeamMember(db.Model):
    __tablename__ = "team_members"
    __table_args__ = (db.UniqueConstraint("team_id", "user_id", name="uq_team_member"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    team_id = db.Column(db.String(36), db.ForeignKey("teams.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    is_lead = db.Column(db.Boolean, nullable=False, default=False)
    joined_at = db.Column(db.DateTime, default=utcnow, nullable=False)

    team = db.relationship("Team", back_populates="members")
    user = db.relationship("User")


class Location(OpsBase, db.Model):
    __tablename__ = "locations"
    name = db.Column(db.String(160), nullable=False)
    description = db.Column(db.Text, nullable=True)
    parent_location_id = db.Column(db.String(36), db.ForeignKey("locations.id"), nullable=True, index=True)
    building = db.Column(db.String(160), nullable=True)
    room = db.Column(db.String(80), nullable=True)
    address_json = db.Column(db.Text, nullable=True)
    is_default = db.Column(db.Boolean, nullable=False, default=False)
    qr_code = db.Column(db.String(120), nullable=True)
    archived_at = db.Column(db.DateTime, nullable=True)

    parent = db.relationship("Location", remote_side="Location.id", foreign_keys=[parent_location_id])


class Category(OpsBase, db.Model):
    __tablename__ = "categories"
    __table_args__ = (db.UniqueConstraint("organization_id", "name", name="uq_categories_org_name"),)
    name = db.Column(db.String(120), nullable=False)
    color = db.Column(db.String(20), nullable=False, default="#0878d1")
    icon = db.Column(db.String(60), nullable=False, default="tag")
    description = db.Column(db.Text, nullable=True)
    archived_at = db.Column(db.DateTime, nullable=True)


class Comment(OpsBase, db.Model):
    __tablename__ = "comments"
    __table_args__ = (db.Index("ix_comments_entity", "entity_type", "entity_id"),)
    entity_type = db.Column(db.String(40), nullable=False)
    entity_id = db.Column(db.String(36), nullable=False)
    author_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    body = db.Column(db.Text, nullable=False)
    parent_comment_id = db.Column(db.String(36), db.ForeignKey("comments.id"), nullable=True)
    edited_at = db.Column(db.DateTime, nullable=True)
    deleted_at = db.Column(db.DateTime, nullable=True)

    author = db.relationship("User", foreign_keys=[author_user_id])


class Attachment(OpsBase, db.Model):
    __tablename__ = "attachments"
    __table_args__ = (db.Index("ix_attachments_entity", "entity_type", "entity_id"),)
    storage_key = db.Column(db.String(300), nullable=False, unique=True)
    original_name = db.Column(db.String(260), nullable=False)
    content_type = db.Column(db.String(120), nullable=False)
    size_bytes = db.Column(db.Integer, nullable=False)
    sha256 = db.Column(db.String(64), nullable=True)
    kind = db.Column(db.String(20), nullable=False, default="file")  # file / image
    uploaded_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    entity_type = db.Column(db.String(40), nullable=False)
    entity_id = db.Column(db.String(36), nullable=False)
    scan_status = db.Column(db.String(20), nullable=False, default="skipped")  # skipped / clean / infected
    deleted_at = db.Column(db.DateTime, nullable=True)

    uploaded_by = db.relationship("User", foreign_keys=[uploaded_by_user_id])


class Notification(OpsBase, db.Model):
    __tablename__ = "notifications"
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    type = db.Column(db.String(60), nullable=False)
    title = db.Column(db.String(220), nullable=False)
    body = db.Column(db.Text, nullable=True)
    entity_type = db.Column(db.String(40), nullable=True)
    entity_id = db.Column(db.String(36), nullable=True)
    read_at = db.Column(db.DateTime, nullable=True)

    user = db.relationship("User", foreign_keys=[user_id])


class SavedFilter(OpsBase, db.Model):
    __tablename__ = "saved_filters"
    entity_type = db.Column(db.String(40), nullable=False, index=True)
    owner_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    name = db.Column(db.String(120), nullable=False)
    visibility = db.Column(db.String(20), nullable=False, default="private")  # private / team / chapter
    team_id = db.Column(db.String(36), db.ForeignKey("teams.id"), nullable=True)
    filter_json = db.Column(db.Text, nullable=False, default="{}")
    sort_json = db.Column(db.Text, nullable=False, default="{}")
    view_type = db.Column(db.String(30), nullable=False, default="panel")
    is_default = db.Column(db.Boolean, nullable=False, default=False)

    owner = db.relationship("User", foreign_keys=[owner_user_id])


class AuditEvent(db.Model):
    """Immutable. Written inside the transaction of the action it records."""

    __tablename__ = "audit_events"
    __table_args__ = (db.Index("ix_audit_events_entity", "entity_type", "entity_id"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    organization_id = db.Column(db.String(36), db.ForeignKey("organizations.id"), nullable=False, index=True)
    actor_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    event_type = db.Column(db.String(80), nullable=False, index=True)
    entity_type = db.Column(db.String(40), nullable=False)
    entity_id = db.Column(db.String(36), nullable=False)
    before_json = db.Column(db.Text, nullable=True)
    after_json = db.Column(db.Text, nullable=True)
    metadata_json = db.Column(db.Text, nullable=False, default="{}")
    request_id = db.Column(db.String(64), nullable=True)
    ip_address = db.Column(db.String(120), nullable=True)
    occurred_at = db.Column(db.DateTime, default=utcnow, nullable=False, index=True)

    actor = db.relationship("User", foreign_keys=[actor_user_id])
