"""Operational projects (programmes) - distinct from the public-site ``projects`` rows,
which remain the published content projection and are linked by ``public_project_id``."""

from __future__ import annotations

from asme.extensions import db
from asme.ops.models.base import OpsBase, new_uuid, utcnow

PROJECT_STATUSES = ("planning", "active", "on_hold", "completed", "archived")
RISK_LEVELS = ("low", "medium", "high")
MILESTONE_STATUSES = ("planned", "in_progress", "complete", "missed")
PROJECT_ROLES = ("lead", "member", "advisor", "observer")


class OpsProject(OpsBase, db.Model):
    __tablename__ = "ops_projects"
    __table_args__ = (db.UniqueConstraint("organization_id", "code", name="uq_ops_projects_org_code"),)
    name = db.Column(db.String(200), nullable=False)
    code = db.Column(db.String(30), nullable=False)
    description = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), nullable=False, default="active", index=True)
    risk_level = db.Column(db.String(20), nullable=False, default="medium")
    lead_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    faculty_advisor_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    competition = db.Column(db.String(200), nullable=True)
    academic_year = db.Column(db.String(9), nullable=True)
    start_date = db.Column(db.Date, nullable=True)
    target_date = db.Column(db.Date, nullable=True)
    budget_amount = db.Column(db.Numeric(12, 2), nullable=True)
    budget_code = db.Column(db.String(60), nullable=True)
    repository_url = db.Column(db.String(500), nullable=True)
    cad_url = db.Column(db.String(500), nullable=True)
    requirements_url = db.Column(db.String(500), nullable=True)
    public_project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=True, unique=True)
    visibility = db.Column(db.String(20), nullable=False, default="members")  # private / members / public
    archived_at = db.Column(db.DateTime, nullable=True)

    lead = db.relationship("User", foreign_keys=[lead_user_id])
    faculty_advisor = db.relationship("User", foreign_keys=[faculty_advisor_user_id])
    public_project = db.relationship("Project", foreign_keys=[public_project_id])
    members = db.relationship("OpsProjectMember", back_populates="project", cascade="all, delete-orphan")
    milestones = db.relationship("Milestone", back_populates="project", cascade="all, delete-orphan", order_by="Milestone.order")

    @property
    def member_user_ids(self) -> set[int]:
        ids = {m.user_id for m in self.members}
        if self.lead_user_id:
            ids.add(self.lead_user_id)
        return ids


class OpsProjectMember(db.Model):
    __tablename__ = "ops_project_members"
    __table_args__ = (db.UniqueConstraint("project_id", "user_id", name="uq_ops_project_member"),)
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)
    project_id = db.Column(db.String(36), db.ForeignKey("ops_projects.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    team_id = db.Column(db.String(36), db.ForeignKey("teams.id"), nullable=True)
    project_role = db.Column(db.String(20), nullable=False, default="member")
    joined_at = db.Column(db.DateTime, default=utcnow, nullable=False)

    project = db.relationship("OpsProject", back_populates="members")
    user = db.relationship("User")
    team = db.relationship("Team", foreign_keys=[team_id])


class Milestone(OpsBase, db.Model):
    __tablename__ = "milestones"
    project_id = db.Column(db.String(36), db.ForeignKey("ops_projects.id"), nullable=False, index=True)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    due_date = db.Column(db.Date, nullable=True)
    status = db.Column(db.String(20), nullable=False, default="planned")
    owner_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    weight = db.Column(db.Integer, nullable=False, default=1)
    order = db.Column(db.Integer, nullable=False, default=0)
    completed_at = db.Column(db.DateTime, nullable=True)

    project = db.relationship("OpsProject", back_populates="milestones")
    owner = db.relationship("User", foreign_keys=[owner_user_id])
