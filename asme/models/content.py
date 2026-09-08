from datetime import datetime

from asme.extensions import db


class Project(db.Model):
    __tablename__ = "projects"
    id = db.Column(db.Integer, primary_key=True)

    slug = db.Column(db.String(160), nullable=False, unique=True, index=True)
    title = db.Column(db.String(200), nullable=False)
    project_type = db.Column(db.String(80), nullable=True)
    summary = db.Column(db.String(320), nullable=False)
    description = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(80), nullable=False, default="Active")
    timeline = db.Column(db.Text, nullable=True)
    gallery_json = db.Column(db.Text, nullable=True)
    lead_name = db.Column(db.String(160), nullable=True)
    image_url = db.Column(db.String(500), nullable=True)
    external_link = db.Column(db.String(500), nullable=True)
    is_joinable = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    memberships = db.relationship("ProjectMembership", back_populates="project")


class ProjectMembership(db.Model):
    """A user on a project team. Feeds the Launchpad "join a team" task."""

    __tablename__ = "project_memberships"
    __table_args__ = (db.UniqueConstraint("project_id", "user_id", name="uq_project_membership"),)
    id = db.Column(db.Integer, primary_key=True)

    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    role = db.Column(db.String(40), nullable=False, default="member")  # member / lead
    joined_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    left_at = db.Column(db.DateTime, nullable=True)

    project = db.relationship("Project", back_populates="memberships")
    user = db.relationship("User")


class WorkLog(db.Model):
    """Hours a member logged against a project. Feeds the Launchpad contributor task."""

    __tablename__ = "work_logs"
    id = db.Column(db.Integer, primary_key=True)

    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=True, index=True)
    hours = db.Column(db.Float, nullable=False)
    note = db.Column(db.String(300), nullable=True)
    logged_for = db.Column(db.Date, nullable=False)
    approved_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    approved_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    user = db.relationship("User", foreign_keys=[user_id])
    project = db.relationship("Project")
    approved_by = db.relationship("User", foreign_keys=[approved_by_user_id])


class ContactMessage(db.Model):
    __tablename__ = "contact_messages"
    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(160), nullable=False)
    email = db.Column(db.String(160), nullable=False)
    kind = db.Column(db.String(40), nullable=False, default="contact")
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    target = db.Column(db.String(80), nullable=True)
    subject = db.Column(db.String(220), nullable=True)
    message = db.Column(db.Text, nullable=False)
    admin_reply = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(40), nullable=False, default="new")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    user = db.relationship("User", foreign_keys=[user_id])


class Announcement(db.Model):
    __tablename__ = "announcements"
    id = db.Column(db.Integer, primary_key=True)

    title = db.Column(db.String(220), nullable=False)
    body = db.Column(db.Text, nullable=False)
    is_published = db.Column(db.Boolean, nullable=False, default=True)
    show_on_public = db.Column(db.Boolean, nullable=False, default=True)
    show_on_member = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    published_at = db.Column(db.DateTime, nullable=True)
