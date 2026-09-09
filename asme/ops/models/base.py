"""Shared columns for every tenant-owned ops table.

* UUID primary keys stored as 36-char strings (portable across SQLite and PostgreSQL).
* ``organization_id`` on every row - the tenancy helpers filter on it for every query.
* created/updated stamps and actor ids for the audit trail.
"""

from __future__ import annotations

import uuid
from datetime import datetime

from sqlalchemy.ext.declarative import declared_attr

from asme.extensions import db


def new_uuid() -> str:
    return str(uuid.uuid4())


def utcnow() -> datetime:
    return datetime.utcnow()


class OpsBase:
    id = db.Column(db.String(36), primary_key=True, default=new_uuid)

    @declared_attr
    def organization_id(cls):
        return db.Column(db.String(36), db.ForeignKey("organizations.id"), nullable=False, index=True)

    created_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=utcnow, onupdate=utcnow, nullable=False)

    @declared_attr
    def created_by_user_id(cls):
        return db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)

    @declared_attr
    def updated_by_user_id(cls):
        return db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
