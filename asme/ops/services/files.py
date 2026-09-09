"""Attachments: validated upload to private storage, signed expiring downloads."""

from __future__ import annotations

from datetime import datetime
from uuid import uuid4

from flask import url_for
from werkzeug.utils import secure_filename

from asme.config import settings
from asme.extensions import db
from asme.ops import audit, storage
from asme.ops.models import Attachment
from asme.ops.services.common import user_ref
from asme.services.errors import Forbidden, Validation


def download_url(attachment: Attachment) -> str:
    token = storage.sign_download(attachment.id)
    return url_for("ops_api.file_download", attachment_id=attachment.id, token=token)


def serialize(attachment: Attachment) -> dict:
    return {
        "id": attachment.id,
        "entity_type": attachment.entity_type,
        "entity_id": attachment.entity_id,
        "original_name": attachment.original_name,
        "content_type": attachment.content_type,
        "size_bytes": attachment.size_bytes,
        "kind": attachment.kind,
        "scan_status": attachment.scan_status,
        "uploaded_by": user_ref(attachment.uploaded_by),
        "created_at": attachment.created_at.isoformat(),
        "download_url": download_url(attachment),
        "expires_in_seconds": settings().file_url_ttl_seconds,
    }


def list_for(ctx, entity_type: str, entity_id: str):
    return (
        Attachment.query.filter(Attachment.organization_id == ctx.organization.id, Attachment.entity_type == entity_type, Attachment.entity_id == entity_id, Attachment.deleted_at.is_(None))
        .order_by(Attachment.created_at.desc())
        .all()
    )


def upload(ctx, entity_type: str, entity_id: str, file_storage) -> Attachment:
    if file_storage is None or not file_storage.filename:
        raise Validation("Choose a file to upload.", field="file")
    original = secure_filename(file_storage.filename)[:260] or "upload"
    try:
        kind = storage.classify(file_storage.content_type or "", original)
    except storage.StorageError as exc:
        raise Validation(str(exc), field="file", code="file_type")
    max_bytes = settings().file_max_mb * 1024 * 1024
    # Size check before writing anything.
    stream = file_storage.stream
    try:
        stream.seek(0, 2)
        size = stream.tell()
        stream.seek(0)
    except Exception:
        size = None
    if size is not None and size > max_bytes:
        raise Validation(f"File is larger than {settings().file_max_mb} MB.", field="file", code="file_size")
    key = f"{ctx.organization.id}/{entity_type}/{entity_id}/{uuid4().hex}_{original}"
    store = storage.get_storage()
    written, digest = store.save(key, stream)
    if written > max_bytes:
        store.delete(key)
        raise Validation(f"File is larger than {settings().file_max_mb} MB.", field="file", code="file_size")
    attachment = Attachment(
        organization_id=ctx.organization.id,
        storage_key=key,
        original_name=original,
        content_type=(file_storage.content_type or "application/octet-stream")[:120],
        size_bytes=written,
        sha256=digest,
        kind=kind,
        uploaded_by_user_id=ctx.user.id,
        entity_type=entity_type,
        entity_id=entity_id,
        scan_status=storage.scan_file(key),
        created_by_user_id=ctx.user.id,
        updated_by_user_id=ctx.user.id,
    )
    db.session.add(attachment)
    db.session.flush()
    audit.record_event("file.uploaded", entity_type, entity_id, organization_id=ctx.organization.id, actor=ctx.user, after={"attachment_id": attachment.id, "name": original, "size": written})
    return attachment


def remove(ctx, attachment: Attachment):
    if attachment.uploaded_by_user_id != ctx.user.id and not ctx.is_admin:
        raise Forbidden("You can only remove files you uploaded.", code="not_owner")
    attachment.deleted_at = datetime.utcnow()
    attachment.updated_by_user_id = ctx.user.id
    audit.record_event("file.deleted", attachment.entity_type, attachment.entity_id, organization_id=ctx.organization.id, actor=ctx.user, metadata={"attachment_id": attachment.id})
    db.session.commit()
    try:
        storage.get_storage().delete(attachment.storage_key)
    except Exception:
        pass
