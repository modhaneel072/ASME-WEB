"""Attachments: upload (multipart), list, signed download, delete."""

from __future__ import annotations

from flask import request, send_file

from asme.config import settings
from asme.ops import authz, storage
from asme.ops.api import bp, ctx, ok
from asme.ops.models import Asset, Attachment, OpsProject
from asme.ops.services import files
from asme.ops.services import work_orders as wo_service
from asme.ops.tenancy import get_or_404
from asme.services.errors import NotFound, Validation

ENTITY_TYPES = {"work_order", "project", "asset"}


def _entity(entity_type: str, entity_id: str):
    if entity_type == "work_order":
        return wo_service.get_visible_or_404(ctx(), entity_id)
    if entity_type == "project":
        return get_or_404(OpsProject, entity_id, ctx(), "Project")
    if entity_type == "asset":
        return get_or_404(Asset, entity_id, ctx(), "Asset")
    raise Validation("Unknown entity type.", field="entity_type")


@bp.get("/files")
def files_list():
    entity_type = (request.args.get("entity_type") or "").strip()
    entity_id = (request.args.get("entity_id") or "").strip()
    if entity_type not in ENTITY_TYPES or not entity_id:
        raise Validation("entity_type and entity_id are required.")
    _entity(entity_type, entity_id)
    return ok([files.serialize(f) for f in files.list_for(ctx(), entity_type, entity_id)])


@bp.post("/files")
def files_upload():
    """Multipart upload: ``file``, ``entity_type``, ``entity_id``."""
    entity_type = (request.form.get("entity_type") or "").strip()
    entity_id = (request.form.get("entity_id") or "").strip()
    if entity_type not in ENTITY_TYPES or not entity_id:
        raise Validation("entity_type and entity_id are required.")
    record = _entity(entity_type, entity_id)
    authz.require(ctx(), "file.upload", record=record)
    attachment = files.upload(ctx(), entity_type, entity_id, request.files.get("file"))
    from asme.extensions import db

    db.session.commit()
    return ok(files.serialize(attachment), status=201)


@bp.get("/files/<attachment_id>")
def files_get(attachment_id):
    row = Attachment.query.filter_by(id=attachment_id, organization_id=ctx().organization.id, deleted_at=None).first()
    if row is None:
        raise NotFound("File not found.")
    _entity(row.entity_type, row.entity_id)
    return ok(files.serialize(row))


@bp.get("/files/<attachment_id>/download")
def file_download(attachment_id):
    """Token-authenticated download; the token expires after ``ASME_FILE_URL_TTL_SECONDS``."""
    token = (request.args.get("token") or "").strip()
    verified = storage.verify_download(token, settings().file_url_ttl_seconds) if token else None
    if verified != attachment_id:
        return {"ok": False, "code": "link_expired", "error": "This download link is invalid or has expired. Reload the page to get a fresh one."}, 403
    row = Attachment.query.filter_by(id=attachment_id, deleted_at=None).first()
    if row is None:
        raise NotFound("File not found.")
    handle = storage.get_storage().open(row.storage_key)
    return send_file(handle, mimetype=row.content_type, as_attachment=(row.kind != "image"), download_name=row.original_name, max_age=0)


@bp.delete("/files/<attachment_id>")
def files_delete(attachment_id):
    row = Attachment.query.filter_by(id=attachment_id, organization_id=ctx().organization.id, deleted_at=None).first()
    if row is None:
        raise NotFound("File not found.")
    _entity(row.entity_type, row.entity_id)
    files.remove(ctx(), row)
    return ok({"deleted": True})
