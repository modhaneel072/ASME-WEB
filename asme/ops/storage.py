"""Private file storage behind an adapter interface.

``LocalStorage`` keeps files under ``instance/private_files`` (outside the static
folder) and every download goes through an HMAC-signed, expiring token. An S3
adapter implements the same three methods when object storage is configured.
"""

from __future__ import annotations

import hashlib
import os
import shutil
from pathlib import Path

from flask import current_app
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

ALLOWED_CONTENT_TYPES = {
    "image/png": "image",
    "image/jpeg": "image",
    "image/webp": "image",
    "image/gif": "image",
    "application/pdf": "file",
    "text/plain": "file",
    "text/csv": "file",
    "application/json": "file",
    "application/zip": "file",
    "application/octet-stream": "file",
    "model/stl": "file",
    "application/sla": "file",
    "application/vnd.ms-excel": "file",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "file",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "file",
}
ALLOWED_EXTENSIONS = {
    "png", "jpg", "jpeg", "webp", "gif", "pdf", "txt", "csv", "json", "zip", "stl", "step", "stp", "gcode", "3mf",
    "xlsx", "xls", "docx", "md",
}


class StorageError(ValueError):
    pass


class Storage:
    name = "abstract"

    def save(self, key: str, stream) -> tuple[int, str]:
        raise NotImplementedError

    def open(self, key: str):
        raise NotImplementedError

    def delete(self, key: str) -> None:
        raise NotImplementedError


class LocalStorage(Storage):
    name = "local"

    def __init__(self, root: Path):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)

    def _path(self, key: str) -> Path:
        path = (self.root / key).resolve()
        if self.root.resolve() not in path.parents:
            raise StorageError("invalid storage key")
        return path

    def save(self, key: str, stream) -> tuple[int, str]:
        path = self._path(key)
        path.parent.mkdir(parents=True, exist_ok=True)
        digest = hashlib.sha256()
        size = 0
        with open(path, "wb") as handle:
            while True:
                chunk = stream.read(1024 * 256)
                if not chunk:
                    break
                handle.write(chunk)
                digest.update(chunk)
                size += len(chunk)
        return size, digest.hexdigest()

    def open(self, key: str):
        return open(self._path(key), "rb")

    def delete(self, key: str) -> None:
        path = self._path(key)
        if path.exists():
            os.remove(path)


def get_storage() -> Storage:
    storage = current_app.extensions.get("asme_storage")
    if storage is None:
        root = Path(current_app.instance_path) / "private_files"
        storage = LocalStorage(root)
        current_app.extensions["asme_storage"] = storage
    return storage


# --------------------------------------------------------------------------- signed access


def _serializer() -> URLSafeTimedSerializer:
    return URLSafeTimedSerializer(current_app.config["SECRET_KEY"], salt="asme-ops-file")


def sign_download(attachment_id: str) -> str:
    return _serializer().dumps({"a": attachment_id})


def verify_download(token: str, max_age: int) -> str | None:
    try:
        payload = _serializer().loads(token, max_age=max_age)
    except (BadSignature, SignatureExpired):
        return None
    return payload.get("a") if isinstance(payload, dict) else None


def classify(content_type: str, filename: str) -> str:
    """Return 'image' or 'file'; raise ``StorageError`` when the type is not allowed."""
    ext = (filename.rsplit(".", 1)[-1].lower() if "." in filename else "")
    ctype = (content_type or "").split(";")[0].strip().lower()
    if ext and ext not in ALLOWED_EXTENSIONS:
        raise StorageError(f".{ext} files are not allowed.")
    if ctype in ALLOWED_CONTENT_TYPES:
        return ALLOWED_CONTENT_TYPES[ctype]
    if ext in {"png", "jpg", "jpeg", "webp", "gif"}:
        return "image"
    if ext in ALLOWED_EXTENSIONS:
        return "file"
    raise StorageError("That file type is not allowed.")


def scan_file(path_or_stream) -> str:
    """Virus-scan hook. Returns 'clean', 'infected' or 'skipped'.

    No scanner is wired in this environment; the status is recorded on the attachment
    so an admin can see that scanning did not run.
    """
    return "skipped"


def remove_tree(root: Path):  # test helper
    shutil.rmtree(root, ignore_errors=True)
