"""List-query conventions: ``filter[key]``, ``sort``, ``page[limit]``, ``page[cursor]``, ``q``.

Every filter key and sort key is validated against an allowlist before it touches SQL.
"""

from __future__ import annotations

import base64
import json
from dataclasses import dataclass, field

from flask import request

from asme.services.errors import Validation


@dataclass
class ListQuery:
    filters: dict[str, list[str]] = field(default_factory=dict)
    sort: str = ""
    limit: int = 50
    cursor: dict | None = None
    q: str = ""
    extra: dict[str, str] = field(default_factory=dict)


def encode_cursor(payload: dict) -> str:
    raw = json.dumps(payload, default=str).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def decode_cursor(value: str | None) -> dict | None:
    if not value:
        return None
    try:
        padded = value + "=" * (-len(value) % 4)
        payload = json.loads(base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8"))
        if not isinstance(payload, dict):
            raise ValueError
        return payload
    except Exception:
        raise Validation("Invalid page cursor.", field="page[cursor]", code="bad_cursor")


def parse_list_query(
    *,
    allowed_filters: set[str],
    allowed_sorts: set[str],
    default_sort: str,
    default_limit: int = 50,
    max_limit: int = 200,
    allowed_extra: set[str] | None = None,
) -> ListQuery:
    query = ListQuery(sort=default_sort, limit=default_limit)
    for key, values in request.args.lists():
        if key.startswith("filter[") and key.endswith("]"):
            name = key[7:-1]
            if name not in allowed_filters:
                raise Validation(f"Unknown filter '{name}'.", field=key, code="unknown_filter")
            parsed: list[str] = []
            for value in values:
                parsed.extend(part.strip() for part in str(value).split(",") if part.strip())
            if parsed:
                query.filters[name] = parsed
        elif key == "sort":
            sort = (values[0] or "").strip()
            if sort and sort not in allowed_sorts:
                raise Validation(f"Unknown sort '{sort}'.", field="sort", code="unknown_sort")
            if sort:
                query.sort = sort
        elif key == "page[limit]":
            try:
                query.limit = max(1, min(int(values[0]), max_limit))
            except ValueError:
                raise Validation("page[limit] must be an integer.", field="page[limit]")
        elif key == "page[cursor]":
            query.cursor = decode_cursor(values[0])
        elif key == "q":
            query.q = (values[0] or "").strip()[:200]
        elif allowed_extra and key in allowed_extra:
            query.extra[key] = (values[0] or "").strip()
    return query
