"""String normalisation helpers shared by identity, inventory and roster code."""

from __future__ import annotations

import csv
import io
import re
import unicodedata


def clean_tag_value(raw):
    value = (raw or "").strip()
    if not value:
        return ""
    return " ".join(value.split())


def normalize_text_key(raw):
    normalized = unicodedata.normalize("NFKD", str(raw or ""))
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "", ascii_only.lower())


def split_name_parts(full_name):
    cleaned = " ".join((full_name or "").strip().split())
    if not cleaned:
        return "", ""
    tokens = cleaned.split(" ")
    first = tokens[0]
    last = " ".join(tokens[1:]) if len(tokens) > 1 else tokens[0]
    return first, last


def slugify(text_value):
    raw = (text_value or "").strip().lower()
    if not raw:
        return ""
    cleaned = []
    prev_dash = False
    for char in raw:
        if char.isalnum():
            cleaned.append(char)
            prev_dash = False
        elif not prev_dash:
            cleaned.append("-")
            prev_dash = True
    slug = "".join(cleaned).strip("-")
    return slug[:150]


def csv_stream_from_rows(headers, rows):
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(headers)
    for row in rows:
        writer.writerow(row)
    return buffer.getvalue()


def truncate(value, limit):
    text = (value or "").strip()
    return text[:limit] if text else None
