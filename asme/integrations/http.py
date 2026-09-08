"""Minimal JSON-over-HTTP client with retry, used by the Outlook adapter."""

from __future__ import annotations

import json
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


def http_json_request(
    method,
    url,
    headers=None,
    form_data=None,
    json_data=None,
    timeout=30,
    retries=0,
    retry_backoff=0.8,
):
    body = None
    request_headers = dict(headers or {})

    if form_data is not None:
        body = urlencode(form_data).encode("utf-8")
        request_headers["Content-Type"] = "application/x-www-form-urlencoded"
    elif json_data is not None:
        body = json.dumps(json_data).encode("utf-8")
        request_headers["Content-Type"] = "application/json"

    req = Request(url=url, data=body, method=method)
    for key, value in request_headers.items():
        req.add_header(key, value)

    for attempt in range(retries + 1):
        try:
            with urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8") if resp else ""
                payload = json.loads(raw) if raw else {}
                return resp.status, payload, None
        except HTTPError as exc:
            error_body = exc.read().decode("utf-8", errors="ignore")
            try:
                payload = json.loads(error_body) if error_body else {}
            except Exception:
                payload = {"raw": error_body}
            if exc.code in (429, 500, 502, 503, 504) and attempt < retries:
                time.sleep(retry_backoff * (2**attempt))
                continue
            return exc.code, payload, error_body
        except URLError as exc:
            if attempt < retries:
                time.sleep(retry_backoff * (2**attempt))
                continue
            return None, {}, str(exc)
        except Exception as exc:
            if attempt < retries:
                time.sleep(retry_backoff * (2**attempt))
                continue
            return None, {}, str(exc)

    return None, {}, "HTTP request failed"
