"""In-memory login rate limiting keyed by (client ip, identifier).

Per-process by design: it exists to slow credential stuffing on a small
instance, not to be a distributed limiter.
"""

from __future__ import annotations

import time
from threading import Lock


class LoginRateLimiter:
    def __init__(self, window_seconds: int, max_attempts: int):
        self.window_seconds = max(1, int(window_seconds))
        self.max_attempts = max(1, int(max_attempts))
        self._rows: dict[str, dict] = {}
        self._lock = Lock()

    @staticmethod
    def key(client_ip: str, identifier: str) -> str:
        return f"{client_ip}|{(identifier or '').strip().lower()}"

    def _sweep(self, now: float) -> None:
        expired = [key for key, info in self._rows.items() if now > info.get("reset_at", 0)]
        for key in expired:
            self._rows.pop(key, None)

    def record_failure(self, client_ip: str, identifier: str) -> None:
        now = time.time()
        with self._lock:
            self._sweep(now)
            key = self.key(client_ip, identifier)
            row = self._rows.get(key)
            if not row or now > row.get("reset_at", 0):
                self._rows[key] = {"count": 1, "reset_at": now + self.window_seconds}
                return
            row["count"] = int(row.get("count", 0)) + 1

    def clear(self, client_ip: str, identifier: str) -> None:
        with self._lock:
            self._rows.pop(self.key(client_ip, identifier), None)

    def is_limited(self, client_ip: str, identifier: str) -> tuple[bool, int]:
        now = time.time()
        with self._lock:
            self._sweep(now)
            row = self._rows.get(self.key(client_ip, identifier))
            if not row:
                return False, 0
            if int(row.get("count", 0)) < self.max_attempts:
                return False, 0
            retry_after = int(max(1, row.get("reset_at", 0) - now))
            return True, retry_after

    def reset_all(self) -> None:
        with self._lock:
            self._rows.clear()
