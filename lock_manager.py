"""Per-workbook and per-user file locks.

Excel operations run synchronously inside worker threads (never on the event
loop), so a synchronous filelock acquired inside that thread is the correct
primitive: it serializes conflicting edits to the same workbook across
processes AND threads, while different users never contend.
"""
from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

from filelock import FileLock, Timeout

from app.core.exceptions import LockConflict

_SAFE_KEY_RE = __import__("re").compile(r"^[A-Za-z0-9_-]+$")


class LockManager:
    def __init__(self, root: Path) -> None:
        self._locks_dir = root / "system" / "locks"
        self._locks_dir.mkdir(parents=True, exist_ok=True)
        self._locks: dict[str, FileLock] = {}

    def _lock_for(self, key: str) -> FileLock:
        if not _SAFE_KEY_RE.match(key):
            raise LockConflict("شناسهٔ قفل نامعتبر است.")
        if key not in self._locks:
            self._locks[key] = FileLock(str(self._locks_dir / f"{key}.lock"), timeout=10)
        return self._locks[key]

    @contextmanager
    def acquire_key(self, key: str, timeout: float = 10.0):
        lock = self._lock_for(key)
        lock._timeout = timeout
        try:
            with lock:
                yield
        except Timeout as exc:
            raise LockConflict(
                "عملیات دیگری روی همین فایل در حال اجراست؛ چند لحظه بعد دوباره تلاش کنید."
            ) from exc

    @contextmanager
    def acquire(self, user_id: int, workbook_id: str, timeout: float = 10.0):
        with self.acquire_key(f"{user_id}-{workbook_id}", timeout=timeout):
            yield