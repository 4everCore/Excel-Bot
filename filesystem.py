"""Filesystem storage layer: user isolation, atomic writes, quotas.

No database. JSON + xlsx files only. All persistence is atomic
(write-to-temp, fsync, os.replace) so crashes never corrupt state.
"""
from __future__ import annotations

import json
import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any

from app.core.exceptions import QuotaExceeded, StorageError, WorkbookNotFound

_MISSING = object()


class FileSystem:
    def __init__(self, root: Path) -> None:
        self.root = root
        for sub in ("users", "temp", "logs", "system"):
            (root / sub).mkdir(parents=True, exist_ok=True)

    # --- paths -------------------------------------------------------------
    def user_dir(self, user_id: int) -> Path:
        if user_id <= 0:
            raise StorageError("شناسهٔ کاربر نامعتبر است.")
        d = self.root / "users" / str(user_id)
        d.mkdir(parents=True, exist_ok=True)
        return d

    def workbooks_dir(self, user_id: int) -> Path:
        d = self.user_dir(user_id) / "workbooks"
        d.mkdir(exist_ok=True)
        return d

    def workbook_dir(self, user_id: int, workbook_id: str, create: bool = False) -> Path:
        """Strictly validate the workbook id before touching the filesystem.

        This is the path-traversal defense: only canonical UUID strings are
        accepted, so user-supplied ids can never escape the user directory.
        """
        try:
            workbook_id = str(uuid.UUID(workbook_id))
        except (ValueError, AttributeError, TypeError) as exc:
            raise WorkbookNotFound("شناسهٔ فایل نامعتبر است.") from exc
        d = self.workbooks_dir(user_id) / workbook_id
        if create:
            d.mkdir(parents=True, exist_ok=True)
        elif not d.is_dir():
            raise WorkbookNotFound("فایل موردنظر یافت نشد.")
        return d

    # --- atomic primitives ---------------------------------------------------
    @staticmethod
    def atomic_write_bytes(path: Path, data: bytes) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp_name = tempfile.mkstemp(dir=path.parent, prefix=".tmp-")
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(data)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_name, path)
        except OSError as exc:
            Path(tmp_name).unlink(missing_ok=True)
            raise StorageError(f"نوشتن فایل ناموفق بود: {exc}") from exc

    def atomic_write_json(self, path: Path, obj: dict) -> None:
        payload = json.dumps(obj, ensure_ascii=False, indent=2, default=str)
        self.atomic_write_bytes(path, payload.encode("utf-8"))

    @staticmethod
    def read_json(path: Path, default: Any = _MISSING) -> Any:
        """Read JSON. Pass `default` explicitly (even None) to tolerate absence."""
        if not path.exists():
            if default is not _MISSING:
                return default
            raise StorageError(f"فایل حالت یافت نشد: {path.name}")
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as exc:
            if default is not _MISSING:
                return default
            raise StorageError(f"خواندن فایل حالت ناموفق بود: {exc}") from exc

    # --- quotas ---------------------------------------------------------------
    def dir_size(self, path: Path) -> int:
        return sum(f.stat().st_size for f in path.rglob("*") if f.is_file())

    def check_user_quota(self, user_id: int, new_bytes: int, settings) -> None:
        wb_dir = self.workbooks_dir(user_id)
        existing = [d for d in wb_dir.iterdir() if d.is_dir()]
        if len(existing) >= settings.max_workbooks_per_user:
            raise QuotaExceeded(
                f"حداکثر {settings.max_workbooks_per_user} فایل همزمان مجاز است؛ "
                "ابتدا یکی از فایل‌ها را حذف کنید."
            )
        used = sum(self.dir_size(d) for d in existing)
        if used + new_bytes > settings.max_storage_per_user_bytes:
            raise QuotaExceeded("سهمیهٔ فضای ذخیره‌سازی شما پر شده است.")

    # --- temp & cleanup ---------------------------------------------------
    def temp_dir(self) -> Path:
        d = self.root / "temp"
        d.mkdir(exist_ok=True)
        return d

    def cleanup_temp(self, older_than_hours: float = 6) -> int:
        cutoff = time.time() - older_than_hours * 3600
        removed = 0
        for f in self.temp_dir().iterdir():
            try:
                if f.is_file() and f.stat().st_mtime < cutoff:
                    f.unlink()
                    removed += 1
            except OSError:
                continue
        return removed

    def delete_workbook(self, user_id: int, workbook_id: str) -> None:
        d = self.workbook_dir(user_id, workbook_id)
        shutil.rmtree(d, ignore_errors=False)