"""JSON-backed session and metadata stores (atomic, restart-safe)."""
from __future__ import annotations

import time
import uuid
from pathlib import Path

from app.core.exceptions import WorkbookNotFound
from app.storage.filesystem import FileSystem


class SessionStore:
    """One session.json per user: active workbook, active sheet, UI state.

    All methods are synchronous and free of awaits, so the read-modify-write
    in set() cannot interleave within one asyncio event loop.
    """

    def __init__(self, fs: FileSystem) -> None:
        self.fs = fs

    def _path(self, user_id: int) -> Path:
        return self.fs.user_dir(user_id) / "session.json"

    def get(self, user_id: int) -> dict:
        return self.fs.read_json(
            self._path(user_id),
            {"active_workbook_id": None, "active_sheet": None, "ui_state": {}, "updated_at": None},
        )

    def set(self, user_id: int, **fields) -> dict:
        session = self.get(user_id)
        session.update(fields)
        session["updated_at"] = time.time()
        self.fs.atomic_write_json(self._path(user_id), session)
        return session

    def reset(self, user_id: int) -> None:
        self.set(user_id, active_workbook_id=None, active_sheet=None, ui_state={})


class MetadataStore:
    """metadata.json inside each workbook directory + the user.json registry."""

    def __init__(self, fs: FileSystem) -> None:
        self.fs = fs

    # --- per-workbook -------------------------------------------------------
    def _meta_path(self, user_id: int, workbook_id: str) -> Path:
        return self.fs.workbook_dir(user_id, workbook_id) / "metadata.json"

    def create(self, user_id: int, original_name: str, size: int, warnings: list[str]) -> dict:
        workbook_id = str(uuid.uuid4())
        self.fs.workbook_dir(user_id, workbook_id, create=True)
        (self.fs.workbook_dir(user_id, workbook_id) / "versions").mkdir(exist_ok=True)
        meta = {
            "id": workbook_id,
            "owner_id": user_id,
            "display_name": original_name,
            "original_name": original_name,
            "size_bytes": size,
            "sheets": [],
            "warnings": warnings,
            "feature_flags": {},
            "version_counter": 0,
            "created_at": time.time(),
            "updated_at": time.time(),
        }
        self.fs.atomic_write_json(self._meta_path(user_id, workbook_id), meta)
        return meta

    def get(self, user_id: int, workbook_id: str) -> dict:
        # Existence is checked here (not inside read_json) so a missing file
        # raises WorkbookNotFound, not StorageError — callers depend on this.
        path = self._meta_path(user_id, workbook_id)
        if not path.exists():
            raise WorkbookNotFound("فایل موردنظر یافت نشد.")
        meta = self.fs.read_json(path)
        if meta.get("owner_id") != user_id:
            # Ownership is enforced here AND at every service call. Do not
            # reveal that the file exists to a non-owner.
            raise WorkbookNotFound("فایل موردنظر یافت نشد.")
        return meta

    def update(self, user_id: int, workbook_id: str, **fields) -> dict:
        meta = self.get(user_id, workbook_id)
        meta.update(fields)
        meta["updated_at"] = time.time()
        self.fs.atomic_write_json(self._meta_path(user_id, workbook_id), meta)
        return meta

    def delete(self, user_id: int, workbook_id: str) -> None:
        self.get(user_id, workbook_id)  # enforce ownership
        self.fs.delete_workbook(user_id, workbook_id)

    # --- registry -------------------------------------------------------
    def _registry_path(self, user_id: int) -> Path:
        return self.fs.user_dir(user_id) / "user.json"

    def list_workbooks(self, user_id: int) -> list[dict]:
        reg = self.fs.read_json(self._registry_path(user_id), {"workbooks": []})
        result = []
        for wid in list(reg.get("workbooks", [])):
            try:
                result.append(self.get(user_id, wid))
            except WorkbookNotFound:
                continue  # stale registry entry, skip silently (logged upstream)
        return result

    def register(self, user_id: int, workbook_id: str) -> None:
        reg = self.fs.read_json(self._registry_path(user_id), {"workbooks": []})
        if workbook_id not in reg["workbooks"]:
            reg["workbooks"].append(workbook_id)
            self.fs.atomic_write_json(self._registry_path(user_id), reg)

    def unregister(self, user_id: int, workbook_id: str) -> None:
        reg = self.fs.read_json(self._registry_path(user_id), {"workbooks": []})
        reg["workbooks"] = [w for w in reg["workbooks"] if w != workbook_id]
        self.fs.atomic_write_json(self._registry_path(user_id), reg)