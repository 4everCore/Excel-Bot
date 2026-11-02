"""Workbook service — the single entry point for every mutation.

Guarantees, for every operation:
  lock -> backup version -> mutate (worker thread) -> save to temp ->
  validate temp -> atomic replace -> metadata update -> lock release.
On any failure the previous valid workbook stays in place.

KNOWN BEHAVIOR (documented honestly):
- asyncio timeouts cancel the awaitable, not the worker thread. On timeout
  the operation may still complete in the background; the lock stays held
  until it does, so later operations receive LockConflict until then.
"""
from __future__ import annotations

import asyncio
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from openpyxl.workbook.workbook import Workbook

from app.core.config import Settings
from app.core.exceptions import (
    OperationTimeout,
    QuotaExceeded,
    UnsupportedFormat,
    WorkbookNotFound,
)
from app.excel.engine import (
    analyze_workbook,
    load_workbook_safe,
    validate_saved_workbook,
    WorkbookAnalysis,
)
from app.excel.operations.worksheets import sheet_summary
from app.storage.filesystem import FileSystem
from app.storage.lock_manager import LockManager
from app.storage.stores import MetadataStore, SessionStore
from app.core import security

EditOperation = Callable[[Workbook], Any]


@dataclass
class OperationResult:
    ok: bool
    summary: str
    data: Any = None
    version: int | None = None


class WorkbookService:
    def __init__(self, fs: FileSystem, settings: Settings) -> None:
        self.fs = fs
        self.settings = settings
        self.locks = LockManager(fs.root)
        self.sessions = SessionStore(fs)
        self.metadata = MetadataStore(fs)
        self._semaphore = asyncio.Semaphore(settings.max_concurrent_jobs)

    # ------------------------------------------------------------------ paths
    def _current_path(self, user_id: int, workbook_id: str) -> Path:
        return self.fs.workbook_dir(user_id, workbook_id) / "current.xlsx"

    def _original_path(self, user_id: int, workbook_id: str) -> Path:
        return self.fs.workbook_dir(user_id, workbook_id) / "original.xlsx"

    def _versions_dir(self, user_id: int, workbook_id: str) -> Path:
        return self.fs.workbook_dir(user_id, workbook_id) / "versions"

    # ------------------------------------------------------------------ upload
    async def upload(self, user_id: int, data: bytes, original_name: str) -> dict:
        """Validate, persist and register an uploaded workbook. Returns metadata."""
        report = await asyncio.to_thread(security.inspect_xlsx_package, data, original_name)
        if len(data) > self.settings.max_upload_size_bytes:
            raise UnsupportedFormat("حجم فایل بیش از حد مجاز است.")

        created_id: str | None = None

        def _persist() -> dict:
            nonlocal created_id
            # Per-user lock: quota check + dir creation must be atomic against
            # concurrent uploads by the same user.
            with self.locks.acquire_key(f"{user_id}-upload"):
                self.fs.check_user_quota(user_id, len(data), self.settings)
                meta = self.metadata.create(
                    user_id, original_name, len(data), list(report.warnings)
                )
                created_id = meta["id"]
                wb_dir = self.fs.workbook_dir(user_id, meta["id"])
                self.fs.atomic_write_bytes(wb_dir / "original.xlsx", data)
                self.fs.atomic_write_bytes(wb_dir / "current.xlsx", data)
            wb = load_workbook_safe(self._current_path(user_id, meta["id"]))
            analysis = analyze_workbook(wb)
            self.metadata.update(
                user_id,
                meta["id"],
                sheets=sheet_summary(wb),
                active_sheet=analysis.active_sheet,
                warnings=report.warnings + analysis.warnings,
                feature_flags=report.feature_flags,
            )
            self.metadata.register(user_id, meta["id"])
            self.sessions.set(user_id, active_workbook_id=meta["id"], active_sheet=analysis.active_sheet)
            return self.metadata.get(user_id, meta["id"])

        async with self._semaphore:
            try:
                return await asyncio.wait_for(asyncio.to_thread(_persist), timeout=60)
            except asyncio.TimeoutError as exc:
                raise OperationTimeout("ذخیرهٔ فایل بیش از حد طول کشید.") from exc
            except Exception:
                # No orphan directories that would still consume quota.
                if created_id:
                    shutil.rmtree(
                        self.fs.workbooks_dir(user_id) / created_id, ignore_errors=True
                    )
                raise

    # ------------------------------------------------------------------ reads
    def get_active_workbook_id(self, user_id: int) -> str:
        wid = self.sessions.get(user_id).get("active_workbook_id")
        if not wid:
            raise WorkbookNotFound("هیچ فایل فعالی انتخاب نشده است؛ ابتدا فایل را بارگذاری کنید.")
        self.metadata.get(user_id, wid)  # ownership check
        return wid

    def read_current(self, user_id: int, workbook_id: str) -> Workbook:
        self.metadata.get(user_id, workbook_id)  # ownership check
        return load_workbook_safe(self._current_path(user_id, workbook_id))

    async def analyze(self, user_id: int, workbook_id: str) -> WorkbookAnalysis:
        async with self._semaphore:
            return await asyncio.to_thread(
                lambda: analyze_workbook(self.read_current(user_id, workbook_id))
            )

    # ------------------------------------------------------------------ mutate
    async def mutate(
        self,
        user_id: int,
        workbook_id: str,
        op_name: str,
        op: EditOperation,
        summary: str,
    ) -> OperationResult:
        """Apply op(workbook) under lock, with backup and atomic validated save."""
        self.metadata.get(user_id, workbook_id)  # ownership check

        def _run() -> OperationResult:
            with self.locks.acquire(user_id, workbook_id):
                current = self._current_path(user_id, workbook_id)
                if not current.exists():
                    raise WorkbookNotFound("نسخهٔ کاری فایل یافت نشد.")

                meta = self.metadata.get(user_id, workbook_id)
                version = int(meta.get("version_counter", 0)) + 1

                # 1. backup current state
                backup = self._versions_dir(user_id, workbook_id) / f"version_{version:03d}.xlsx"
                shutil.copy2(current, backup)

                # 2. load, mutate
                wb = load_workbook_safe(current)
                result = op(wb)

                # 3. save to temp, validate, atomic replace
                tmp = self.fs.temp_dir() / f"{workbook_id}-{int(time.time() * 1000)}.xlsx"
                try:
                    wb.save(tmp)
                    validate_saved_workbook(tmp)
                    shutil.move(str(tmp), current)
                finally:
                    tmp.unlink(missing_ok=True)

                # 4. refresh cheap metadata (sheet names/states/active)
                self.metadata.update(
                    user_id,
                    workbook_id,
                    version_counter=version,
                    sheets=sheet_summary(wb),
                    active_sheet=wb.active.title if wb.active is not None else None,
                )
                self._rotate_versions(user_id, workbook_id)
                return OperationResult(True, summary, data=result, version=version)

        async with self._semaphore:
            try:
                return await asyncio.wait_for(
                    asyncio.to_thread(_run),
                    timeout=self.settings.workbook_operation_timeout_seconds,
                )
            except asyncio.TimeoutError as exc:
                raise OperationTimeout(
                    "عملیات بیش از حد طول کشید. ممکن است در پس‌زمینه کامل شود؛ "
                    "پیش از ادامه، وضعیت فایل را با «اطلاعات فایل» بررسی کنید."
                ) from exc

    def _rotate_versions(self, user_id: int, workbook_id: str) -> None:
        versions = sorted(self._versions_dir(user_id, workbook_id).glob("version_*.xlsx"))
        max_keep = self.settings.max_backup_versions
        for old in versions[:-max_keep] if len(versions) > max_keep else []:
            old.unlink(missing_ok=True)

    # ------------------------------------------------------------------ history
    def list_versions(self, user_id: int, workbook_id: str) -> list[dict]:
        self.metadata.get(user_id, workbook_id)
        return [
            {"file": f.name, "size": f.stat().st_size, "mtime": f.stat().st_mtime}
            for f in sorted(self._versions_dir(user_id, workbook_id).glob("version_*.xlsx"))
        ]

    async def restore_version(self, user_id: int, workbook_id: str, version_file: str) -> OperationResult:
        """Restore a backup as a NEW version (never destructive)."""

        def _restore() -> OperationResult:
            self.metadata.get(user_id, workbook_id)
            if "/" in version_file or "\\" in version_file or ".." in version_file:
                raise WorkbookNotFound("نام نسخهٔ نامعتبر است.")
            src = self._versions_dir(user_id, workbook_id) / version_file
            if not src.exists():
                raise WorkbookNotFound("نسخهٔ موردنظر یافت نشد.")
            return self._restore_sync(user_id, workbook_id, src)

        async with self._semaphore:
            return await asyncio.to_thread(_restore)

    def _restore_sync(self, user_id: int, workbook_id: str, src: Path) -> OperationResult:
        with self.locks.acquire(user_id, workbook_id):
            current = self._current_path(user_id, workbook_id)
            meta = self.metadata.get(user_id, workbook_id)
            version = int(meta.get("version_counter", 0)) + 1
            shutil.copy2(
                current, self._versions_dir(user_id, workbook_id) / f"version_{version:03d}.xlsx"
            )
            validate_saved_workbook(src)
            shutil.copy2(src, current)
            self._rotate_versions(user_id, workbook_id)
            self.metadata.update(user_id, workbook_id, version_counter=version)
            return OperationResult(True, f"نسخهٔ {src.name} بازیابی شد.", version=version)

    async def undo(self, user_id: int, workbook_id: str) -> OperationResult:
        versions = self.list_versions(user_id, workbook_id)
        if not versions:
            raise WorkbookNotFound("تغییری برای بازگردانی وجود ندارد.")
        return await self.restore_version(user_id, workbook_id, versions[-1]["file"])

    async def restore_original(self, user_id: int, workbook_id: str) -> OperationResult:
        self.metadata.get(user_id, workbook_id)
        original = self._original_path(user_id, workbook_id)
        if not original.exists():
            raise WorkbookNotFound("فایل اصلی یافت نشد.")
        async with self._semaphore:
            return await asyncio.to_thread(self._restore_sync, user_id, workbook_id, original)

    # ------------------------------------------------------------------ files
    def export_current(self, user_id: int, workbook_id: str) -> tuple[Path, str]:
        meta = self.metadata.get(user_id, workbook_id)
        return self._current_path(user_id, workbook_id), meta["display_name"]

    def list_workbooks(self, user_id: int) -> list[dict]:
        return self.metadata.list_workbooks(user_id)

    def delete_workbook(self, user_id: int, workbook_id: str) -> None:
        self.metadata.delete(user_id, workbook_id)
        self.metadata.unregister(user_id, workbook_id)
        session = self.sessions.get(user_id)
        if session.get("active_workbook_id") == workbook_id:
            self.sessions.reset(user_id)