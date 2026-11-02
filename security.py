"""Security inspection for untrusted Excel packages.

Uploaded spreadsheets are hostile input. Every upload passes through
inspect_xlsx_package() before anything is written to disk.
"""
from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass, field

from app.core.exceptions import CorruptedWorkbook, DangerousPackage

# Hard limits for the OOXML zip package itself
MAX_ARCHIVE_ENTRIES = 5_000
MAX_TOTAL_UNCOMPRESSED = 400 * 1024 * 1024  # 400 MB expanded
MAX_SINGLE_ENTRY_UNCOMPRESSED = 200 * 1024 * 1024
SUSPICIOUS_RATIO = 1_000  # expanded/compressed
SUSPICIOUS_MIN_SIZE = 10 * 1024 * 1024  # only flag big entries

REQUIRED_PARTS = ("[Content_Types].xml", "xl/workbook.xml")

# Parts that mean features we will NOT execute or silently drop
MACRO_PARTS = ("xl/vbaProject.bin",)
EXTERNAL_PARTS = ("xl/externalLinks/", "xl/connections.xml", "xl/embeddings/")


@dataclass
class PackageReport:
    ok: bool = True
    warnings: list[str] = field(default_factory=list)
    feature_flags: dict[str, bool] = field(default_factory=dict)


def _fail(msg: str) -> None:
    raise DangerousPackage(msg)


def inspect_xlsx_package(data: bytes, original_name: str) -> PackageReport:
    """Validate the zip package of an uploaded workbook. Raises on fatal issues."""
    name = original_name.lower()
    if name.endswith(".xlsm"):
        raise DangerousPackage(
            "فایل‌های ماکرو‌دار (.xlsm) در این نسخه ویرایش نمی‌شوند؛ ماکروها اجرا یا حفظ نمی‌شوند."
        )
    if not name.endswith(".xlsx"):
        raise DangerousPackage("فقط فایل‌های .xlsx پذیرفته می‌شوند.")

    report = PackageReport()

    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise CorruptedWorkbook("فایل یک بستهٔ معتبر Excel (OOXML) نیست.")

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            bad = zf.testzip()
            if bad is not None:
                raise CorruptedWorkbook(f"بستهٔ فایل خراب است (عضو آسیب‌دیده: {bad[:40]}).")

            names = zf.namelist()
            if len(names) > MAX_ARCHIVE_ENTRIES:
                _fail("ساختار فایل مشکوک است (تعداد اعضای بسته بیش از حد مجاز).")

            # Zip-slip defense: reject hostile entry names even though we
            # never extract members to disk ourselves.
            for entry in names:
                if ".." in entry or "\\" in entry or entry.startswith("/"):
                    _fail("نام یکی از اجزای بسته نامعتبر است.")

            total_uncompressed = 0
            for info in zf.infolist():
                total_uncompressed += info.file_size
                if info.file_size > MAX_SINGLE_ENTRY_UNCOMPRESSED:
                    _fail("حجم یکی از اجزای فایل غیرعادی است.")
                ratio = info.file_size / max(info.compress_size, 1)
                if ratio > SUSPICIOUS_RATIO and info.file_size > SUSPICIOUS_MIN_SIZE:
                    _fail("الگوی فشرده‌سازی فایل مشکوک است (احتمال zip bomb).")

            if total_uncompressed > MAX_TOTAL_UNCOMPRESSED:
                _fail("حجم واقعی محتوای فایل بیش از حد مجاز است.")

            missing = [p for p in REQUIRED_PARTS if p not in names]
            if missing:
                raise CorruptedWorkbook("ساختار داخلی فایل Excel کامل نیست.")

            if any(n in MACRO_PARTS for n in names):
                _fail("فایل حاوی ماکرو است و پذیرفته نمی‌شود.")

            has_external = any(
                n.startswith(prefix) or n == prefix for n in names for prefix in EXTERNAL_PARTS
            )
            if has_external:
                report.feature_flags["external_links"] = True
                report.warnings.append(
                    "این فایل پیوند خارجی یا دادهٔ embedded دارد؛ در ویرایش ممکن است از دست برود."
                )

            # openpyxl does NOT read charts back from existing files: any
            # chart in this workbook will be lost on the first edit-save.
            # Never silent — surface it at upload time.
            if any(n.startswith("xl/charts/") for n in names):
                report.feature_flags["charts"] = True
                report.warnings.append(
                    "این فایل نمودار دارد. متأسفانه نمودارهای موجود هنگام ذخیرهٔ "
                    "ویرایش‌ها حذف می‌شوند؛ برای حفظ نمودارها، ویرایش را روی کپی "
                    "دیگری انجام دهید یا نمودار را دوباره بسازید."
                )
            if any(n.startswith("xl/media/") for n in names):
                report.feature_flags["images"] = True
                report.warnings.append(
                    "این فایل تصویر دارد؛ حفظ تصاویر به نصب بودن Pillow روی سرور وابسته است."
                )
            if any(n.startswith("xl/pivotTables/") for n in names):
                report.feature_flags["pivot_tables"] = True
                report.warnings.append(
                    "این فایل جدول محوری (PivotTable) دارد که در ویرایش ممکن است از دست برود."
                )
    except zipfile.BadZipFile as exc:
        raise CorruptedWorkbook("بستهٔ ZIP فایل مخدوش است.") from exc

    return report