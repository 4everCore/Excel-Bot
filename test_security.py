"""Security tests: hostile archives, macros, path traversal, isolation."""
from __future__ import annotations

import io
import zipfile

import pytest

from app.core import security
from app.core.exceptions import CorruptedWorkbook, DangerousPackage
from app.storage.filesystem import FileSystem
from app.core.exceptions import WorkbookNotFound


def _zip_bytes(entries: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)
    return buf.getvalue()


GOOD_PARTS = {
    "[Content_Types].xml": b"<Types/>",
    "xl/workbook.xml": b"<workbook/>",
    "xl/worksheets/sheet1.xml": b"<worksheet/>",
}


class TestPackageInspection:
    def test_valid_package_passes(self):
        report = security.inspect_xlsx_package(_zip_bytes(GOOD_PARTS), "report.xlsx")
        assert report.ok

    def test_rejects_non_xlsx_extension(self):
        with pytest.raises(DangerousPackage):
            security.inspect_xlsx_package(b"whatever", "data.csv")

    def test_rejects_xlsm(self):
        with pytest.raises(DangerousPackage):
            security.inspect_xlsx_package(_zip_bytes(GOOD_PARTS), "macros.xlsm")

    def test_rejects_corrupted_zip(self):
        with pytest.raises(CorruptedWorkbook):
            security.inspect_xlsx_package(b"this is not a zip at all", "bad.xlsx")

    def test_rejects_zip_missing_required_parts(self):
        with pytest.raises(CorruptedWorkbook):
            security.inspect_xlsx_package(_zip_bytes({"random.txt": b"hi"}), "bad.xlsx")

    def test_rejects_macro_workbook(self):
        entries = {**GOOD_PARTS, "xl/vbaProject.bin": b"MZ fake binary"}
        with pytest.raises(DangerousPackage):
            security.inspect_xlsx_package(_zip_bytes(entries), "virus.xlsx")

    def test_flags_external_links_as_warning(self):
        entries = {**GOOD_PARTS, "xl/externalLinks/externalLink1.xml": b"<links/>"}
        report = security.inspect_xlsx_package(_zip_bytes(entries), "linked.xlsx")
        assert report.feature_flags.get("external_links") is True

    def test_charts_and_pivots_flagged_not_rejected(self):
        entries = {
            **GOOD_PARTS,
            "xl/charts/chart1.xml": b"<chart/>",
            "xl/pivotTables/pivotTable1.xml": b"<pivot/>",
        }
        report = security.inspect_xlsx_package(_zip_bytes(entries), "ch.xlsx")
        assert report.feature_flags.get("charts") is True
        assert report.feature_flags.get("pivot_tables") is True
        assert any("نمودار" in w for w in report.warnings)

    def test_rejects_zip_slip_entry_names(self):
        entries = {**GOOD_PARTS, "../evil.xml": b"x"}
        with pytest.raises(DangerousPackage):
            security.inspect_xlsx_package(_zip_bytes(entries), "evil.xlsx")

    def test_rejects_zip_bomb_shape(self):
        # Central directory claims 300 MB uncompressed for a 3 MB entry —
        # must be rejected by the single-entry limit.
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(zipfile.ZipInfo("xl/worksheets/huge.xml"), b"0" * (3 * 1024 * 1024))
            zf.filelist[-1].file_size = 300 * 1024 * 1024
        with pytest.raises(DangerousPackage):
            security.inspect_xlsx_package(buf.getvalue(), "bomb.xlsx")


class TestPathSafety:
    def test_workbook_id_path_traversal_rejected(self, tmp_path):
        fs = FileSystem(tmp_path)
        for evil in ("../../etc", "..", "abc/def", "not-a-uuid", ""):
            with pytest.raises(WorkbookNotFound):
                fs.workbook_dir(1, evil)

    def test_uuid_ids_resolve_only_inside_user_dir(self, tmp_path):
        import uuid

        fs = FileSystem(tmp_path)
        wid = str(uuid.uuid4())
        created = fs.workbook_dir(1, wid, create=True)
        assert created.is_relative_to(tmp_path / "users" / "1")
        with pytest.raises(WorkbookNotFound):  # another user cannot see it
            fs.workbook_dir(2, wid)