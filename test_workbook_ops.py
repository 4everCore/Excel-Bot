"""Excel engine + service tests: fidelity of real workbooks through real operations."""
from __future__ import annotations

import asyncio
import io

import pytest
from openpyxl import Workbook, load_workbook

from app.core.config import Settings
from app.core.exceptions import WorkbookError, WorkbookNotFound, InvalidSheetName
from app.excel.operations import cells, worksheets
from app.services.workbook_service import WorkbookService
from app.storage.filesystem import FileSystem


def build_workbook_bytes() -> bytes:
    """Fixture workbook: formulas, styling, merged cells, hidden sheet, freeze panes.

    Styles are assigned as fresh immutable objects — openpyxl 3.x has no
    .copy(**kw) method on style objects.
    """
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "Financial"
    ws["A1"] = "Item"
    ws["B1"] = "Amount"
    ws["A1"].font = Font(bold=True)
    ws["B1"].fill = PatternFill("solid", fgColor="FFFF00")
    rows = [("Rent", 100), ("Food", 250), ("Transport", 60)]
    for i, (name, amount) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=name)
        c = ws.cell(row=i, column=2, value=amount)
        c.number_format = "#,##0.00"
    ws["B5"] = "=SUM(B2:B4)"
    ws.merge_cells("A7:B7")
    ws["A7"] = "Total report"
    ws.freeze_panes = "A2"
    wb.create_sheet("Hidden").sheet_state = "hidden"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture()
def service(tmp_path):
    s = Settings(bot_token="1:test")  # type: ignore[call-arg]
    s.storage_root = tmp_path
    fs = FileSystem(tmp_path)
    return WorkbookService(fs, s)


@pytest.fixture()
def workbook_id(service):
    meta = service.metadata.create(1, "test.xlsx", 0, [])
    path = service._current_path(1, meta["id"])
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(build_workbook_bytes())
    service.metadata.register(1, meta["id"])
    return meta["id"]


class TestCellOperations:
    def test_parse_value_variants(self):
        import datetime

        assert cells.parse_value("42") == 42
        assert cells.parse_value("3.5") == 3.5
        assert cells.parse_value("TRUE") is True
        assert cells.parse_value("=SUM(A1:A2)") == "=SUM(A1:A2)"
        assert cells.parse_value("hello") == "hello"
        assert cells.parse_value("2024-01-05") == datetime.date(2024, 1, 5)

    def test_write_and_read_roundtrip(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        n = cells.write_cells(wb, "Financial", "C1:C3", ["added"], mode="fill")
        assert n == 3
        cells.write_cells(wb, "Financial", "D1", ["=SUM(A1:A2)"], mode="fill")
        snap = cells.read_cell(wb, "Financial", "D1")
        assert snap.value == "=SUM(A1:A2)"

    def test_row_insert_shifts_values_but_not_formulas(self, workbook_id, service):
        """Documents the real openpyxl behavior: values/styles shift, formula
        references are NOT translated. This is a known limitation, not a bug."""
        wb = service.read_current(1, workbook_id)
        cells.insert_rows(wb, "Financial", 3, 1)
        assert wb["Financial"]["B5"].value == 60  # old row 4 shifted down
        assert wb["Financial"]["B6"].value == "=SUM(B2:B4)"  # moved, untranslated

    def test_merge_and_unmerge(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        coord = cells.merge_cells(wb, "Financial", "D1:E2")
        assert coord in [str(r) for r in wb["Financial"].merged_cells.ranges]
        cells.unmerge_cells(wb, "Financial", "D1:E2")
        assert coord not in [str(r) for r in wb["Financial"].merged_cells.ranges]

    def test_oversized_range_rejected(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        with pytest.raises(WorkbookError):
            cells.write_cells(wb, "Financial", "A1:ZZ50000", ["x"], mode="fill")

    def test_beyond_xfd_column_rejected(self):
        from app.core.exceptions import InvalidCellReference

        with pytest.raises(InvalidCellReference):
            cells.parse_value  # noqa: B018 - placeholder removal check
        from app.utils.range_utils import parse_cell

        with pytest.raises(InvalidCellReference):
            parse_cell("ZZZ1")

    def test_fidelity_unrelated_features_survive_edit(self, workbook_id, service):
        """Editing one cell must not damage merges, freeze panes, number
        formats or the hidden sheet elsewhere in the workbook."""
        wb = service.read_current(1, workbook_id)
        cells.write_cells(wb, "Financial", "C1", ["hello"], mode="fill")
        ws = wb["Financial"]
        assert "A7:B7" in [str(r) for r in ws.merged_cells.ranges]
        assert ws.freeze_panes == "A2"
        assert ws["B2"].number_format == "#,##0.00"
        assert wb["Hidden"].sheet_state == "hidden"
        assert wb["Financial"]["A1"].font.bold is True  # untouched header style


class TestWorksheetOperations:
    def test_create_duplicate_rejected(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        with pytest.raises(InvalidSheetName):
            worksheets.create_sheet(wb, "Financial")

    def test_delete_last_sheet_rejected(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        worksheets.delete_sheet(wb, "Hidden")
        with pytest.raises(WorkbookError):
            worksheets.delete_sheet(wb, "Financial")

    def test_illegal_name_rejected(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        with pytest.raises(InvalidSheetName):
            worksheets.rename_sheet(wb, "Hidden", "bad:name")

    def test_hide_last_visible_sheet_rejected(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        worksheets.delete_sheet(wb, "Hidden")
        with pytest.raises(WorkbookError):
            worksheets.set_sheet_state(wb, "Financial", "hidden")

    def test_delete_active_sheet_reaims_active(self, workbook_id, service):
        wb = service.read_current(1, workbook_id)
        worksheets.delete_sheet(wb, "Financial")
        assert wb.active is not None
        assert wb.active.title in wb.sheetnames
        assert wb.active.sheet_state == "visible"


class TestServiceSafety:
    async def test_upload_validates_and_stores(self, service, tmp_path):
        meta = await service.upload(1, build_workbook_bytes(), "report.xlsx")
        assert meta["active_sheet"] == "Financial"
        assert len(meta["sheets"]) == 2
        # original preserved: compare loaded structure (zip timestamps make
        # raw bytes of two saves differ even for identical content)
        orig = (tmp_path / "users" / "1" / "workbooks" / meta["id"] / "original.xlsx").read_bytes()
        a = load_workbook(io.BytesIO(orig))
        b = load_workbook(io.BytesIO(build_workbook_bytes()))
        assert a.sheetnames == b.sheetnames
        assert a["Financial"]["B5"].value == b["Financial"]["B5"].value

    async def test_mutate_creates_version_and_undo_restores(self, service, workbook_id):
        r1 = await service.mutate(
            1, workbook_id, "edit", lambda wb: wb["Financial"].__setitem__("A1", "CHANGED"),
            "ویرایش انجام شد",
        )
        assert r1.ok and r1.version == 1
        wb = service.read_current(1, workbook_id)
        assert wb["Financial"]["A1"].value == "CHANGED"

        await service.undo(1, workbook_id)
        wb = service.read_current(1, workbook_id)
        assert wb["Financial"]["A1"].value == "Item"

    async def test_failed_mutation_keeps_previous_state(self, service, workbook_id):
        def bad_op(wb):
            wb["Financial"]["A1"] = "PARTIAL"
            raise RuntimeError("simulated crash mid-operation")

        with pytest.raises(RuntimeError):
            await service.mutate(1, workbook_id, "edit", bad_op, "never")
        wb = service.read_current(1, workbook_id)
        # working copy untouched (save happens only after op succeeds)
        assert wb["Financial"]["A1"].value == "Item"

    async def test_ownership_isolated_between_users(self, service, workbook_id):
        with pytest.raises(WorkbookNotFound):
            service.metadata.get(2, workbook_id)

    async def test_version_rotation(self, service, workbook_id):
        service.settings.max_backup_versions = 2
        for i in range(4):
            await service.mutate(1, workbook_id, "edit", lambda wb: None, f"op {i}")
        assert len(service.list_versions(1, workbook_id)) <= 2

    async def test_restore_original(self, service, workbook_id):
        await service.mutate(
            1, workbook_id, "edit", lambda wb: wb["Financial"].__setattr__("A1", "X"), "x"
        )
        await service.restore_original(1, workbook_id)
        wb = service.read_current(1, workbook_id)
        assert wb["Financial"]["A1"].value == "Item"

    async def test_concurrent_mutates_serialize(self, service, workbook_id):
        """Two concurrent edits of the SAME workbook must serialize under the
        file lock and produce distinct versions."""

        def make_op(i):
            return lambda wb: wb["Financial"].__setitem__("Z1", i)

        results = await asyncio.gather(
            *[service.mutate(1, workbook_id, "edit", make_op(i), f"op{i}") for i in range(3)]
        )
        assert all(r.ok for r in results)
        assert sorted(r.version for r in results) == [1, 2, 3]

    async def test_mutate_refreshes_sheet_metadata(self, service, workbook_id):
        def rename(wb):
            worksheets.rename_sheet(wb, "Hidden", "Renamed")

        await service.mutate(1, workbook_id, "rename", rename, "renamed")
        meta = service.metadata.get(1, workbook_id)
        names = [s["name"] for s in meta["sheets"]]
        assert "Renamed" in names and "Hidden" not in names