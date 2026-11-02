"""Cell editing operations — pure functions over an open Workbook instance.

KNOWN LIMITATIONS (documented, never hidden from the user):
- insert/delete rows/columns shift values and styles but do NOT translate
  formula references (openpyxl behavior). LibreOffice recalculation, when
  enabled, is the separate remedy.
- copy_range copies formula strings verbatim (relative references are not
  adjusted).
- move_range moves cells without translating formulas.
"""
from __future__ import annotations

import datetime as _dt
from copy import copy as _copy
from dataclasses import dataclass

from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import absolute_coordinate
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.core.exceptions import InvalidValue, WorkbookError
from app.excel.operations.worksheets import get_sheet
from app.utils.range_utils import bounds_within_sheet, parse_cell, parse_range, range_size

MAX_RANGE_CELLS = 50_000  # guard against accidental whole-sheet operations


@dataclass
class CellSnapshot:
    ref: str
    value: object
    data_type: str
    number_format: str


def _check_bounds(ws: Worksheet, bounds: tuple[int, int, int, int]) -> None:
    if not bounds_within_sheet(bounds, ws.max_row, ws.max_column):
        raise WorkbookError("مرجع در محدودهٔ مجاز شیت نیست.")
    if range_size(bounds) > MAX_RANGE_CELLS:
        raise WorkbookError("محدودهٔ انتخابی بیش از حد بزرگ است (حداکثر ۵۰٬۰۰۰ سلول).")


def parse_value(raw: str) -> object:
    """Interpret user text as an Excel value. Never uses eval/exec."""
    raw = raw.strip()
    if raw == "":
        return None
    if raw.startswith("="):
        return raw  # formula
    low = raw.lower()
    if low in ("true", "درست"):
        return True
    if low in ("false", "نادرست"):
        return False
    try:
        return int(raw.replace(",", ""))
    except ValueError:
        pass
    try:
        return float(raw.replace(",", ""))
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return _dt.datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return raw


def read_range(
    wb: Workbook, sheet: str | None, ref: str, limit: int = 500
) -> tuple[list[list[CellSnapshot]], bool]:
    """Return (rows, truncated). Rows are capped at `limit`."""
    ws = get_sheet(wb, sheet)
    bounds = parse_range(ref, max_row=ws.max_row, max_column=ws.max_column)
    _check_bounds(ws, bounds)
    min_col, min_row, max_col, max_row = bounds
    truncated = (max_row - min_row + 1) > limit
    if truncated:
        max_row = min_row + limit - 1
    out: list[list[CellSnapshot]] = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        out.append(
            [
                CellSnapshot(f"{get_column_letter(c.column)}{c.row}", c.value, c.data_type, c.number_format)
                for c in row
            ]
        )
    return out, truncated


def read_cell(wb: Workbook, sheet: str | None, ref: str) -> CellSnapshot:
    rows, _ = read_range(wb, sheet, ref)
    return rows[0][0]


def write_cells(wb: Workbook, sheet: str | None, ref: str, values: list[str], mode: str = "fill") -> int:
    """mode='fill' writes values[0] to every cell; mode='sequence' fills row-major."""
    ws = get_sheet(wb, sheet)
    bounds = parse_range(ref, max_row=ws.max_row, max_column=ws.max_column)
    _check_bounds(ws, bounds)
    min_col, min_row, max_col, max_row = bounds
    if mode == "fill":
        parsed = parse_value(values[0])
        count = 0
        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for c in row:
                c.value = parsed
                count += 1
        return count
    needed = range_size(bounds)
    if len(values) < needed:
        raise InvalidValue(f"تعداد مقادیر ({len(values)}) کمتر از سلول‌های محدوده ({needed}) است.")
    i = 0
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for c in row:
            c.value = parse_value(values[i])
            i += 1
    return needed


def clear_range(wb: Workbook, sheet: str | None, ref: str, contents: bool = True, formats: bool = False) -> int:
    ws = get_sheet(wb, sheet)
    bounds = parse_range(ref, max_row=ws.max_row, max_column=ws.max_column)
    _check_bounds(ws, bounds)
    min_col, min_row, max_col, max_row = bounds
    count = 0
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for c in row:
            if contents:
                c.value = None
            if formats:
                c.font, c.fill, c.border, c.alignment = Font(), PatternFill(), Border(), Alignment()
            count += 1
    return count


def insert_rows(wb: Workbook, sheet: str | None, at: int, count: int = 1) -> None:
    """NOTE: does not translate formula references (see module docstring)."""
    get_sheet(wb, sheet).insert_rows(at, count)


def delete_rows(wb: Workbook, sheet: str | None, at: int, count: int = 1) -> None:
    """NOTE: does not translate formula references (see module docstring)."""
    get_sheet(wb, sheet).delete_rows(at, count)


def insert_cols(wb: Workbook, sheet: str | None, at: int, count: int = 1) -> None:
    get_sheet(wb, sheet).insert_cols(at, count)


def delete_cols(wb: Workbook, sheet: str | None, at: int, count: int = 1) -> None:
    get_sheet(wb, sheet).delete_cols(at, count)


def merge_cells(wb: Workbook, sheet: str | None, ref: str) -> str:
    ws = get_sheet(wb, sheet)
    bounds = parse_range(ref, max_row=ws.max_row, max_column=ws.max_column)
    if range_size(bounds) < 2:
        raise WorkbookError("برای ادغام حداقل دو سلول انتخاب کنید.")
    min_col, min_row, max_col, max_row = bounds
    coord = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    ws.merge_cells(str(absolute_coordinate(coord)))
    return coord


def unmerge_cells(wb: Workbook, sheet: str | None, ref: str) -> str:
    ws = get_sheet(wb, sheet)
    bounds = parse_range(ref, max_row=ws.max_row, max_column=ws.max_column)
    min_col, min_row, max_col, max_row = bounds
    coord = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    if coord not in [str(r) for r in ws.merged_cells.ranges]:
        raise WorkbookError(f"محدودهٔ «{coord}» ادغام‌شده نیست.")
    ws.unmerge_cells(str(absolute_coordinate(coord)))
    return coord


def set_row_height(wb: Workbook, sheet: str | None, row: int, height: float) -> None:
    if row < 1 or not (0 < height < 500):
        raise InvalidValue("ارتفاع سطر باید بین ۰ تا ۵۰۰ باشد.")
    get_sheet(wb, sheet).row_dimensions[row].height = height


def set_col_width(wb: Workbook, sheet: str | None, col: int, width: float) -> None:
    if col < 1 or not (0 < width < 300):
        raise InvalidValue("عرض ستون باید بین ۰ تا ۳۰۰ باشد.")
    get_sheet(wb, sheet).column_dimensions[get_column_letter(col)].width = width


def set_hidden(wb: Workbook, sheet: str | None, what: str, start: int, end: int, hidden: bool) -> None:
    ws = get_sheet(wb, sheet)
    dims = ws.row_dimensions if what == "row" else ws.column_dimensions
    for i in range(start, end + 1):
        key = i if what == "row" else get_column_letter(i)
        dims[key].hidden = hidden


def copy_range(wb: Workbook, sheet: str | None, source: str, target: str) -> int:
    """Copies values + styles. Formulas are copied verbatim (not translated)."""
    ws = get_sheet(wb, sheet)
    src = parse_range(source, max_row=ws.max_row, max_column=ws.max_column)
    t_row, t_col = parse_cell(target)
    _check_bounds(ws, src)
    min_col, min_row, max_col, max_row = src
    count = 0
    for r_off, row in enumerate(
        ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)
    ):
        for c_off, cell in enumerate(row):
            dst = ws.cell(row=t_row + r_off, column=t_col + c_off)
            dst.value = cell.value
            if cell.has_style:
                dst.font, dst.border, dst.fill = _copy(cell.font), _copy(cell.border), _copy(cell.fill)
                dst.number_format, dst.alignment = cell.number_format, _copy(cell.alignment)
            count += 1
    return count


def move_range(wb: Workbook, sheet: str | None, source: str, target: str) -> str:
    """Moves cells (values+styles) without translating formulas."""
    ws = get_sheet(wb, sheet)
    src = parse_range(source, max_row=ws.max_row, max_column=ws.max_column)
    t_row, t_col = parse_cell(target)
    _check_bounds(ws, src)
    min_col, min_row, max_col, max_row = src
    coord = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    ws.move_range(str(absolute_coordinate(coord)), rows=t_row - min_row, cols=t_col - min_col)
    return coord