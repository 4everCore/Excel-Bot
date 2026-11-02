"""Worksheet operations — pure functions over an open Workbook instance."""
from __future__ import annotations

from openpyxl.workbook.workbook import Workbook

from app.core.exceptions import InvalidSheetName, WorkbookError
from app.excel.engine import SheetInfo

ILLEGAL_SHEET_CHARS = set("[]:*?/\\")
MAX_SHEET_NAME_LEN = 31


def get_sheet(wb: Workbook, name: str | None):
    if name:
        if name not in wb.sheetnames:
            raise WorkbookError(f"شیت «{name}» وجود ندارد.")
        return wb[name]
    if wb.active is None:
        raise WorkbookError("شیت فعالی وجود ندارد.")
    return wb.active


def list_sheets(wb: Workbook) -> list[SheetInfo]:
    return [
        SheetInfo(
            ws.title,
            str(ws.sheet_state),
            ws.max_row,
            ws.max_column,
            sum(1 for row in ws.iter_rows() for c in row if c.value is not None),
        )
        for ws in wb.worksheets
    ]


def sheet_summary(wb: Workbook) -> list[dict]:
    """Cheap sheet snapshot for metadata (no cell iteration)."""
    return [
        {
            "name": ws.title,
            "state": str(ws.sheet_state),
            "max_row": ws.max_row,
            "max_column": ws.max_column,
        }
        for ws in wb.worksheets
    ]


def _validate_name(name: str) -> str:
    name = name.strip()
    if not name or len(name) > MAX_SHEET_NAME_LEN:
        raise InvalidSheetName("نام شیت باید ۱ تا ۳۱ کاراکتر باشد.")
    if any(ch in ILLEGAL_SHEET_CHARS for ch in name):
        raise InvalidSheetName("نام شیت نمی‌تواند شامل کاراکترهای [ ] : * ? / \\ باشد.")
    if name.startswith("'") or name.endswith("'"):
        raise InvalidSheetName("نام شیت نمی‌تواند با آپاستروف شروع یا تمام شود.")
    return name


def create_sheet(wb: Workbook, name: str, position: int | None = None) -> str:
    name = _validate_name(name)
    if name in wb.sheetnames:
        raise InvalidSheetName(f"شیت «{name}» از قبل وجود دارد.")
    ws = wb.create_sheet(name, position)
    wb.active = wb.sheetnames.index(name)
    return ws.title


def delete_sheet(wb: Workbook, name: str | None) -> None:
    ws = get_sheet(wb, name)
    if len(wb.sheetnames) <= 1:
        raise WorkbookError("حذف آخرین شیت مجاز نیست.")
    was_active = wb.active is not None and wb.active.title == ws.title
    wb.remove(ws)
    if was_active:
        # openpyxl leaves wb.active pointing at a stale/possibly hidden index;
        # re-aim it at the first visible sheet.
        for idx, sheet in enumerate(wb.worksheets):
            if sheet.sheet_state == "visible":
                wb.active = idx
                break
        else:
            wb.active = 0


def rename_sheet(wb: Workbook, old: str | None, new: str) -> str:
    ws = get_sheet(wb, old)
    new = _validate_name(new)
    if new != ws.title and new in wb.sheetnames:
        raise InvalidSheetName(f"شیت «{new}» از قبل وجود دارد.")
    ws.title = new
    return new


def copy_sheet(wb: Workbook, name: str | None, new_name: str | None = None) -> str:
    src = get_sheet(wb, name)
    target = _validate_name(new_name) if new_name else _validate_name(f"{src.title} (کپی)")
    base, i = target, 2
    while target in wb.sheetnames:
        target = f"{base} {i}"
        i += 1
    wb.copy_worksheet(src).title = target
    return target


def reorder_sheet(wb: Workbook, name: str, offset: int) -> list[str]:
    if name not in wb.sheetnames:
        raise WorkbookError(f"شیت «{name}» وجود ندارد.")
    wb.move_sheet(name, offset=offset)
    return wb.sheetnames


def set_sheet_state(wb: Workbook, name: str, state: str) -> None:
    if state not in ("visible", "hidden"):
        raise InvalidSheetName("وضعیت شیت باید visible یا hidden باشد.")
    ws = get_sheet(wb, name)
    visible_count = sum(1 for s in wb.worksheets if s.sheet_state == "visible")
    if state == "hidden" and visible_count <= 1:
        raise WorkbookError("حداقل یک شیت باید نمایان بماند.")
    ws.sheet_state = state


def set_active_sheet(wb: Workbook, name: str) -> str:
    ws = get_sheet(wb, name)
    wb.active = wb.sheetnames.index(ws.title)
    return ws.title