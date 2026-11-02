"""Excel A1-style reference parsing and validation.

Raises InvalidCellReference for anything ambiguous — never guess.
"""
from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string, get_column_letter

from app.core.exceptions import InvalidCellReference

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$")

MAX_COLUMN = 16384  # XFD
MAX_ROW = 1048576


def parse_cell(ref: str) -> tuple[int, int]:
    """Return (row, column) for a single cell reference like 'B5' or '$B$5'."""
    m = _CELL_RE.match(ref.strip())
    if not m:
        raise InvalidCellReference(f"مرجع سلول نامعتبر است: «{ref}» — مثال درست: B5")
    try:
        col = column_index_from_string(m.group(1).upper())
    except ValueError as exc:  # e.g. 'ZZZ' beyond XFD
        raise InvalidCellReference(f"ستون «{m.group(1)}» خارج از محدودهٔ Excel است (حداکثر XFD).") from exc
    if col > MAX_COLUMN:
        raise InvalidCellReference("شمارهٔ ستون خارج از محدودهٔ Excel است.")
    row = int(m.group(2))
    if row > MAX_ROW:
        raise InvalidCellReference("شمارهٔ سطر خارج از محدودهٔ Excel است.")
    return row, col


def cell_name(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def parse_range(
    ref: str, max_row: int | None = None, max_column: int | None = None
) -> tuple[int, int, int, int]:
    """Normalize a reference to (min_col, min_row, max_col, max_row).

    Accepts: A1, A1:D10, A:A, 1:10. Full-column/row references require the
    caller to supply sheet dimensions; otherwise they are rejected.
    """
    ref = ref.strip().replace("$", "")
    if not ref:
        raise InvalidCellReference("مرجع خالی است.")

    if ":" not in ref:
        row, col = parse_cell(ref)
        return col, row, col, row

    left, _, right = ref.partition(":")
    left, right = left.strip(), right.strip()

    # Full column range: A:D
    if left.isalpha() and right.isalpha():
        if max_row is None:
            raise InvalidCellReference("برای مرجع ستونی کامل ابتدا شیت فعال را انتخاب کنید.")
        try:
            c1, c2 = column_index_from_string(left.upper()), column_index_from_string(right.upper())
        except ValueError as exc:
            raise InvalidCellReference("نام ستون نامعتبر است.") from exc
        c1, c2 = min(c1, c2), max(c1, c2)
        return c1, 1, c2, max_row

    # Full row range: 1:10
    if left.isdigit() and right.isdigit():
        r1, r2 = int(left), int(right)
        if r1 < 1 or r2 < 1:
            raise InvalidCellReference("شمارهٔ سطر باید ۱ یا بزرگ‌تر باشد.")
        if max_column is None:
            raise InvalidCellReference("برای مرجع سطری کامل ابتدا شیت فعال را انتخاب کنید.")
        return 1, min(r1, r2), max_column, max(r1, r2)

    # Cell rectangle: A1:D10
    r1, c1 = parse_cell(left)
    r2, c2 = parse_cell(right)
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def bounds_within_sheet(bounds: tuple[int, int, int, int], max_row: int, max_column: int) -> bool:
    min_col, min_row, max_col, max_row_ = bounds
    return (
        min_col >= 1
        and min_row >= 1
        and max_col <= max(max_column, 1)
        and max_row_ <= max(max_row, 1)
    )


def range_size(bounds: tuple[int, int, int, int]) -> int:
    min_col, min_row, max_col, max_row = bounds
    return (max_col - min_col + 1) * (max_row - min_row + 1)