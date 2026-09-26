"""
tools/excel-engine/core/cell_addressing.py - Chuyen doi dia chi o A1 <-> (row, col) va dinh dang
locator o.

Quy uoc: row va col deu 1-indexed (A1 == (1, 1)); sheet index trong locator cung 1-indexed theo
thu tu xuat hien trong `xl/workbook.xml`. Locator o co dang `cell_s{sheet}_r{row}_c{col}`, vi du
`cell_s1_r3_c2` == sheet thu 1, o B3.
"""

from __future__ import annotations

import re
from typing import Tuple

MAX_ROW = 1_048_576
MAX_COL = 16_384  # XFD

_A1_RE = re.compile(r"^([A-Za-z]{1,3})([1-9][0-9]{0,6})$")
_RANGE_RE = re.compile(r"^([A-Za-z]{1,3}[1-9][0-9]{0,6})(?::([A-Za-z]{1,3}[1-9][0-9]{0,6}))?$")
_LOCATOR_RE = re.compile(r"^cell_s([1-9][0-9]*)_r([1-9][0-9]*)_c([1-9][0-9]*)$")


class CellAddressError(ValueError):
    pass


def col_to_index(letters: str) -> int:
    if not letters or not letters.isalpha() or not letters.isascii():
        raise CellAddressError(f"Ten cot khong hop le: {letters!r}")
    index = 0
    for ch in letters.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    if not 1 <= index <= MAX_COL:
        raise CellAddressError(f"Cot {letters!r} nam ngoai gioi han A..XFD")
    return index


def index_to_col(index: int) -> str:
    if not isinstance(index, int) or isinstance(index, bool) or not 1 <= index <= MAX_COL:
        raise CellAddressError(f"Chi so cot ngoai gioi han 1..{MAX_COL}: {index!r}")
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def a1_to_rc(ref: str) -> Tuple[int, int]:
    """'B3' -> (3, 2). Khong chap nhan dia chi tuyet doi ($) hay ten sheet."""
    match = _A1_RE.match(ref.strip() if isinstance(ref, str) else "")
    if not match:
        raise CellAddressError(f"Dia chi o A1 khong hop le: {ref!r}")
    col = col_to_index(match.group(1))
    row = int(match.group(2))
    if row > MAX_ROW:
        raise CellAddressError(f"Dong {row} vuot gioi han {MAX_ROW}")
    return row, col


def rc_to_a1(row: int, col: int) -> str:
    if not isinstance(row, int) or isinstance(row, bool) or not 1 <= row <= MAX_ROW:
        raise CellAddressError(f"Chi so dong ngoai gioi han 1..{MAX_ROW}: {row!r}")
    return f"{index_to_col(col)}{row}"


def is_valid_a1(ref: str) -> bool:
    try:
        a1_to_rc(ref)
    except CellAddressError:
        return False
    return True


def parse_range(ref: str) -> Tuple[int, int, int, int]:
    """'A1:C3' -> (r1, c1, r2, c2) da chuan hoa min/max. 'B2' -> (2, 2, 2, 2)."""
    match = _RANGE_RE.match(ref.strip() if isinstance(ref, str) else "")
    if not match:
        raise CellAddressError(f"Vung o khong hop le: {ref!r}")
    r1, c1 = a1_to_rc(match.group(1))
    r2, c2 = a1_to_rc(match.group(2) or match.group(1))
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def range_to_a1(r1: int, c1: int, r2: int, c2: int) -> str:
    if (r1, c1) == (r2, c2):
        return rc_to_a1(r1, c1)
    return f"{rc_to_a1(r1, c1)}:{rc_to_a1(r2, c2)}"


def in_range(row: int, col: int, rng: Tuple[int, int, int, int]) -> bool:
    r1, c1, r2, c2 = rng
    return r1 <= row <= r2 and c1 <= col <= c2


def make_locator(sheet_index: int, row: int, col: int) -> str:
    if sheet_index < 1:
        raise CellAddressError(f"sheet_index phai >= 1: {sheet_index!r}")
    rc_to_a1(row, col)  # validate bounds
    return f"cell_s{sheet_index}_r{row}_c{col}"


def parse_locator(locator: str) -> Tuple[int, int, int]:
    """'cell_s1_r3_c2' -> (sheet_index=1, row=3, col=2)."""
    match = _LOCATOR_RE.match(locator if isinstance(locator, str) else "")
    if not match:
        raise CellAddressError(f"Locator o khong hop le: {locator!r}")
    sheet_index, row, col = (int(g) for g in match.groups())
    rc_to_a1(row, col)  # validate bounds
    return sheet_index, row, col


def is_valid_locator(locator: str) -> bool:
    try:
        parse_locator(locator)
    except CellAddressError:
        return False
    return True
