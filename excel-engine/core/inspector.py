"""
tools/excel-engine/core/inspector.py - Doc (read-only) workbook .xlsx bang lxml truc tiep: cau
truc workbook, locator tung o, dimension, merged cells, Table, vung array formula, cong voi Reader
(`read_sheet`, `read_cell`) tra ve JSON sach.

Read-only tuyet doi: khong ghi bat ky file nao (khong sidecar index, khong `.jarvis/`). Moi XML
duoc parse qua `safety_gateway.parse_xml_safe()` (tat XXE).

Quy uoc gia tri o cong thuc: `cached_value` la gia tri Excel da luu trong `<v>`; neu o cong thuc
chua co `<v>` (vd o cong thuc moi vua duoc patch) HOAC workbook dang dat `fullCalcOnLoad="1"`
(gia tri luu co the cu) thi `cached_value: null`/gia tri cu va `stale: true`. Engine KHONG tu
tinh cong thuc.
"""

from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Tuple, Union

from lxml import etree

from core.cell_addressing import (
    a1_to_rc,
    in_range,
    make_locator,
    parse_locator,
    parse_range,
    range_to_a1,
    rc_to_a1,
)
from core.safety_gateway import compute_document_revision, parse_xml_safe
from core.xlsx_package import XlsxFormatError, XlsxPackage, reject_legacy_suffix

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DEFAULT_LOCATOR_LIMIT = 5000
_INT_RE = re.compile(r"^-?[0-9]{1,15}$")


def q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


class SheetNotFoundError(LookupError):
    pass


class LocatorNotFoundError(LookupError):
    pass


class UnsupportedSheetError(ValueError):
    """Sheet khong phai worksheet (chartsheet, dialogsheet, macrosheet...): khong co du lieu o."""


Rect = Tuple[int, int, int, int]


@dataclass
class SheetInfo:
    index: int  # 1-indexed
    name: str
    sheet_id: str
    state: str
    part: str
    rid: str
    kind: str = "worksheet"  # theo Type cua relationship: worksheet | chartsheet | dialogsheet | ...


@dataclass
class TableInfo:
    name: str
    part: str
    ref: Rect
    header_rows: int


# ---------------------------------------------------------------------------
# Package-level helpers
# ---------------------------------------------------------------------------
def _resolve_target(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _rels_part_for(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _read_relationships(pkg: XlsxPackage, part: str) -> List[Dict[str, str]]:
    rels_part = _rels_part_for(part)
    if not pkg.has(rels_part):
        return []
    root = parse_xml_safe(pkg.get(rels_part))
    return [
        {"id": r.get("Id", ""), "type": r.get("Type", ""), "target": _resolve_target(part, r.get("Target", "")),
         "mode": r.get("TargetMode", "")}
        for r in root
        if isinstance(r.tag, str)
    ]


def load_workbook_root(pkg: XlsxPackage) -> etree._Element:
    return parse_xml_safe(pkg.get("xl/workbook.xml"))


def load_sheets(pkg: XlsxPackage) -> List[SheetInfo]:
    root = load_workbook_root(pkg)
    rels = {r["id"]: r for r in _read_relationships(pkg, "xl/workbook.xml")}
    sheets_el = root.find(q("sheets"))
    if sheets_el is None:
        raise XlsxFormatError("workbook.xml khong co <sheets>.")
    out: List[SheetInfo] = []
    for idx, el in enumerate(sheets_el.findall(q("sheet")), start=1):
        rid = el.get(f"{{{NS_REL}}}id", "")
        rel = rels.get(rid)
        if rel is None or not pkg.has(rel["target"]):
            raise XlsxFormatError(f"Sheet {el.get('name')!r}: relationship {rid!r} khong tro toi part ton tai.")
        out.append(SheetInfo(idx, el.get("name", ""), el.get("sheetId", ""), el.get("state", "visible"), rel["target"], rid,
                        rel["type"].rsplit("/", 1)[-1]))
    return out


def resolve_sheet(sheets: List[SheetInfo], ref: Union[str, int]) -> SheetInfo:
    """Tim sheet theo ten (chinh xac), neu khong thi theo chi so 1-indexed (int hoac chuoi so)."""
    for s in sheets:
        if s.name == ref:
            return s
    if (isinstance(ref, int) and not isinstance(ref, bool)) or (isinstance(ref, str) and ref.isdigit()):
        idx = int(ref)
        if 1 <= idx <= len(sheets):
            return sheets[idx - 1]
    raise SheetNotFoundError(f"Khong tim thay sheet {ref!r}. Co: {[s.name for s in sheets]}")


def load_shared_strings(pkg: XlsxPackage) -> List[str]:
    if not pkg.has("xl/sharedStrings.xml"):
        return []
    root = parse_xml_safe(pkg.get("xl/sharedStrings.xml"))
    out: List[str] = []
    for si in root.findall(q("si")):
        # Ghep moi <t> (ke ca rich-text run <r><t>), bo qua phonetic <rPh>.
        parts: List[str] = []
        for node in si.iter(q("t")):
            parent = node.getparent()
            if parent is not None and parent.tag == q("rPh"):
                continue
            parts.append(node.text or "")
        out.append("".join(parts))
    return out


def sheet_tables(pkg: XlsxPackage, sheet: SheetInfo) -> List[TableInfo]:
    tables: List[TableInfo] = []
    for rel in _read_relationships(pkg, sheet.part):
        if not rel["type"].endswith("/table") or not pkg.has(rel["target"]):
            continue
        troot = parse_xml_safe(pkg.get(rel["target"]))
        ref = troot.get("ref")
        if not ref:
            continue
        header_rows = int(troot.get("headerRowCount", "1") or "0")
        tables.append(TableInfo(troot.get("displayName") or troot.get("name") or "", rel["target"], parse_range(ref), header_rows))
    return tables


def sheet_merged(root: etree._Element) -> List[Rect]:
    merged = root.find(q("mergeCells"))
    if merged is None:
        return []
    return [parse_range(m.get("ref")) for m in merged.findall(q("mergeCell")) if m.get("ref")]


def sheet_array_ranges(root: etree._Element) -> List[Rect]:
    """Vung array formula / data table (`<f t="array|dataTable" ref=...>`)."""
    out: List[Rect] = []
    for f in root.iter(q("f")):
        if f.get("t") in ("array", "dataTable") and f.get("ref"):
            out.append(parse_range(f.get("ref")))
    return out


def sheet_data(root: etree._Element) -> etree._Element:
    data = root.find(q("sheetData"))
    if data is None:
        raise XlsxFormatError("worksheet khong co <sheetData>.")
    return data


def iter_cells(root: etree._Element) -> Iterator[Tuple[int, int, etree._Element]]:
    """Duyet moi <c> theo thu tu tai lieu, tra (row, col, element). Suy ra chi so khi thieu `r`."""
    prev_row = 0
    for row in sheet_data(root).findall(q("row")):
        r = int(row.get("r")) if row.get("r") else prev_row + 1
        prev_row = r
        prev_col = 0
        for c in row.findall(q("c")):
            ref = c.get("r")
            if ref:
                rr, cc = a1_to_rc(ref)
                if rr != r:
                    raise XlsxFormatError(f"O {ref} nam trong <row r={r}> - workbook khong nhat quan.")
            else:
                cc = prev_col + 1
            prev_col = cc
            yield r, cc, c


def full_calc_on_load(pkg: XlsxPackage) -> bool:
    calc = load_workbook_root(pkg).find(q("calcPr"))
    return calc is not None and calc.get("fullCalcOnLoad") in ("1", "true")


# ---------------------------------------------------------------------------
# Cell parsing
# ---------------------------------------------------------------------------
def _num(text: Optional[str]) -> Any:
    if text is None:
        return None
    text = text.strip()
    if _INT_RE.match(text):
        return int(text)
    try:
        return float(text)
    except ValueError:
        return text


_X_ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def decode_x_escapes(text: str) -> str:
    """Giai ma `_xHHHH_` nhu Excel (`_x005F_` -> `_`). Escape surrogate rieng le duoc giu nguyen."""
    def repl(m: "re.Match[str]") -> str:
        code = int(m.group(1), 16)
        return m.group(0) if 0xD800 <= code <= 0xDFFF else chr(code)
    return _X_ESCAPE_RE.sub(repl, text)


def _parse_value(c: etree._Element, shared: List[str]) -> Tuple[str, Any]:
    t = c.get("t", "n")
    v = c.find(q("v"))
    if t == "inlineStr":
        is_el = c.find(q("is"))
        if is_el is None:
            return "blank", None
        text = "".join((n.text or "") for n in is_el.iter(q("t")) if n.getparent() is None or n.getparent().tag != q("rPh"))
        return "string", decode_x_escapes(text)
    if v is None or v.text is None:
        return ("string", "") if (t == "str" and v is not None) else ("blank", None)
    if t == "s":
        try:
            return "string", decode_x_escapes(shared[int(v.text)])
        except (ValueError, IndexError):
            return "error", f"#SHAREDSTRING[{v.text}]"
    if t == "str":
        return "string", decode_x_escapes(v.text)
    if t == "b":
        return "bool", v.text.strip() == "1"
    if t == "e":
        return "error", v.text
    if t == "d":
        return "date", v.text
    parsed = _num(v.text)
    return ("number", parsed) if not isinstance(parsed, str) else ("string", parsed)


def parse_cell(
    c: etree._Element, row: int, col: int, sheet_index: int, shared: List[str],
    merges: List[Rect], calc_on_load: bool,
) -> Dict[str, Any]:
    vtype, value = _parse_value(c, shared)
    f = c.find(q("f"))
    has_formula = f is not None
    formula_text = (f.text if has_formula else None) or None
    merge: Optional[Dict[str, Any]] = None
    for m in merges:
        if in_range(row, col, m):
            merge = {"range": range_to_a1(*m), "anchor": (row, col) == (m[0], m[1])}
            break
    has_cache = c.find(q("v")) is not None
    stale = bool(has_formula and (not has_cache or calc_on_load))
    return {
        "ref": rc_to_a1(row, col),
        "locator": make_locator(sheet_index, row, col),
        "row": row,
        "col": col,
        "type": vtype,
        "value": None if (has_formula and not has_cache) else value,
        "has_formula": has_formula,
        "formula": formula_text,
        "formula_kind": (f.get("t", "normal") if has_formula else None),
        "cached_value": (value if has_cache else None) if has_formula else None,
        "stale": stale,
        "style": c.get("s"),
        "merge": merge,
    }


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def _load(path: Union[str, Path]) -> Tuple[XlsxPackage, str]:
    reject_legacy_suffix(path)
    data = Path(path).read_bytes()
    return XlsxPackage.from_bytes(data), compute_document_revision(data)


def inspect_workbook(path: Union[str, Path], locator_limit: int = DEFAULT_LOCATOR_LIMIT) -> Dict[str, Any]:
    pkg, revision = _load(path)
    sheets = load_sheets(pkg)
    calc_flag = full_calc_on_load(pkg)
    shared = load_shared_strings(pkg)
    locators: List[Dict[str, Any]] = []
    truncated = False
    sheet_reports: List[Dict[str, Any]] = []
    for s in sheets:
        if s.kind != "worksheet":
            sheet_reports.append({
                "index": s.index, "name": s.name, "sheet_id": s.sheet_id, "state": s.state, "part": s.part,
                "kind": s.kind, "dimension": None, "merged_cells": [], "tables": [], "array_formulas": [],
                "cell_count": 0, "formula_count": 0,
            })
            continue
        root = parse_xml_safe(pkg.get(s.part))
        merges = sheet_merged(root)
        dim = root.find(q("dimension"))
        cell_count = formula_count = 0
        for r, cc, c in iter_cells(root):
            cell_count += 1
            info = parse_cell(c, r, cc, s.index, shared, merges, calc_flag)
            formula_count += 1 if info["has_formula"] else 0
            if len(locators) < locator_limit:
                locators.append({
                    "locator": info["locator"], "sheet": s.name, "ref": info["ref"],
                    "type": info["type"], "has_formula": info["has_formula"],
                })
            else:
                truncated = True
        sheet_reports.append({
            "index": s.index,
            "name": s.name,
            "sheet_id": s.sheet_id,
            "state": s.state,
            "part": s.part,
            "kind": s.kind,
            "dimension": dim.get("ref") if dim is not None else None,
            "merged_cells": [range_to_a1(*m) for m in merges],
            "tables": [{"name": t.name, "ref": range_to_a1(*t.ref), "header_rows": t.header_rows} for t in sheet_tables(pkg, s)],
            "array_formulas": [range_to_a1(*a) for a in sheet_array_ranges(root)],
            "cell_count": cell_count,
            "formula_count": formula_count,
        })
    return {
        "document_revision": revision,
        "sheets": sheet_reports,
        "calc": {"full_calc_on_load": calc_flag, "has_calc_chain": pkg.has("xl/calcChain.xml")},
        "locators": locators,
        "locators_truncated": truncated,
    }


def read_sheet(
    path: Union[str, Path], sheet: Union[str, int], cell_range: Optional[str] = None
) -> Dict[str, Any]:
    """Doc cac o CO TRONG XML cua sheet (o rong khong duoc tao ra), tuy chon gioi han theo vung."""
    pkg, revision = _load(path)
    info = resolve_sheet(load_sheets(pkg), sheet)
    if info.kind != "worksheet":
        raise UnsupportedSheetError(f"Sheet {info.name!r} la {info.kind}, khong co o de doc.")
    root = parse_xml_safe(pkg.get(info.part))
    merges = sheet_merged(root)
    shared = load_shared_strings(pkg)
    calc_flag = full_calc_on_load(pkg)
    rect = parse_range(cell_range) if cell_range else None
    cells = [
        parse_cell(c, r, cc, info.index, shared, merges, calc_flag)
        for r, cc, c in iter_cells(root)
        if rect is None or in_range(r, cc, rect)
    ]
    return {"document_revision": revision, "sheet": info.name, "sheet_index": info.index, "range": cell_range, "cells": cells}


def read_cell(path: Union[str, Path], sheet: Union[str, int], ref: str) -> Dict[str, Any]:
    """Doc 1 o. O chua co trong XML tra ve type 'blank' (khong loi)."""
    row, col = a1_to_rc(ref)
    result = read_sheet(path, sheet, ref)
    for cell in result["cells"]:
        if (cell["row"], cell["col"]) == (row, col):
            found = cell
            break
    else:
        pkg_sheet_index = result["sheet_index"]
        found = {
            "ref": rc_to_a1(row, col), "locator": make_locator(pkg_sheet_index, row, col), "row": row, "col": col,
            "type": "blank", "value": None, "has_formula": False, "formula": None, "formula_kind": None,
            "cached_value": None, "stale": False, "style": None, "merge": None,
        }
    return {"document_revision": result["document_revision"], "sheet": result["sheet"], "cell": found}


def locate(pkg: XlsxPackage, locator: str) -> Tuple[SheetInfo, int, int]:
    """Giai ma locator `cell_s{idx}_r{row}_c{col}` thanh (SheetInfo, row, col)."""
    sheet_index, row, col = parse_locator(locator)
    sheets = load_sheets(pkg)
    if sheet_index > len(sheets):
        raise LocatorNotFoundError(f"Locator {locator}: workbook chi co {len(sheets)} sheet.")
    return sheets[sheet_index - 1], row, col
