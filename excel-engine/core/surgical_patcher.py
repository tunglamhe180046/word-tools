"""
tools/excel-engine/core/surgical_patcher.py - Sua phau thuat (surgical) tung o trong .xlsx: chi
part worksheet bi sua va `xl/workbook.xml`/calcChain/[Content_Types].xml/rels khi bat buoc; moi
part khac giu nguyen 100% byte (SHA-256 khong doi).

Quy tac ghi:
  - Chuoi: `<c t="inlineStr"><is><t xml:space="preserve">...</t></is></c>` (khong dong den
    sharedStrings). Escape XML do lxml lam; ky tu `_xHHHH_` duoc escape thanh `_x005F_xHHHH_`; ky
    tu dieu khien ASCII khong hop le (XML 1.0) bi TU CHOI; chuoi > 32767 don vi UTF-16 bi TU CHOI
    (hoac cat neu `truncate=True`).
  - Thu tu the con trong `<c>`: `<f>` truoc, roi `<v>`/`<is>`. Thu tu `<c>` tang theo cot, `<row>`
    tang theo dong; dong/o chua co duoc chen dung vi tri.
  - Giu nguyen thuoc tinh style `s="..."` cua o (o moi duoc thua style cua `<row s customFormat>`
    neu co).
  - Fail-closed: khong sua o con nam trong merged range (chi sua anchor top-left), khong sua header
    cua Table, khong sua o trong vung array formula / data table, khong ghi de o goc cua shared
    formula (se lam hong cac o con).
  - Moi lan patch dat `fullCalcOnLoad="1"` trong `<calcPr>` cua `xl/workbook.xml`. Khi sua/xoa mot
    o cong thuc hoac ghi cong thuc moi, `xl/calcChain.xml` bi go khoi goi cung entry
    Override trong `[Content_Types].xml` va Relationship trong `xl/_rels/workbook.xml.rels`, de
    Excel tu dung lai chuoi tinh ma khong bao "repair".
  - `batch_patch`: kiem tra + ap dung toan bo trong bo nho; loi bat ky op nao la huy ca lo, chi
    commit dung 1 lan (atomic) qua Commit Broker (`core/safety_gateway.py`).
  - Luon can `expected_revision` (SHA-256 `document_revision`); khac dia hien tai -> DocumentDriftError.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple, Union

from lxml import etree

from core.cell_addressing import a1_to_rc, in_range, parse_range, range_to_a1, rc_to_a1
from core.inspector import (
    SheetInfo,
    load_sheets,
    locate,
    q,
    resolve_sheet,
    sheet_array_ranges,
    sheet_data,
    sheet_merged,
    sheet_tables,
)
from core.safety_gateway import (
    DocumentDriftError,
    PostconditionVerificationError,
    assert_revision,
    cleanup_job_work_dir,
    issue_job,
    parse_xml_safe,
    request_commit,
    sha256_bytes,
)
from core.xlsx_package import UnsupportedWorkbookError, XlsxPackage

DEFAULT_ACTOR = "excel-engine"
MAX_TEXT_UTF16 = 32767
MAX_FORMULA_CHARS = 8192
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

_ILLEGAL_XML_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿\ud800-\udfff]")
_X_ESCAPE_RE = re.compile(r"_(x[0-9A-Fa-f]{4}_)")
_WORKBOOK_ORDER = [
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
    "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize",
    "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
    "fileRecoveryPr", "webPublishObjects", "extLst",
]


# ---------------------------------------------------------------------------
# Exceptions (fail-closed)
# ---------------------------------------------------------------------------
class PatchRejectedError(RuntimeError):
    """Op bi tu choi truoc khi ghi bat ky thu gi."""


class InvalidValueError(PatchRejectedError):
    pass


class TextTooLongError(InvalidValueError):
    pass


class MergedCellError(PatchRejectedError):
    pass


class TableHeaderError(PatchRejectedError):
    pass


class ArrayFormulaError(PatchRejectedError):
    pass


class SharedFormulaError(PatchRejectedError):
    pass


class DuplicateTargetError(PatchRejectedError):
    pass


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
@dataclass
class CellOp:
    kind: str  # text | number | bool | formula | blank
    value: Any = None
    sheet: Optional[Union[str, int]] = None
    ref: Optional[str] = None  # A1
    locator: Optional[str] = None  # cell_s{idx}_r{row}_c{col}, thay cho sheet+ref


@dataclass
class PatchResult:
    outcome: str
    new_document_revision: str
    job_id: str
    backup_path: Optional[str]
    ops_applied: int
    modified_parts: List[str] = field(default_factory=list)
    removed_parts: List[str] = field(default_factory=list)
    calc_chain_removed: bool = False
    unchanged_part_count: int = 0


def op_from_native(value: Any, sheet: Optional[Union[str, int]] = None, ref: Optional[str] = None,
                   locator: Optional[str] = None) -> CellOp:
    """str -> text, bool -> bool, int/float -> number, None -> blank, {"formula": "..."} -> formula."""
    if value is None:
        kind = "blank"
    elif isinstance(value, bool):
        kind = "bool"
    elif isinstance(value, (int, float)):
        kind = "number"
    elif isinstance(value, str):
        kind = "text"
    elif isinstance(value, dict) and set(value) == {"formula"}:
        return CellOp("formula", value["formula"], sheet, ref, locator)
    else:
        raise InvalidValueError(f"Gia tri khong ho tro: {value!r}")
    return CellOp(kind, value, sheet, ref, locator)


def op_from_dict(d: Dict[str, Any]) -> CellOp:
    """{"sheet","cell"|"ref"|"target_id", "value" | "formula", ["type"]}."""
    if not isinstance(d, dict):
        raise InvalidValueError(f"Op phai la object: {d!r}")
    sheet, ref, locator = d.get("sheet"), d.get("cell") or d.get("ref"), d.get("target_id")
    if locator is None and (sheet is None or ref is None):
        raise InvalidValueError(f"Op thieu dia chi (can target_id hoac sheet+cell): {d!r}")
    if "formula" in d:
        return CellOp("formula", d["formula"], sheet, ref, locator)
    if "value" not in d:
        raise InvalidValueError(f"Op thieu 'value' hoac 'formula': {d!r}")
    declared = d.get("type")
    if declared is None:
        return op_from_native(d["value"], sheet, ref, locator)
    if declared not in ("text", "number", "bool", "blank"):
        raise InvalidValueError(f"type khong hop le: {declared!r}")
    return CellOp(declared, d["value"], sheet, ref, locator)


# ---------------------------------------------------------------------------
# Value validation & encoding
# ---------------------------------------------------------------------------
def _utf16_len(text: str) -> int:
    return len(text.encode("utf-16-le")) // 2


def sanitize_text(text: Any, truncate: bool = False) -> str:
    if not isinstance(text, str):
        raise InvalidValueError(f"Gia tri chuoi phai la str: {text!r}")
    bad = _ILLEGAL_XML_RE.search(text)
    if bad:
        raise InvalidValueError(f"Ky tu dieu khien khong hop le U+{ord(bad.group()):04X} trong chuoi.")
    if _utf16_len(text) > MAX_TEXT_UTF16:
        if not truncate:
            raise TextTooLongError(f"Chuoi dai {_utf16_len(text)} > {MAX_TEXT_UTF16} ky tu.")
        # cat 1 lan theo don vi UTF-16; 'ignore' bo surrogate mo coi neu diem cat roi giua 1 cap
        text = text.encode("utf-16-le")[: MAX_TEXT_UTF16 * 2].decode("utf-16-le", errors="ignore")
    return _X_ESCAPE_RE.sub(r"_x005F_\1", text)


def sanitize_formula(formula: Any) -> str:
    if not isinstance(formula, str):
        raise InvalidValueError(f"Cong thuc phai la str: {formula!r}")
    body = formula[1:] if formula.startswith("=") else formula
    if not body.strip():
        raise InvalidValueError("Cong thuc rong.")
    if _ILLEGAL_XML_RE.search(body):
        raise InvalidValueError("Cong thuc chua ky tu dieu khien khong hop le.")
    if len(body) > MAX_FORMULA_CHARS:
        raise InvalidValueError(f"Cong thuc dai {len(body)} > {MAX_FORMULA_CHARS} ky tu.")
    return body


def _number_text(value: Any) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise InvalidValueError(f"Gia tri so khong hop le: {value!r}")
    if isinstance(value, float) and not math.isfinite(value):
        raise InvalidValueError(f"So khong huu han (NaN/Inf): {value!r}")
    return repr(value)


# ---------------------------------------------------------------------------
# XML mutation helpers
# ---------------------------------------------------------------------------
def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _materialize_refs(root: etree._Element) -> None:
    """Ghi tuong minh `r` len moi <row>/<c> con thieu (nghia khong doi) de viec chen giu duoc thu tu."""
    prev_row = 0
    for row in sheet_data(root).findall(q("row")):
        if not row.get("r"):
            row.set("r", str(prev_row + 1))
        prev_row = int(row.get("r"))
        prev_col = 0
        for c in row.findall(q("c")):
            if not c.get("r"):
                c.set("r", rc_to_a1(prev_row, prev_col + 1))
            prev_col = a1_to_rc(c.get("r"))[1]


def _get_or_create_cell(data: etree._Element, row: int, col: int) -> etree._Element:
    row_el: Optional[etree._Element] = None
    before_row: Optional[etree._Element] = None
    for candidate in data.findall(q("row")):
        rn = int(candidate.get("r"))
        if rn == row:
            row_el = candidate
            break
        if rn > row:
            before_row = candidate
            break
    if row_el is None:
        row_el = etree.Element(q("row"))
        row_el.set("r", str(row))
        if before_row is not None:
            before_row.addprevious(row_el)
        else:
            data.append(row_el)

    before_cell: Optional[etree._Element] = None
    for cell in row_el.findall(q("c")):
        cn = a1_to_rc(cell.get("r"))[1]
        if cn == col:
            return cell
        if cn > col:
            before_cell = cell
            break
    cell = etree.Element(q("c"))
    cell.set("r", rc_to_a1(row, col))
    if row_el.get("customFormat") in ("1", "true") and row_el.get("s"):
        cell.set("s", row_el.get("s"))
    if before_cell is not None:
        before_cell.addprevious(cell)
    else:
        row_el.append(cell)
    if "spans" in row_el.attrib:
        del row_el.attrib["spans"]  # toi uu tuy chon, co the sai sau khi chen o
    return cell


def _write_content(cell: etree._Element, op: CellOp, truncate: bool) -> None:
    """Thay noi dung o, giu nguyen `s` va cac thuoc tinh khac (chi bo `t`, `vm`, `cm`)."""
    text = formula = number = None
    if op.kind == "text":
        text = sanitize_text(op.value, truncate)
    elif op.kind == "formula":
        formula = sanitize_formula(op.value)
    elif op.kind == "number":
        number = _number_text(op.value)
    elif op.kind == "bool":
        if not isinstance(op.value, bool):
            raise InvalidValueError(f"Gia tri bool khong hop le: {op.value!r}")
    elif op.kind != "blank":
        raise InvalidValueError(f"Loai op khong hop le: {op.kind!r}")

    for child in list(cell):
        cell.remove(child)
    for attr in ("t", "vm", "cm"):
        if attr in cell.attrib:
            del cell.attrib[attr]

    if text is not None:
        cell.set("t", "inlineStr")
        is_el = etree.SubElement(cell, q("is"))
        t_el = etree.SubElement(is_el, q("t"))
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_el.text = text
    elif formula is not None:
        f_el = etree.SubElement(cell, q("f"))
        f_el.text = formula
    elif number is not None:
        etree.SubElement(cell, q("v")).text = number
    elif op.kind == "bool":
        cell.set("t", "b")
        etree.SubElement(cell, q("v")).text = "1" if op.value else "0"


def _expand_dimension(root: etree._Element, touched: Iterable[Tuple[int, int]]) -> None:
    dim = root.find(q("dimension"))
    if dim is None or not dim.get("ref"):
        return
    r1, c1, r2, c2 = parse_range(dim.get("ref"))
    for r, c in touched:
        r1, c1, r2, c2 = min(r1, r), min(c1, c), max(r2, r), max(c2, c)
    dim.set("ref", range_to_a1(r1, c1, r2, c2))


def _ensure_full_calc_on_load(pkg: XlsxPackage) -> None:
    root = parse_xml_safe(pkg.get("xl/workbook.xml"))
    calc = root.find(q("calcPr"))
    if calc is not None and calc.get("fullCalcOnLoad") == "1":
        return
    if calc is None:
        calc = etree.Element(q("calcPr"))
        after = set(_WORKBOOK_ORDER[_WORKBOOK_ORDER.index("calcPr") + 1:])
        anchor = next((el for el in root if isinstance(el.tag, str) and etree.QName(el).localname in after), None)
        if anchor is not None:
            anchor.addprevious(calc)
        else:
            root.append(calc)
    calc.set("fullCalcOnLoad", "1")
    pkg.set("xl/workbook.xml", _serialize(root))


def _drop_calc_chain(pkg: XlsxPackage) -> bool:
    if not pkg.has("xl/calcChain.xml"):
        return False
    pkg.remove("xl/calcChain.xml")
    ct = parse_xml_safe(pkg.get("[Content_Types].xml"))
    for el in list(ct):
        if el.tag == f"{{{NS_CT}}}Override" and el.get("PartName") == "/xl/calcChain.xml":
            ct.remove(el)
    pkg.set("[Content_Types].xml", _serialize(ct))
    rels_part = "xl/_rels/workbook.xml.rels"
    if pkg.has(rels_part):
        rels = parse_xml_safe(pkg.get(rels_part))
        for el in list(rels):
            if isinstance(el.tag, str) and (el.get("Type", "").endswith("/calcChain") or el.get("Target", "").endswith("calcChain.xml")):
                rels.remove(el)
        pkg.set(rels_part, _serialize(rels))
    return True


# ---------------------------------------------------------------------------
# Validation against structural constraints
# ---------------------------------------------------------------------------
def _validate_target(ref: str, row: int, col: int, merges, arrays, tables) -> None:
    for m in merges:
        if in_range(row, col, m) and (row, col) != (m[0], m[1]):
            raise MergedCellError(
                f"O {ref} nam trong merged range {range_to_a1(*m)} nhung khong phai o anchor {rc_to_a1(m[0], m[1])}."
            )
    for a in arrays:
        if in_range(row, col, a):
            raise ArrayFormulaError(f"O {ref} nam trong vung array formula {range_to_a1(*a)}.")
    for t in tables:
        if t.header_rows > 0 and in_range(row, col, (t.ref[0], t.ref[1], t.ref[0] + t.header_rows - 1, t.ref[3])):
            raise TableHeaderError(f"O {ref} la header cua Table {t.name!r} ({range_to_a1(*t.ref)}).")


# ---------------------------------------------------------------------------
# In-memory patch (khong ghi dia)
# ---------------------------------------------------------------------------
def apply_ops_to_package(pkg: XlsxPackage, ops: Sequence[CellOp], truncate: bool = False) -> Dict[str, Any]:
    """Kiem tra va ap dung TOAN BO ops len goi trong bo nho. Loi o bat ky op nao -> nem, va goi
    `pkg` khong con ve trang thai dung (caller phai bo di); khong co thu gi duoc ghi ra dia."""
    if not ops:
        raise InvalidValueError("Danh sach op rong.")
    sheets = load_sheets(pkg)

    resolved: List[Tuple[SheetInfo, int, int, CellOp]] = []
    seen = set()
    for op in ops:
        if op.locator:
            sheet, row, col = locate(pkg, op.locator)
        else:
            sheet = resolve_sheet(sheets, op.sheet)
            row, col = a1_to_rc(op.ref)
        if sheet.kind != "worksheet":
            raise PatchRejectedError(f"Sheet {sheet.name!r} la {sheet.kind}, khong co o de ghi.")
        key = (sheet.index, row, col)
        if key in seen:
            raise DuplicateTargetError(f"O {rc_to_a1(row, col)} trong sheet {sheet.name!r} bi ghi nhieu lan trong cung lo.")
        seen.add(key)
        resolved.append((sheet, row, col, op))

    by_sheet: Dict[int, List[Tuple[SheetInfo, int, int, CellOp]]] = {}
    for item in resolved:
        by_sheet.setdefault(item[0].index, []).append(item)

    formula_touched = False
    for group in by_sheet.values():
        sheet = group[0][0]
        root = parse_xml_safe(pkg.get(sheet.part))
        merges, arrays, tables = sheet_merged(root), sheet_array_ranges(root), sheet_tables(pkg, sheet)
        for _, row, col, _op in group:
            _validate_target(rc_to_a1(row, col), row, col, merges, arrays, tables)

        _materialize_refs(root)
        data = sheet_data(root)
        for _, row, col, op in group:
            cell = _get_or_create_cell(data, row, col)
            old_f = cell.find(q("f"))
            if old_f is not None:
                if old_f.get("t") == "shared" and old_f.get("ref"):
                    raise SharedFormulaError(
                        f"O {rc_to_a1(row, col)} la goc cua shared formula {old_f.get('ref')}; ghi de se lam hong cac o con."
                    )
                formula_touched = True
            if op.kind == "formula":
                formula_touched = True
            _write_content(cell, op, truncate)
        _expand_dimension(root, [(r, c) for _, r, c, _o in group])
        pkg.set(sheet.part, _serialize(root))

    _ensure_full_calc_on_load(pkg)
    calc_chain_removed = _drop_calc_chain(pkg) if formula_touched else False
    return {"ops_applied": len(resolved), "calc_chain_removed": calc_chain_removed}


# ---------------------------------------------------------------------------
# Commit qua Commit Broker
# ---------------------------------------------------------------------------
def batch_patch(
    path: Union[str, Path],
    ops: Sequence[CellOp],
    expected_revision: str,
    work_dir: Optional[Union[str, Path]] = None,
    allowed_roots: Optional[List[str]] = None,
    actor: str = DEFAULT_ACTOR,
    truncate: bool = False,
) -> PatchResult:
    path = Path(path).resolve()
    if path.suffix.lower() != ".xlsx":
        raise UnsupportedWorkbookError(f"Chi patch duoc file .xlsx (nhan {path.suffix!r}).")
    work_dir = Path(work_dir).resolve() if work_dir else path.parent
    data = path.read_bytes()
    assert_revision(data, expected_revision)

    pkg = XlsxPackage.from_bytes(data)
    summary = apply_ops_to_package(pkg, ops, truncate=truncate)
    new_bytes = pkg.to_bytes()
    unchanged = len(pkg.unchanged_hashes())

    ticket = issue_job(path, actor=actor, allowed_roots=allowed_roots or [str(path.parent)], work_dir=work_dir)
    keep_work_dir = False
    try:
        if ticket.baseline_sha256 != sha256_bytes(data):
            raise DocumentDriftError(expected_revision, f"sha256:{ticket.baseline_sha256}")
        (Path(ticket.work_dir) / f"candidate{path.suffix}").write_bytes(new_bytes)
        result = request_commit(ticket.job_id, work_dir=work_dir)
    except PostconditionVerificationError:
        keep_work_dir = True  # giu staged de phuc hoi thu cong
        raise
    finally:
        if not keep_work_dir:
            cleanup_job_work_dir(ticket.work_dir)

    return PatchResult(
        outcome="patched",
        new_document_revision=f"sha256:{result['sha256']}",
        job_id=ticket.job_id,
        backup_path=result.get("backup_path"),
        ops_applied=summary["ops_applied"],
        modified_parts=sorted(pkg.modified_parts()),
        removed_parts=sorted(pkg.removed_parts()),
        calc_chain_removed=summary["calc_chain_removed"],
        unchanged_part_count=unchanged,
    )


def patch_cell(
    path: Union[str, Path],
    sheet: Optional[Union[str, int]],
    ref: Optional[str],
    value: Any,
    expected_revision: str,
    kind: Optional[str] = None,
    locator: Optional[str] = None,
    **kwargs: Any,
) -> PatchResult:
    """`kind` None -> suy ra tu kieu cua `value` (xem op_from_native)."""
    op = op_from_native(value, sheet, ref, locator) if kind is None else CellOp(kind, value, sheet, ref, locator)
    return batch_patch(path, [op], expected_revision, **kwargs)


def patch_range(
    path: Union[str, Path],
    sheet: Union[str, int],
    cell_range: str,
    values: Sequence[Sequence[Any]],
    expected_revision: str,
    **kwargs: Any,
) -> PatchResult:
    """Ghi ma tran 2D `values` (kich thuoc phai khop vung) - 1 lo atomic."""
    r1, c1, r2, c2 = parse_range(cell_range)
    if len(values) != r2 - r1 + 1 or any(not isinstance(row, (list, tuple)) or len(row) != c2 - c1 + 1 for row in values):
        raise InvalidValueError(f"Ma tran values khong khop vung {cell_range} ({r2 - r1 + 1}x{c2 - c1 + 1}).")
    ops = [
        op_from_native(values[r - r1][c - c1], sheet, rc_to_a1(r, c))
        for r in range(r1, r2 + 1)
        for c in range(c1, c2 + 1)
    ]
    return batch_patch(path, ops, expected_revision, **kwargs)
