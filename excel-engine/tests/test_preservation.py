import io
import zipfile

import pytest
from conftest import build_minimal_xlsx, file_revision, part_hashes, part_names

from core.inspector import read_sheet
from core.surgical_patcher import CellOp, apply_ops_to_package, batch_patch, patch_cell
from core.xlsx_package import XlsxPackage

SHEET1 = "xl/worksheets/sheet1.xml"


def _unchanged_expected(before, after, touched):
    """Moi part khong nam trong `touched` phai co SHA-256 giong het truoc/sau."""
    for name, digest in before.items():
        if name in touched:
            continue
        assert name in after, f"part {name} bi mat"
        assert after[name] == digest, f"part {name} bi thay doi"


def test_value_patch_touches_only_sheet_and_workbook(xlsx_path):
    before, names_before = part_hashes(xlsx_path), part_names(xlsx_path)
    patch_cell(xlsx_path, "Data", "A2", "changed", file_revision(xlsx_path))
    after = part_hashes(xlsx_path)
    _unchanged_expected(before, after, {SHEET1, "xl/workbook.xml"})
    assert after[SHEET1] != before[SHEET1] and after["xl/workbook.xml"] != before["xl/workbook.xml"]
    assert part_names(xlsx_path) == names_before  # thu tu entry giu nguyen


def test_formula_patch_touches_exact_set_and_preserves_order(xlsx_path):
    before, names_before = part_hashes(xlsx_path), part_names(xlsx_path)
    patch_cell(xlsx_path, "Data", "C4", {"formula": "B4+1"}, file_revision(xlsx_path))
    after = part_hashes(xlsx_path)
    touched = {SHEET1, "xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/calcChain.xml"}
    _unchanged_expected(before, after, touched)
    assert [n for n in names_before if n != "xl/calcChain.xml"] == part_names(xlsx_path)
    # cac part nhay cam khong bi cham: anh, customXml, sharedStrings, styles, table, sheet2, rels cua sheet
    for name in ("xl/media/image1.png", "customXml/item1.xml", "xl/sharedStrings.xml", "xl/styles.xml",
                 "xl/tables/table1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/_rels/sheet1.xml.rels", "docProps/core.xml"):
        assert after[name] == before[name], name


def test_untouched_cells_are_semantically_identical(xlsx_path):
    before = {c["ref"]: c for c in read_sheet(xlsx_path, "Data")["cells"]}
    patch_cell(xlsx_path, "Data", "A2", "changed", file_revision(xlsx_path))
    after = {c["ref"]: c for c in read_sheet(xlsx_path, "Data")["cells"]}
    assert set(before) == set(after)
    for ref, cell in before.items():
        if ref == "A2":
            continue
        for key in ("type", "formula", "style", "merge", "row", "col"):
            assert after[ref][key] == cell[key], (ref, key)
        if not cell["has_formula"]:
            assert after[ref]["value"] == cell["value"], ref
    # cong thuc giu nguyen text, cache van con (chi danh dau stale vi fullCalcOnLoad)
    assert after["B4"]["formula"] == "SUM(B2:B3)" and after["B4"]["cached_value"] == 12 and after["B4"]["stale"] is True


def test_in_memory_package_roundtrip_is_byte_preserving_for_unmodified_parts(xlsx_path):
    data = xlsx_path.read_bytes()
    pkg = XlsxPackage.from_bytes(data)
    assert pkg.modified_parts() == [] and pkg.removed_parts() == []
    rebuilt = pkg.to_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as a, zipfile.ZipFile(io.BytesIO(rebuilt)) as b:
        assert a.namelist() == b.namelist()
        for info in a.infolist():
            assert a.read(info.filename) == b.read(info.filename)
            binfo = b.getinfo(info.filename)
            assert (binfo.date_time, binfo.compress_type) == (info.date_time, info.compress_type)


def test_package_tracks_modified_and_removed(xlsx_path):
    pkg = XlsxPackage.open(xlsx_path)
    apply_ops_to_package(pkg, [CellOp("formula", "1+1", "Data", "A2")])
    assert set(pkg.modified_parts()) == {SHEET1, "xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}
    assert pkg.removed_parts() == ["xl/calcChain.xml"]
    assert len(pkg.unchanged_hashes()) == len(pkg.names) - 4


def test_patched_workbook_opens_in_openpyxl(xlsx_path):
    openpyxl = pytest.importorskip("openpyxl")
    batch_patch(xlsx_path, [
        CellOp("text", "Zed", "Data", "A2"),
        CellOp("number", 42, "Data", "B2"),
        CellOp("formula", "B2+1", "Data", "C2"),
        CellOp("text", "s2", "Second", "B1"),
    ], file_revision(xlsx_path))
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["Data"]
    assert ws["A2"].value == "Zed" and ws["B2"].value == 42 and ws["C2"].value == "=B2+1"
    assert "A6:B6" in [str(r) for r in ws.merged_cells.ranges]
    assert wb["Second"]["B1"].value == "s2"


def test_plain_workbook_without_extras_still_preserves(tmp_path):
    path = build_minimal_xlsx(tmp_path / "p.xlsx", calc_chain=False, merge=False, table=False, array=False)
    before = part_hashes(path)
    patch_cell(path, "Second", "B1", "x", file_revision(path))
    _unchanged_expected(before, part_hashes(path), {"xl/worksheets/sheet2.xml", "xl/workbook.xml"})
