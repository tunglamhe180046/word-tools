import pytest
from conftest import build_minimal_xlsx, file_revision, part_names, read_part
from lxml import etree

from core.inspector import read_cell, read_sheet
from core.safety_gateway import list_backups
from core.surgical_patcher import (
    CellOp,
    DuplicateTargetError,
    InvalidValueError,
    MergedCellError,
    batch_patch,
    op_from_dict,
    patch_cell,
    patch_range,
)

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _values(path, sheet):
    return {c["ref"]: c["value"] for c in read_sheet(path, sheet)["cells"]}


def test_batch_multi_sheet_is_one_atomic_commit(xlsx_path):
    ops = [
        CellOp("text", "A", "Data", "A2"),
        CellOp("number", 1, "Data", "B2"),
        CellOp("text", "S2", "Second", "C1"),
        CellOp("bool", False, "Second", "D1"),
    ]
    result = batch_patch(xlsx_path, ops, file_revision(xlsx_path))
    assert result.ops_applied == 4 and result.outcome == "patched"
    assert result.new_document_revision == file_revision(xlsx_path)
    assert len(list_backups(xlsx_path)) == 1  # dung 1 commit
    assert _values(xlsx_path, "Data")["A2"] == "A" and _values(xlsx_path, "Second")["C1"] == "S2"


def test_batch_aborts_entirely_on_one_bad_op(xlsx_path, tmp_path):
    original = xlsx_path.read_bytes()
    ops = [
        CellOp("text", "fine", "Data", "A2"),
        CellOp("text", "fine too", "Second", "C1"),
        CellOp("text", "boom", "Data", "B6"),  # o merged khong phai anchor
    ]
    with pytest.raises(MergedCellError):
        batch_patch(xlsx_path, ops, file_revision(xlsx_path))
    assert xlsx_path.read_bytes() == original
    assert list_backups(xlsx_path) == []
    assert not (tmp_path / ".jarvis" / "backups").exists()


@pytest.mark.parametrize("bad_op", [
    CellOp("text", "x" * 40000, "Data", "A3"),
    CellOp("number", float("nan"), "Data", "A3"),
    CellOp("weird", 1, "Data", "A3"),
    CellOp("bool", "yes", "Data", "A3"),
])
def test_batch_aborts_on_invalid_value_in_later_op(xlsx_path, bad_op):
    original = xlsx_path.read_bytes()
    with pytest.raises(InvalidValueError):
        batch_patch(xlsx_path, [CellOp("text", "ok", "Data", "A2"), bad_op], file_revision(xlsx_path))
    assert xlsx_path.read_bytes() == original


def test_batch_rejects_duplicates_empty_and_unknown_sheet(xlsx_path):
    rev = file_revision(xlsx_path)
    with pytest.raises(DuplicateTargetError):
        batch_patch(xlsx_path, [CellOp("text", "a", "Data", "A2"), CellOp("text", "b", "Data", "a2")], rev)
    with pytest.raises(InvalidValueError):
        batch_patch(xlsx_path, [], rev)
    with pytest.raises(LookupError):
        batch_patch(xlsx_path, [CellOp("text", "a", "Ghost", "A1")], rev)
    assert file_revision(xlsx_path) == rev


def test_patch_range_writes_matrix_and_validates_shape(xlsx_path):
    rev = file_revision(xlsx_path)
    with pytest.raises(InvalidValueError):
        patch_range(xlsx_path, "Data", "A10:B11", [["only-one-row", 1]], rev)
    with pytest.raises(InvalidValueError):
        patch_range(xlsx_path, "Data", "A10:B11", [["a"], ["b"]], rev)
    assert file_revision(xlsx_path) == rev
    patch_range(xlsx_path, "Data", "A10:C11", [["a", 1, True], [None, 2.5, {"formula": "SUM(B10:B11)"}]], rev)
    vals = _values(xlsx_path, "Data")
    assert vals["A10"] == "a" and vals["B10"] == 1 and vals["C10"] is True and vals["B11"] == 2.5
    assert read_cell(xlsx_path, "Data", "C11")["cell"]["formula"] == "SUM(B10:B11)"


def test_patch_range_over_merged_area_is_rejected_atomically(xlsx_path):
    original = xlsx_path.read_bytes()
    with pytest.raises(MergedCellError):
        patch_range(xlsx_path, "Data", "A6:B6", [["ok-anchor", "not-ok"]], file_revision(xlsx_path))
    assert xlsx_path.read_bytes() == original


def test_op_from_dict_forms(xlsx_path):
    assert op_from_dict({"sheet": "Data", "cell": "A2", "value": "x"}).kind == "text"
    assert op_from_dict({"sheet": "Data", "ref": "A2", "value": 3}).kind == "number"
    assert op_from_dict({"sheet": "Data", "cell": "A2", "formula": "=A1"}).kind == "formula"
    assert op_from_dict({"target_id": "cell_s1_r2_c1", "value": None}).kind == "blank"
    assert op_from_dict({"sheet": "Data", "cell": "A2", "value": "5", "type": "text"}).value == "5"
    for bad in ({"sheet": "Data", "value": "x"}, {"sheet": "Data", "cell": "A2"}, "nope",
                {"sheet": "Data", "cell": "A2", "value": 1, "type": "matrix"}):
        with pytest.raises(InvalidValueError):
            op_from_dict(bad)


# ---- calcChain / fullCalcOnLoad ------------------------------------------------
def _calc(path):
    root = etree.fromstring(read_part(path, "xl/workbook.xml"))
    return root.xpath("//m:calcPr", namespaces=NS)


def test_value_patch_sets_full_calc_and_keeps_calc_chain(xlsx_path):
    patch_cell(xlsx_path, "Data", "A2", "plain", file_revision(xlsx_path))
    assert _calc(xlsx_path)[0].get("fullCalcOnLoad") == "1"
    assert _calc(xlsx_path)[0].get("calcId") == "191029"  # thuoc tinh cu con nguyen
    assert "xl/calcChain.xml" in part_names(xlsx_path)  # khong dung toi o cong thuc


@pytest.mark.parametrize("op", [
    CellOp("formula", "1+1", "Data", "A2"),  # ghi cong thuc moi
    CellOp("number", 1, "Data", "B4"),  # ghi de o cong thuc bang gia tri
    CellOp("blank", None, "Data", "C4"),  # xoa o cong thuc
])
def test_formula_edits_drop_calc_chain_consistently(xlsx_path, op):
    result = batch_patch(xlsx_path, [op], file_revision(xlsx_path))
    assert result.calc_chain_removed is True and result.removed_parts == ["xl/calcChain.xml"]
    assert "xl/calcChain.xml" not in part_names(xlsx_path)
    ct = read_part(xlsx_path, "[Content_Types].xml").decode()
    rels = read_part(xlsx_path, "xl/_rels/workbook.xml.rels").decode()
    assert "calcChain" not in ct and "calcChain" not in rels
    assert "sheet1.xml" in ct and "sharedStrings" in rels  # cac muc khac con nguyen
    assert set(result.modified_parts) == {"xl/worksheets/sheet1.xml", "xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}


def test_no_calc_chain_workbook_and_missing_calc_pr(tmp_path):
    path = build_minimal_xlsx(tmp_path / "nocc.xlsx", calc_chain=False, calc_pr=False)
    result = patch_cell(path, "Data", "B4", 1, file_revision(path))
    assert result.calc_chain_removed is False
    calc = _calc(path)
    assert len(calc) == 1 and calc[0].get("fullCalcOnLoad") == "1"
    order = [etree.QName(e).localname for e in etree.fromstring(read_part(path, "xl/workbook.xml"))]
    assert order == ["workbookPr", "sheets", "definedNames", "calcPr"]  # dung vi tri theo schema


def test_full_calc_already_set_leaves_workbook_part_untouched(tmp_path):
    path = build_minimal_xlsx(tmp_path / "fc.xlsx", full_calc_on_load=True)
    before = read_part(path, "xl/workbook.xml")
    result = patch_cell(path, "Data", "A2", "z", file_revision(path))
    assert read_part(path, "xl/workbook.xml") == before
    assert "xl/workbook.xml" not in result.modified_parts
