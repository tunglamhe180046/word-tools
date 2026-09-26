import time

import pytest
from conftest import build_minimal_xlsx, file_revision, read_part
from lxml import etree

from core.cell_addressing import CellAddressError, make_locator
from core.inspector import read_cell
from core.surgical_patcher import (
    ArrayFormulaError,
    InvalidValueError,
    MergedCellError,
    SharedFormulaError,
    TableHeaderError,
    TextTooLongError,
    patch_cell,
    sanitize_text,
)

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _patch(path, sheet, ref, value, **kw):
    return patch_cell(path, sheet, ref, value, file_revision(path), **kw)


def _sheet(path, part="xl/worksheets/sheet1.xml"):
    return etree.fromstring(read_part(path, part))


def _cell(path, ref, part="xl/worksheets/sheet1.xml"):
    found = _sheet(path, part).xpath(f'//m:c[@r="{ref}"]', namespaces=NS)
    assert len(found) == 1, ref
    return found[0]


def test_write_text_uses_inline_str_and_preserves_style(xlsx_path):
    _patch(xlsx_path, "Data", "A2", "Carol")
    c = _cell(xlsx_path, "A2")
    assert c.get("t") == "inlineStr" and c.get("s") == "2"
    t = c.xpath("m:is/m:t", namespaces=NS)[0]
    assert t.text == "Carol" and t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert read_cell(xlsx_path, "Data", "A2")["cell"]["value"] == "Carol"


def test_overwrite_shared_string_cell_and_numeric_cell(xlsx_path):
    _patch(xlsx_path, "Data", "A4", "  padded  ")  # o tro toi sharedStrings
    assert read_cell(xlsx_path, "Data", "A4")["cell"]["value"] == "  padded  "
    _patch(xlsx_path, "Data", "B3", 99.5)
    c = _cell(xlsx_path, "B3")
    assert c.get("t") is None and c.xpath("m:v", namespaces=NS)[0].text == "99.5"
    _patch(xlsx_path, "Data", "B3", True)
    c = _cell(xlsx_path, "B3")
    assert c.get("t") == "b" and c.xpath("m:v", namespaces=NS)[0].text == "1"


def test_blank_keeps_style_and_removes_content(xlsx_path):
    _patch(xlsx_path, "Data", "B2", None)
    c = _cell(xlsx_path, "B2")
    assert c.get("s") == "3" and len(c) == 0 and c.get("t") is None


def test_xml_escape_and_x_escape(xlsx_path):
    nasty = '<a href="x">&amp; _x0041_ _x005F_ \t line\nbreak</a>'
    _patch(xlsx_path, "Data", "A2", nasty)
    raw = read_part(xlsx_path, "xl/worksheets/sheet1.xml").decode()
    assert "<a href" not in raw and "&lt;a href" in raw
    stored = _cell(xlsx_path, "A2").xpath("m:is/m:t", namespaces=NS)[0].text
    assert "_x005F_x0041_" in stored  # `_x0041_` literal -> `_x005F_x0041_`
    assert stored.count("_x005F_") == 2  # ca `_x005F_` co san cung duoc escape: `_x005F_x005F_`


def test_sanitize_text_rules():
    assert sanitize_text("plain_x12") == "plain_x12"  # khong du 4 hex + '_'
    assert sanitize_text("_x00Zz_") == "_x00Zz_"
    assert sanitize_text("é中\U0001F600 ok") == "é中\U0001F600 ok"
    with pytest.raises(InvalidValueError):
        sanitize_text(123)


@pytest.mark.parametrize("bad", ["a\x00b", "a\x01b", "\x08", "\x0b", "\x0c", "\x1f", "￾", "￿", "\ud800"])
def test_illegal_control_chars_rejected(xlsx_path, bad):
    before = file_revision(xlsx_path)
    with pytest.raises(InvalidValueError):
        _patch(xlsx_path, "Data", "A2", bad)
    assert file_revision(xlsx_path) == before


def test_text_length_limit(xlsx_path):
    _patch(xlsx_path, "Data", "A2", "x" * 32767)  # dung gioi han: OK
    with pytest.raises(TextTooLongError):
        _patch(xlsx_path, "Data", "A2", "x" * 32768)
    # emoji la 2 don vi UTF-16: 16384 emoji = 32768 don vi -> vuot
    with pytest.raises(TextTooLongError):
        _patch(xlsx_path, "Data", "A2", "\U0001F600" * 16384)
    _patch(xlsx_path, "Data", "A2", "y" * 40000, truncate=True)
    assert len(read_cell(xlsx_path, "Data", "A2")["cell"]["value"]) == 32767
    # cat theo emoji khong bao gio cat doi mot cap surrogate
    _patch(xlsx_path, "Data", "A3", "a" + "\U0001F600" * 20000, truncate=True)
    value = read_cell(xlsx_path, "Data", "A3")["cell"]["value"]
    assert len(value.encode("utf-16-le")) // 2 <= 32767 and value.endswith("\U0001F600")


def test_number_validation(xlsx_path):
    for bad in (float("nan"), float("inf"), float("-inf")):
        with pytest.raises(InvalidValueError):
            _patch(xlsx_path, "Data", "B2", bad)
    with pytest.raises(InvalidValueError):
        _patch(xlsx_path, "Data", "B2", [1, 2])


def test_formula_write_has_f_first_no_value_and_is_stale(xlsx_path):
    _patch(xlsx_path, "Data", "B3", {"formula": "=SUM(B1:B2)+1"})
    c = _cell(xlsx_path, "B3")
    assert [etree.QName(x).localname for x in c] == ["f"]
    assert c[0].text == "SUM(B1:B2)+1"
    got = read_cell(xlsx_path, "Data", "B3")["cell"]
    assert got["formula"] == "SUM(B1:B2)+1" and got["cached_value"] is None and got["stale"] is True


def test_formula_validation(xlsx_path):
    for bad in ("=", "   ", "=" + "A" * 8193, "=A1\x00"):
        with pytest.raises(InvalidValueError):
            _patch(xlsx_path, "Data", "B3", {"formula": bad})


def test_text_starting_with_equals_stays_text(xlsx_path):
    _patch(xlsx_path, "Data", "A2", "=1+1")
    c = _cell(xlsx_path, "A2")
    assert c.get("t") == "inlineStr" and not c.xpath("m:f", namespaces=NS)


def test_insert_new_cells_and_rows_sorted(xlsx_path):
    _patch(xlsx_path, "Data", "D2", "mid-row")  # o moi, dong co san (sau B, truoc het)
    _patch(xlsx_path, "Data", "C2", "before-D")  # chen truoc D2, sau B2
    _patch(xlsx_path, "Data", "A5", "new-row-5")  # dong moi giua 4 va 6
    _patch(xlsx_path, "Data", "A12", "tail")  # dong moi cuoi
    root = _sheet(xlsx_path)
    rows = [int(r.get("r")) for r in root.xpath("//m:row", namespaces=NS)]
    assert rows == sorted(rows) and 5 in rows and 12 in rows
    r2 = root.xpath('//m:row[@r="2"]/m:c', namespaces=NS)
    assert [c.get("r") for c in r2] == ["A2", "B2", "C2", "D2"]
    assert "spans" not in root.xpath('//m:row[@r="2"]', namespaces=NS)[0].attrib
    # dimension duoc mo rong de bao ca A12
    assert root.xpath("//m:dimension", namespaces=NS)[0].get("ref") == "A1:E12"


def test_insert_row_before_first_row(tmp_path):
    path = build_minimal_xlsx(tmp_path / "b.xlsx")
    _patch(path, "Second", "A1", "x")  # o co san
    _patch(path, "Second", "C1", "y")
    root = _sheet(path, "xl/worksheets/sheet2.xml")
    assert [c.get("r") for c in root.xpath('//m:row[@r="1"]/m:c', namespaces=NS)] == ["A1", "C1"]


def test_patch_by_locator(xlsx_path):
    loc = make_locator(1, 2, 1)  # A2
    patch_cell(xlsx_path, None, None, "via-locator", file_revision(xlsx_path), locator=loc)
    assert read_cell(xlsx_path, "Data", "A2")["cell"]["value"] == "via-locator"


def test_new_cell_inherits_row_style(tmp_path):
    path = build_minimal_xlsx(tmp_path / "r.xlsx")
    # bien row 6 thanh row co style mac dinh, roi chen o moi
    import zipfile
    with zipfile.ZipFile(path) as zf:
        items = [(n, zf.read(n)) for n in zf.namelist()]
    with zipfile.ZipFile(path, "w") as zf:
        for n, d in items:
            if n == "xl/worksheets/sheet2.xml":
                d = d.replace(b'<row r="1">', b'<row r="1" s="2" customFormat="1">')
            zf.writestr(n, d)
    _patch(path, "Second", "B1", "styled")
    assert _cell(path, "B1", "xl/worksheets/sheet2.xml").get("s") == "2"


def test_rows_and_cells_without_r_are_materialized(tmp_path):
    import zipfile
    path = build_minimal_xlsx(tmp_path / "n.xlsx")
    with zipfile.ZipFile(path) as zf:
        items = [(n, zf.read(n)) for n in zf.namelist()]
    with zipfile.ZipFile(path, "w") as zf:
        for n, d in items:
            if n == "xl/worksheets/sheet2.xml":
                d = (b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
                     b"<row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row></sheetData></worksheet>")
            zf.writestr(n, d)
    _patch(path, "Second", "A2", "second-row")
    root = _sheet(path, "xl/worksheets/sheet2.xml")
    assert [r.get("r") for r in root.xpath("//m:row", namespaces=NS)] == ["1", "2"]
    assert [c.get("r") for c in root.xpath("//m:c", namespaces=NS)] == ["A1", "B1", "A2"]
    assert read_cell(path, "Second", "A2")["cell"]["value"] == "second-row"


# ---- fail-closed structural rules ------------------------------------------------
def test_merged_non_anchor_rejected_anchor_allowed(xlsx_path):
    before = file_revision(xlsx_path)
    with pytest.raises(MergedCellError):
        _patch(xlsx_path, "Data", "B6", "nope")
    assert file_revision(xlsx_path) == before
    _patch(xlsx_path, "Data", "A6", "anchor ok")
    assert read_cell(xlsx_path, "Data", "A6")["cell"]["value"] == "anchor ok"


def test_table_header_rejected_body_allowed(xlsx_path):
    for ref in ("A1", "B1"):
        with pytest.raises(TableHeaderError):
            _patch(xlsx_path, "Data", ref, "renamed")
    _patch(xlsx_path, "Data", "A2", "body ok")


def test_array_formula_region_rejected(xlsx_path):
    for ref in ("D8", "E8", "D9", "E9"):
        with pytest.raises(ArrayFormulaError):
            _patch(xlsx_path, "Data", ref, 1)


def test_shared_formula_master_rejected_child_allowed(xlsx_path):
    with pytest.raises(SharedFormulaError):
        _patch(xlsx_path, "Second", "A2", 5)
    _patch(xlsx_path, "Second", "A3", 5)
    assert read_cell(xlsx_path, "Second", "A3")["cell"]["formula"] is None


def test_out_of_range_and_bad_refs(xlsx_path):
    for ref in ("A0", "XFE1", "A1048577", "$A$1", "A1:B2"):
        with pytest.raises(CellAddressError):
            _patch(xlsx_path, "Data", ref, "x")


def test_no_table_no_merge_variant_allows_everything(tmp_path):
    path = build_minimal_xlsx(tmp_path / "plain.xlsx", merge=False, table=False, array=False)
    _patch(path, "Data", "A1", "free header")
    assert read_cell(path, "Data", "A1")["cell"]["value"] == "free header"


def test_truncate_is_linear_time(xlsx_path):
    start = time.perf_counter()
    _patch(xlsx_path, "Data", "A2", "z" * 500_000, truncate=True)
    assert time.perf_counter() - start < 3.0  # ban cu (cat tung ky tu) mat > 10 giay
    assert len(read_cell(xlsx_path, "Data", "A2")["cell"]["value"]) == 32767


def test_only_xlsx_can_be_patched(tmp_path):
    from core.xlsx_package import UnsupportedWorkbookError

    macro = build_minimal_xlsx(tmp_path / "book.xlsm")
    original = macro.read_bytes()
    with pytest.raises(UnsupportedWorkbookError):
        _patch(macro, "Data", "A2", "x")
    assert macro.read_bytes() == original
    assert read_cell(macro, "Data", "A2")["cell"]["value"] == "Alice"  # doc van duoc


def test_overwrite_drops_cell_metadata_attrs(tmp_path):
    import zipfile

    path = build_minimal_xlsx(tmp_path / "cm.xlsx")
    with zipfile.ZipFile(path) as zf:
        items = [(n, zf.read(n)) for n in zf.namelist()]
    with zipfile.ZipFile(path, "w") as zf:
        for n, d in items:
            if n == "xl/worksheets/sheet1.xml":
                d = d.replace(b'<c r="B3">', b'<c r="B3" cm="1" vm="2">')
            zf.writestr(n, d)
    _patch(path, "Data", "B3", "plain")
    c = _cell(path, "B3")
    assert c.get("cm") is None and c.get("vm") is None and c.get("t") == "inlineStr"
