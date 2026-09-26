import zipfile

import pytest
from conftest import build_minimal_xlsx, file_revision

from core.inspector import SheetNotFoundError, inspect_workbook, read_cell, read_sheet
from core.safety_gateway import UnsafeXmlError, parse_xml_safe
from core.xlsx_package import UnsupportedWorkbookError, XlsxFormatError, XlsxPackage


def _cells(result):
    return {c["ref"]: c for c in result["cells"]}


def test_inspect_structure(xlsx_path):
    report = inspect_workbook(xlsx_path)
    assert report["document_revision"] == file_revision(xlsx_path)
    data, second = report["sheets"]
    assert (data["index"], data["name"], data["dimension"]) == (1, "Data", "A1:E9")
    assert second["name"] == "Second" and second["index"] == 2
    assert data["merged_cells"] == ["A6:B6"]
    assert data["tables"] == [{"name": "Table1", "ref": "A1:B3", "header_rows": 1}]
    assert data["array_formulas"] == ["D8:E9"]
    assert data["formula_count"] == 3  # B4, C4, D8
    assert report["calc"] == {"full_calc_on_load": False, "has_calc_chain": True}
    locs = {loc["locator"]: loc for loc in report["locators"]}
    assert locs["cell_s1_r2_c1"]["ref"] == "A2" and locs["cell_s1_r2_c1"]["sheet"] == "Data"
    assert locs["cell_s2_r2_c1"]["has_formula"] is True
    assert report["locators_truncated"] is False


def test_inspect_locator_limit(xlsx_path):
    report = inspect_workbook(xlsx_path, locator_limit=3)
    assert len(report["locators"]) == 3 and report["locators_truncated"] is True


def test_read_sheet_value_types(xlsx_path):
    cells = _cells(read_sheet(xlsx_path, "Data"))
    assert cells["A1"]["value"] == "Name" and cells["A1"]["type"] == "string"
    assert cells["B1"]["value"] == "Qty"  # rich text runs duoc ghep: "Q" + "ty"
    assert cells["A2"]["value"] == "Alice" and cells["A2"]["style"] == "2"
    assert cells["B2"]["value"] == 5 and cells["B2"]["type"] == "number"
    assert cells["A4"]["value"] == "Total"  # shared string
    assert cells["B6"]["type"] == "blank" and cells["B6"]["value"] is None


def test_formula_cache_and_stale(xlsx_path):
    cells = _cells(read_sheet(xlsx_path, "Data"))
    b4, c4 = cells["B4"], cells["C4"]
    assert b4["has_formula"] and b4["formula"] == "SUM(B2:B3)" and b4["cached_value"] == 12 and b4["stale"] is False
    assert c4["has_formula"] and c4["formula"] == "B4*2"
    assert c4["cached_value"] is None and c4["value"] is None and c4["stale"] is True  # o cong thuc chua co cache


def test_full_calc_on_load_marks_all_formulas_stale(tmp_path):
    path = build_minimal_xlsx(tmp_path / "f.xlsx", full_calc_on_load=True)
    b4 = _cells(read_sheet(path, "Data"))["B4"]
    assert b4["cached_value"] == 12 and b4["stale"] is True


def test_shared_formula_child_flagged(xlsx_path):
    cells = _cells(read_sheet(xlsx_path, "Second"))
    assert cells["A2"]["formula"] == "A1+1" and cells["A2"]["formula_kind"] == "shared"
    assert cells["A3"]["has_formula"] is True and cells["A3"]["formula"] is None and cells["A3"]["formula_kind"] == "shared"


def test_merge_info(xlsx_path):
    cells = _cells(read_sheet(xlsx_path, "Data"))
    assert cells["A6"]["merge"] == {"range": "A6:B6", "anchor": True}
    assert cells["B6"]["merge"] == {"range": "A6:B6", "anchor": False}
    assert cells["A2"]["merge"] is None


def test_read_sheet_range_filter_and_sheet_by_index(xlsx_path):
    result = read_sheet(xlsx_path, "1", "A1:B2")
    assert sorted(_cells(result)) == ["A1", "A2", "B1", "B2"]
    assert result["sheet"] == "Data"
    assert read_sheet(xlsx_path, 2)["sheet"] == "Second"


def test_read_cell_existing_and_missing(xlsx_path):
    assert read_cell(xlsx_path, "Data", "B2")["cell"]["value"] == 5
    missing = read_cell(xlsx_path, "Data", "Z99")["cell"]
    assert missing["type"] == "blank" and missing["locator"] == "cell_s1_r99_c26" and missing["formula"] is None


def test_unknown_sheet(xlsx_path):
    with pytest.raises(SheetNotFoundError):
        read_sheet(xlsx_path, "Nope")
    with pytest.raises(SheetNotFoundError):
        read_sheet(xlsx_path, "9")


def test_inspect_is_read_only(xlsx_path, tmp_path):
    before = sorted(p.name for p in tmp_path.rglob("*"))
    rev = file_revision(xlsx_path)
    inspect_workbook(xlsx_path)
    read_sheet(xlsx_path, "Data")
    read_cell(xlsx_path, "Data", "A1")
    assert sorted(p.name for p in tmp_path.rglob("*")) == before
    assert file_revision(xlsx_path) == rev


# ---- package rejection / XXE -------------------------------------------------
def test_reject_non_zip_and_legacy_and_encrypted(tmp_path):
    ole = tmp_path / "old.xlsx"
    ole.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 64)
    with pytest.raises(UnsupportedWorkbookError, match="OLE2"):
        XlsxPackage.open(ole)
    junk = tmp_path / "junk.xlsx"
    junk.write_bytes(b"not a zip")
    with pytest.raises(UnsupportedWorkbookError, match="magic"):
        XlsxPackage.open(junk)
    for suffix in (".xls", ".xlsb"):
        p = tmp_path / f"legacy{suffix}"
        p.write_bytes(b"PK\x03\x04")
        with pytest.raises(UnsupportedWorkbookError):
            XlsxPackage.open(p)


def test_reject_xlsb_payload_and_missing_workbook(tmp_path):
    xlsb = tmp_path / "b.xlsx"
    with zipfile.ZipFile(xlsb, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("xl/workbook.bin", b"x")
    with pytest.raises(UnsupportedWorkbookError, match="xlsb"):
        XlsxPackage.open(xlsb)
    empty = tmp_path / "e.xlsx"
    with zipfile.ZipFile(empty, "w") as zf:
        zf.writestr("hello.txt", "hi")
    with pytest.raises(XlsxFormatError, match="Thieu part"):
        XlsxPackage.open(empty)


def test_reject_zip_slip(tmp_path):
    bad = tmp_path / "slip.xlsx"
    with zipfile.ZipFile(bad, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("xl/workbook.xml", "<workbook/>")
        zf.writestr("../evil.xml", "x")
    with pytest.raises(XlsxFormatError, match="khong an toan"):
        XlsxPackage.open(bad)


def test_xxe_and_entities_rejected(tmp_path):
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET")
    payload = f'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///{secret.as_posix()}">]><x>&e;</x>'.encode()
    with pytest.raises(UnsafeXmlError):
        parse_xml_safe(payload)
    with pytest.raises(UnsafeXmlError):
        parse_xml_safe(b'<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY a "aaaa">]><lolz>&a;</lolz>')
    assert parse_xml_safe(b"<a><b/></a>").tag == "a"
    with pytest.raises(ValueError):
        parse_xml_safe(b"<a><b></a>")


def test_workbook_with_xxe_sheet_is_rejected_end_to_end(tmp_path):
    path = build_minimal_xlsx(tmp_path / "x.xlsx")
    with zipfile.ZipFile(path) as zf:
        items = [(n, zf.read(n)) for n in zf.namelist()]
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in items:
            if name == "xl/worksheets/sheet1.xml":
                data = b'<?xml version="1.0"?><!DOCTYPE s [<!ENTITY e SYSTEM "file:///etc/passwd">]><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&e;</t></is></c></row></sheetData></worksheet>'
            zf.writestr(name, data)
    with pytest.raises(UnsafeXmlError):
        read_sheet(path, "Data")


# ---- chartsheet / dialogsheet ---------------------------------------------------
def _with_chartsheet(path):
    with zipfile.ZipFile(path) as zf:
        items = [(n, zf.read(n)) for n in zf.namelist()]
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in items:
            if name == "xl/workbook.xml":
                data = data.replace(b'<sheet name="Second"', b'<sheet name="Chart1" sheetId="9" r:id="rId9"/><sheet name="Second"')
            elif name == "xl/_rels/workbook.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>',
                )
            zf.writestr(name, data)
        zf.writestr("xl/chartsheets/sheet1.xml", '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>')
    return path


def test_chartsheet_does_not_break_inspect_read_or_patch(tmp_path):
    from core.inspector import UnsupportedSheetError
    from core.surgical_patcher import PatchRejectedError, patch_cell

    path = _with_chartsheet(build_minimal_xlsx(tmp_path / "chart.xlsx"))
    report = inspect_workbook(path)
    kinds = {s["name"]: s["kind"] for s in report["sheets"]}
    assert kinds == {"Data": "worksheet", "Chart1": "chartsheet", "Second": "worksheet"}
    assert next(s for s in report["sheets"] if s["name"] == "Chart1")["cell_count"] == 0
    assert read_sheet(path, "Data")["sheet"] == "Data"
    with pytest.raises(UnsupportedSheetError):
        read_sheet(path, "Chart1")
    original = path.read_bytes()
    with pytest.raises(PatchRejectedError):
        patch_cell(path, "Chart1", "A1", "x", file_revision(path))
    assert path.read_bytes() == original
    patch_cell(path, "Second", "B1", "ok", file_revision(path))  # sheet thuong (chi so 3) van sua duoc


def test_read_decodes_x_escapes_roundtrip(xlsx_path):
    from core.surgical_patcher import patch_cell

    literal = "x _x0041_ _x005F_ end"
    patch_cell(xlsx_path, "Data", "A2", literal, file_revision(xlsx_path))
    assert read_cell(xlsx_path, "Data", "A2")["cell"]["value"] == literal
    from core.inspector import decode_x_escapes

    assert decode_x_escapes("_x0041_") == "A" and decode_x_escapes("_x005F_x0041_") == "_x0041_"
    assert decode_x_escapes("_xD83D_") == "_xD83D_"  # surrogate rieng le giu nguyen
    assert decode_x_escapes("_x12_") == "_x12_"


def test_doctype_rejected_even_when_utf16_encoded():
    xml = '<?xml version="1.0" encoding="UTF-16"?><!DOCTYPE x [<!ENTITY e "boom">]><x>&e;</x>'.encode("utf-16")
    with pytest.raises(UnsafeXmlError):
        parse_xml_safe(xml)
