"""Pytest configuration for tools/excel-engine/tests/.

- Them tools/excel-engine/ (cha cua `core/`) vao sys.path.
- Fixture autouse dat cwd va work_dir ve `tmp_path`: moi test chay hoan toan trong thu muc tam,
  khong tao file rac tren o dia (`.jarvis/`, backup, lock deu nam trong tmp_path).
- `build_minimal_xlsx()` dung zipfile truc tiep dung 1 workbook nho co calcChain, inlineStr,
  sharedStrings, mergedCells, Table, array formula, shared formula, style, part phu (anh, customXml).
"""
import hashlib
import sys
import zipfile
from pathlib import Path
from typing import Dict, List

import pytest

sys.dont_write_bytecode = True

WORKSPACE_ROOT = Path(__file__).resolve().parent.parent
if str(WORKSPACE_ROOT) not in sys.path:
    sys.path.insert(0, str(WORKSPACE_ROOT))

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


@pytest.fixture(autouse=True)
def _isolate_cwd(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)


@pytest.fixture
def work_dir(tmp_path) -> Path:
    return tmp_path


def _sheet1_xml(merge: bool, table: bool, array: bool) -> str:
    rows = [
        '<row r="1" spans="1:2"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" t="s" s="1"><v>1</v></c></row>',
        '<row r="2" spans="1:2"><c r="A2" t="inlineStr" s="2"><is><t>Alice</t></is></c><c r="B2" s="3"><v>5</v></c></row>',
        '<row r="3" spans="1:2"><c r="A3" t="inlineStr"><is><t>Bob</t></is></c><c r="B3"><v>7</v></c></row>',
        '<row r="4"><c r="A4" t="s"><v>2</v></c><c r="B4" s="3"><f>SUM(B2:B3)</f><v>12</v></c><c r="C4"><f>B4*2</f></c></row>',
    ]
    if merge:
        rows.append('<row r="6"><c r="A6" t="inlineStr" s="2"><is><t>Merged</t></is></c><c r="B6" s="2"/></row>')
    if array:
        rows.append(
            '<row r="8"><c r="D8"><f t="array" ref="D8:E9">B2:B3*2</f><v>10</v></c><c r="E8"><v>14</v></c></row>'
            '<row r="9"><c r="D9"><v>10</v></c><c r="E9"><v>14</v></c></row>'
        )
    parts = [DECL, f'<worksheet xmlns="{NS}" xmlns:r="{NS_R}">', '<dimension ref="A1:E9"/>',
             '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
             '<cols><col min="1" max="2" width="14" customWidth="1"/></cols>',
             "<sheetData>", *rows, "</sheetData>"]
    if merge:
        parts.append('<mergeCells count="1"><mergeCell ref="A6:B6"/></mergeCells>')
    if table:
        parts.append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
    parts.append("</worksheet>")
    return "".join(parts)


def build_minimal_xlsx(
    path: Path,
    *,
    calc_chain: bool = True,
    merge: bool = True,
    table: bool = True,
    array: bool = True,
    calc_pr: bool = True,
    full_calc_on_load: bool = False,
) -> Path:
    """Dung workbook 2 sheet: "Data" (Table, merge, array, formula, sparse rows) va "Second"
    (shared formula). Tra ve `path`."""
    path = Path(path)
    ct_over = [
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
        ("/xl/worksheets/sheet2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
        ("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
    ]
    if calc_chain:
        ct_over.append(("/xl/calcChain.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"))
    if table:
        ct_over.append(("/xl/tables/table1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"))
    content_types = (
        DECL + f'<Types xmlns="{NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
        + "".join(f'<Override PartName="{n}" ContentType="{t}"/>' for n, t in ct_over) + "</Types>"
    )
    wb_rels = [
        ("rId1", "worksheet", "worksheets/sheet1.xml"),
        ("rId2", "worksheet", "worksheets/sheet2.xml"),
        ("rId3", "styles", "styles.xml"),
        ("rId4", "sharedStrings", "sharedStrings.xml"),
    ]
    if calc_chain:
        wb_rels.append(("rId5", "calcChain", "calcChain.xml"))
    wb_rels_xml = DECL + f'<Relationships xmlns="{NS_REL}">' + "".join(
        f'<Relationship Id="{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{t}" Target="{tg}"/>'
        for i, t, tg in wb_rels
    ) + "</Relationships>"
    calc = ""
    if calc_pr:
        calc = '<calcPr calcId="191029"' + (' fullCalcOnLoad="1"' if full_calc_on_load else "") + "/>"
    workbook = (
        DECL + f'<workbook xmlns="{NS}" xmlns:r="{NS_R}"><workbookPr/>'
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets>'
        '<definedNames><definedName name="Total">Data!$B$4</definedName></definedNames>'
        + calc + "</workbook>"
    )
    sheet2 = (
        DECL + f'<worksheet xmlns="{NS}"><dimension ref="A1:A3"/><sheetData>'
        '<row r="1"><c r="A1"><v>1</v></c></row>'
        '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">A1+1</f><v>2</v></c></row>'
        '<row r="3"><c r="A3"><f t="shared" si="0"/><v>3</v></c></row>'
        "</sheetData></worksheet>"
    )
    entries: List[tuple] = [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", DECL + f'<Relationships xmlns="{NS_REL}"><Relationship Id="rId1" '
         'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
        ("docProps/core.xml", DECL + '<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>'),
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", wb_rels_xml),
        ("xl/styles.xml", DECL + f'<styleSheet xmlns="{NS}"><fonts count="1"><font><sz val="11"/></font></fonts>'
         '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>'
         '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="4"><xf/><xf/><xf/><xf/></cellXfs></styleSheet>'),
        ("xl/sharedStrings.xml", DECL + f'<sst xmlns="{NS}" count="3" uniqueCount="3">'
         '<si><t>Name</t></si><si><r><t>Q</t></r><r><t>ty</t></r></si><si><t>Total</t></si></sst>'),
        ("xl/worksheets/sheet1.xml", _sheet1_xml(merge, table, array)),
        ("xl/worksheets/sheet2.xml", sheet2),
        ("customXml/item1.xml", DECL + "<root>keep-me</root>"),
        ("xl/media/image1.png", b"\x89PNG\r\n\x1a\n" + bytes(range(256)) * 4),
    ]
    if table:
        entries.insert(9, ("xl/worksheets/_rels/sheet1.xml.rels", DECL + f'<Relationships xmlns="{NS_REL}"><Relationship Id="rId1" '
                           'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>'))
        entries.append(("xl/tables/table1.xml", DECL + f'<table xmlns="{NS}" id="1" name="Table1" displayName="Table1" ref="A1:B3">'
                        '<tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/></tableColumns></table>'))
    if calc_chain:
        entries.append(("xl/calcChain.xml", DECL + f'<calcChain xmlns="{NS}"><c r="B4" i="1"/><c r="C4" i="1"/><c r="A2" i="2"/></calcChain>'))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries:
            zf.writestr(name, content)
    return path


def part_hashes(path: Path) -> Dict[str, str]:
    with zipfile.ZipFile(path) as zf:
        return {n: hashlib.sha256(zf.read(n)).hexdigest() for n in zf.namelist()}


def part_names(path: Path) -> List[str]:
    with zipfile.ZipFile(path) as zf:
        return zf.namelist()


def read_part(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path) as zf:
        return zf.read(name)


def file_revision(path: Path) -> str:
    return "sha256:" + hashlib.sha256(Path(path).read_bytes()).hexdigest()


@pytest.fixture
def xlsx_path(tmp_path) -> Path:
    return build_minimal_xlsx(tmp_path / "book.xlsx")
