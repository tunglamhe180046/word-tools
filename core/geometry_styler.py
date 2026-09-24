"""
tools/word-engine/core/geometry_styler.py - Geometry & Layout Standardization (Phase 2) cho
tools/word-engine/.

Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md muc 4 ("Phase 2: Dan trang & Hinh hoc") va
nhiem vu Phase 2 duoc giao: tach bach TUYET DOI 3 nhom cau hinh OOXML doc lap voi nhau:
  1. Pagination (chong vo trang): w:cantSplit tren moi w:tr, w:tblHeader tren hang dau, w:keepNext
     tren doan van ngay truoc bang.
  2. Borders (vien bang): w:tblBorders (w:tblPr) + w:tcBorders (tung w:tcPr) voi 4 preset
     all|none|horizontal-only|notary-standard.
  3. Page Setup (kho giay & le): w:pgSz + w:pgMar trong w:sectPr cuoi cung cua tai lieu.

Module nay la lop dieu phoi (orchestration) GHEP LAI cac primitive da co san cua chinh
tools/word-engine/ - khong tu viet lai Commit Broker/locator state machine:
  - core/inspector.py: state machine locator (resolve_in_document, load_document_root,
    inspect_document) khi table_locator duoc truyen.
  - core/safety_gateway.py: sha256_file (drift check truoc khi sua).
  - core/surgical_patcher.py: PatchResult (tai su dung nguyen dataclass, replace_mode=None cho
    geometry vi khong lien quan Multi-Run Text Resolver), _repack_document_xml() va
    _stage_and_commit() (dung LAI nguyen ven co che dong goi lai goi OPC + Commit Broker 18 buoc
    da kiem toan, khong viet lai - dung tinh than nac 2 cua CLAUDE.md goc: tai su dung noi bo dung
    du an con dang sua truoc khi viet moi).

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" (xem
core/__init__.py va docs/plans muc 1.2).

Ghi chu ve ten the OOXML "keep with next" (quan trong, doc truoc khi sua module nay): de bai
nhiem vu Phase 2 goi ten than thien "w:keepWithNext" cho hanh vi "doan van tieu de/doan van ngay
truoc bang luon di lien voi bang phia sau". The OOXML THUC TE thuc hien dung hanh vi nay la
w:keepNext (CT_PPrBase, "Keep Paragraph With Next Paragraph") - KHONG ton tai the w:keepWithNext
nao trong dac ta OOXML (ECMA-376/ISO 29500). Module nay dung w:keepNext de tai lieu sinh ra hop
le va co hieu luc that trong Word, thay vi ghi mot the khong ton tai chi vi khop dung ten goi
trong de bai (fabricate mot XML tag gia se lam Word bo qua no lang le, phan tac dung voi chinh
muc dich "chong vo trang" cua nhiem vu).

3 nhom hoan toan doc lap ve mat cai dat: moi ham public trong module nay CHI dung/xoa dung cac
the XML cua DUNG 1 nhom (set_pagination() chi dung cantSplit/tblHeader/keepNext; set_borders() chi
dung tblBorders/tcBorders; set_page_size()/set_page_margins() chi dung pgSz/pgMar) - khong ham nao
doc lai hay xoa element cua nhom khac, nen goi tuan tu ca 3 trong apply_geometry() khong bao gio
lam mat cau hinh cua nhau. Thu tu chen element con trong tung parent (w:tblPr/w:trPr/w:tcPr/
w:pPr/w:sectPr) tuan theo dung sequence cua schema CT_* tuong ung (xem cac hang so *_ORDER duoi
day) qua ham dung chung _set_ordered_child() - khong chen bua vi Word co the tu choi mo file neu
thu tu element trong 1 CT_* sai (validation nghiem ngat hon nhieu so voi HTML/JSON).
"""

from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
from typing import Dict, List, Optional, Union

from lxml import etree

from core.inspector import (
    WORD_NS,
    DocumentDriftError,
    LocatorNotFoundError,
    ObjectLocator,
    inspect_document,
    load_document_root,
    resolve_in_document,
)
from core.safety_gateway import sha256_file
from core.surgical_patcher import PatchResult, _repack_document_xml, _stage_and_commit  # noqa: F401 (repack tai su dung qua _stage_and_commit)

_NSMAP = {"w": WORD_NS}


def _w(tag: str) -> str:
    return f"{{{WORD_NS}}}{tag}"


_BODY_TAG = _w("body")
_TBL_TAG = _w("tbl")
_TR_TAG = _w("tr")
_TC_TAG = _w("tc")
_P_TAG = _w("p")

# ---------------------------------------------------------------------------
# Schema child-sequence order (CT_* tuong ung, ECMA-376 Part 1) - dung boi _set_ordered_child()
# de chen/thay element con dung vi tri, khong lam file .docx sinh ra bi Word tu choi mo.
# ---------------------------------------------------------------------------
TBLPR_ORDER = [
    _w(t) for t in [
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
        "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
    ]
]
TRPR_ORDER = [
    _w(t) for t in [
        "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
        "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange",
    ]
]
TCPR_ORDER = [
    _w(t) for t in [
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
        "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel",
        "cellMerge", "tcPrChange",
    ]
]
PPR_ORDER = [
    _w(t) for t in [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
        "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId",
        "cnfStyle", "rPr", "sectPr", "pPrChange",
    ]
]
SECTPR_ORDER = [
    _w(t) for t in [
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
        "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
        "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
        "sectPrChange",
    ]
]

_BORDER_EDGES = ["top", "left", "bottom", "right", "insideH", "insideV"]


# ---------------------------------------------------------------------------
# Generic XML element-ordering helpers (dung chung cho ca 3 nhom)
# ---------------------------------------------------------------------------
def _get_or_insert_first(parent: etree._Element, tag: str) -> etree._Element:
    """Lay element con <tag> da co, hoac tao moi va chen lam con DAU TIEN - w:tblPr/w:trPr/
    w:tcPr/w:pPr bat buoc phai la con dau tien cua parent theo schema OOXML."""
    existing = parent.find(tag)
    if existing is not None:
        return existing
    new_el = etree.Element(tag)
    parent.insert(0, new_el)
    return new_el


def _set_ordered_child(parent: etree._Element, tag: str, order: List[str], new_element: etree._Element) -> None:
    """Xoa element con <tag> cu (neu co) roi chen new_element vao DUNG vi tri theo `order` (danh
    sach Clark-notation qname theo dung sequence cua CT_* tuong ung) - khong bao gio dung
    append() bua, vi Word tu choi mo file .docx co thu tu element con sai schema."""
    existing = parent.find(tag)
    if existing is not None:
        parent.remove(existing)
    try:
        tag_rank = order.index(tag)
    except ValueError:
        parent.append(new_element)
        return
    insert_at = len(parent)
    for i, child in enumerate(parent):
        try:
            child_rank = order.index(child.tag)
        except ValueError:
            continue
        if child_rank > tag_rank:
            insert_at = i
            break
    parent.insert(insert_at, new_element)


# ---------------------------------------------------------------------------
# Nhom 1: Pagination (chong vo trang)
# ---------------------------------------------------------------------------
def set_pagination(tbl: etree._Element) -> None:
    """Gan w:cantSplit cho MOI w:tr cua bang (cam ngat hang qua 2 trang), w:tblHeader cho hang
    dau tien (lap lai tieu de khi bang dai sang trang sau), va w:keepNext cho doan van (w:p) nam
    NGAY TRUOC bang trong w:body - neu co - de doan van do luon di lien voi bang (xem module
    docstring ve ten the that su la w:keepNext, khong phai "w:keepWithNext"). Khong dung cham den
    bat ky the w:tblBorders/w:tcBorders/w:pgSz/w:pgMar nao."""
    for i, tr in enumerate(tbl.findall(_TR_TAG)):
        tr_pr = _get_or_insert_first(tr, _w("trPr"))
        _set_ordered_child(tr_pr, _w("cantSplit"), TRPR_ORDER, etree.Element(_w("cantSplit")))
        if i == 0:
            _set_ordered_child(tr_pr, _w("tblHeader"), TRPR_ORDER, etree.Element(_w("tblHeader")))

    prev_sibling = tbl.getprevious()
    if prev_sibling is not None and prev_sibling.tag == _P_TAG:
        p_pr = _get_or_insert_first(prev_sibling, _w("pPr"))
        _set_ordered_child(p_pr, _w("keepNext"), PPR_ORDER, etree.Element(_w("keepNext")))


# ---------------------------------------------------------------------------
# Nhom 2: Borders (vien bang)
# ---------------------------------------------------------------------------
def _edge_attrs(val: str, sz: Optional[str] = None, color: Optional[str] = None) -> Dict[str, str]:
    attrs = {"val": val}
    if sz is not None:
        attrs["sz"] = sz
        attrs["space"] = "0"
    if color is not None:
        attrs["color"] = color
    return attrs


def _border_spec(preset: str) -> Dict[str, Dict[str, str]]:
    """Tra ve dict {top|left|bottom|right|insideH|insideV: attrs} cho dung 1 trong 4 preset -
    dung chung cho ca w:tblBorders (cap bang) va w:tcBorders (cap tung o, xem module docstring va
    de bai nhiem vu: "Thiet lap w:tblBorders... va w:tcBorders (tung o)"). val="nil" = tuong minh
    khong ke (uu tien hon bat ky vien ke thua tu table style), khong phai bo qua the."""
    single_normal = _edge_attrs("single", sz="4", color="000000")
    single_thin = _edge_attrs("single", sz="2", color="000000")
    nil = _edge_attrs("nil")
    if preset == "all":
        return {edge: single_normal for edge in _BORDER_EDGES}
    if preset == "none":
        return {edge: nil for edge in _BORDER_EDGES}
    if preset == "horizontal-only":
        return {
            "top": single_normal, "bottom": single_normal, "insideH": single_normal,
            "left": nil, "right": nil, "insideV": nil,
        }
    if preset == "notary-standard":
        # "vien ngoai don, duong ngang mong, khong ke doc" (de bai nhiem vu Phase 2 muc 2).
        return {
            "top": single_normal, "left": single_normal, "bottom": single_normal, "right": single_normal,
            "insideH": single_thin, "insideV": nil,
        }
    raise ValueError(f"borders_preset khong hop le: {preset!r} (chi nhan all|none|horizontal-only|notary-standard)")


def _build_border_element(tag: str, spec: Dict[str, Dict[str, str]]) -> etree._Element:
    new_el = etree.Element(tag)
    for edge in _BORDER_EDGES:
        edge_el = etree.SubElement(new_el, _w(edge))
        for attr_name, attr_val in spec[edge].items():
            edge_el.set(_w(attr_name), attr_val)
    return new_el


def _iter_direct_cells(tbl: etree._Element):
    """Chi cac w:tc la con TRUC TIEP cua tbl (qua tung w:tr) - khong dung vao o cua bang long ben
    trong (nested table), giu dung pham vi 1 bang duy nhat dang duoc styling."""
    for tr in tbl.findall(_TR_TAG):
        for tc in tr.findall(_TC_TAG):
            yield tc


def set_borders(tbl: etree._Element, preset: str) -> None:
    """Thiet lap w:tblBorders (trong w:tblPr) va w:tcBorders cho TUNG o truc tiep cua bang (trong
    w:tcPr cua moi o) theo dung 1 trong 4 preset. Khong dung cham den bat ky the w:cantSplit/
    w:tblHeader/w:keepNext/w:pgSz/w:pgMar nao (xem module docstring)."""
    spec = _border_spec(preset)

    tbl_pr = _get_or_insert_first(tbl, _w("tblPr"))
    _set_ordered_child(tbl_pr, _w("tblBorders"), TBLPR_ORDER, _build_border_element(_w("tblBorders"), spec))

    for tc in _iter_direct_cells(tbl):
        tc_pr = _get_or_insert_first(tc, _w("tcPr"))
        _set_ordered_child(tc_pr, _w("tcBorders"), TCPR_ORDER, _build_border_element(_w("tcBorders"), spec))


# ---------------------------------------------------------------------------
# Nhom 3: Page Setup (kho giay & le)
# ---------------------------------------------------------------------------
_PAGE_SIZES = {
    "A4": {"w": "11906", "h": "16838"},
    "A5": {"w": "8391", "h": "11906"},
}
_MARGIN_PRESETS = {
    "notary": {"left": "1701", "right": "850", "top": "1134", "bottom": "1134"},
    "standard": {"left": "1440", "right": "1440", "top": "1440", "bottom": "1440"},
    "compact": {"left": "720", "right": "720", "top": "720", "bottom": "720"},
}


def _get_or_create_final_sect_pr(body: etree._Element) -> etree._Element:
    """sectPr CUOI CUNG cua tai lieu la con truc tiep cuoi cung cua w:body (CT_Body:
    (EG_BlockLevelElts)* sectPr?) - Phase 1/2 MVP chi xu ly dung section nay, cung pham vi voi
    core/inspector.py (chi doc con truc tiep cua w:body). Tao moi bang append() neu chua co (dung
    vi tri cuoi cung vi sectPr phai la con cuoi cung khi ton tai)."""
    sect_pr = body.find(_w("sectPr"))
    if sect_pr is None:
        sect_pr = etree.SubElement(body, _w("sectPr"))
    return sect_pr


def set_page_size(sect_pr: etree._Element, page_size: str) -> None:
    if page_size not in _PAGE_SIZES:
        raise ValueError(f"page_size khong hop le: {page_size!r} (chi nhan A4|A5)")
    dims = _PAGE_SIZES[page_size]
    existing = sect_pr.find(_w("pgSz"))
    new_el = etree.Element(_w("pgSz"))
    if existing is not None:
        orient = existing.get(_w("orient"))
        if orient is not None:
            new_el.set(_w("orient"), orient)
    new_el.set(_w("w"), dims["w"])
    new_el.set(_w("h"), dims["h"])
    _set_ordered_child(sect_pr, _w("pgSz"), SECTPR_ORDER, new_el)


def set_page_margins(sect_pr: etree._Element, margins: str) -> None:
    if margins not in _MARGIN_PRESETS:
        raise ValueError(f"margins khong hop le: {margins!r} (chi nhan notary|standard|compact)")
    values = _MARGIN_PRESETS[margins]
    existing = sect_pr.find(_w("pgMar"))
    new_el = etree.Element(_w("pgMar"))
    # Giu nguyen header/footer/gutter cu neu w:pgMar da ton tai (chi doi top/right/bottom/left) -
    # mac dinh 720/720/0 (twips) neu chua co w:pgMar nao, dung quy uoc mac dinh cua Word.
    header = existing.get(_w("header")) if existing is not None else "720"
    footer = existing.get(_w("footer")) if existing is not None else "720"
    gutter = existing.get(_w("gutter")) if existing is not None else "0"
    new_el.set(_w("top"), values["top"])
    new_el.set(_w("right"), values["right"])
    new_el.set(_w("bottom"), values["bottom"])
    new_el.set(_w("left"), values["left"])
    if header is not None:
        new_el.set(_w("header"), header)
    if footer is not None:
        new_el.set(_w("footer"), footer)
    if gutter is not None:
        new_el.set(_w("gutter"), gutter)
    _set_ordered_child(sect_pr, _w("pgMar"), SECTPR_ORDER, new_el)


# ---------------------------------------------------------------------------
# Resolve table_locator -> dung 1 w:tbl (state machine, cung nguyen tac voi surgical_patcher.py)
# ---------------------------------------------------------------------------
def _resolve_target_table(
    docx_path: Path, root: etree._Element, table_locator: Optional[ObjectLocator]
) -> Optional[etree._Element]:
    """table_locator la 1 ObjectLocator kind == "table_cell" (core/inspector.py Phase 1 MVP chua
    co kind "table" rieng - chi "paragraph" | "table_cell"), dung de xac dinh DUNG bang can ap
    dung pagination/borders khi tai lieu co nhieu hon 1 bang: resolve dung state machine cua
    inspector.py (kiem tra DOCUMENT_DRIFT truoc, roi resolve_in_document() voi fast structural
    lookup + fallback context_sha256) roi leo tu w:tc tim duoc len w:tr/w:tbl chua no - KHONG dung
    truc tiep table_locator.table_index (co the da dich chuyen ke tu luc quet, day chinh la ly do
    resolve_in_document() ton tai). Tra ve None neu table_locator la None (nghia la: ap dung cho
    TOAN BO bang trong tai lieu - xem apply_geometry() docstring)."""
    if table_locator is None:
        return None
    if table_locator.kind != "table_cell":
        raise ValueError(f"table_locator.kind phai la 'table_cell', nhan duoc {table_locator.kind!r}")

    current_revision = f"sha256:{sha256_file(docx_path)}"
    if current_revision != table_locator.revision:
        raise DocumentDriftError(table_locator.revision, current_revision)

    tc_element = resolve_in_document(root, asdict(table_locator))
    tbl = tc_element
    while tbl is not None and tbl.tag != _TBL_TAG:
        tbl = tbl.getparent()
    if tbl is None:
        raise LocatorNotFoundError(table_locator.object_id)
    return tbl


_VALID_ALIGNMENTS = {"left", "center", "right", "both"}


def set_paragraph_alignment(p: etree._Element, alignment: str) -> None:
    """Thiet lap w:jc (can le) cho doan van (w:p).
    Gia tri hop le: 'left', 'center', 'right', 'both'.
    Chen vao dung vi tri schema PPR_ORDER."""
    if alignment not in _VALID_ALIGNMENTS:
        raise ValueError(f"alignment khong hop le: {alignment!r} (chi nhan left|center|right|both)")
    p_pr = _get_or_insert_first(p, _w("pPr"))
    jc_el = etree.Element(_w("jc"))
    jc_el.set(_w("val"), alignment)
    _set_ordered_child(p_pr, _w("jc"), PPR_ORDER, jc_el)


def _resolve_target_element(
    docx_path: Path, root: etree._Element, locator: Optional[ObjectLocator]
) -> Optional[etree._Element]:
    if locator is None:
        return None
    current_revision = f"sha256:{sha256_file(docx_path)}"
    if current_revision != locator.revision:
        raise DocumentDriftError(locator.revision, current_revision)
    return resolve_in_document(root, asdict(locator))


def balance_table_grid(tbl: etree._Element, target_width: int = 10368) -> bool:
    """Can chinh lai luoi cot va chieu rong cac o cua bang neu phat hien bang bi lech
    (lech chieu rong giua cac hang, hoac gridCol khong khop voi cac hang).
    Dac biet xu ly bang tong ket hoc ba (6 cot noi dung va 1 hang cuoi span toan bo).
    Tra ve True neu da can chinh, False neu khong khop mau bang can xu ly."""
    trs = tbl.findall(_TR_TAG)
    if len(trs) < 2:
        return False

    last_tr = trs[-1]
    last_tcs = last_tr.findall(_TC_TAG)
    first_tr_tcs = trs[0].findall(_TC_TAG)

    if len(last_tcs) == 1 and len(first_tr_tcs) in (4, 6):
        cols = [1548, 1702, 1702, 2012, 1702, 1702]
        total_w = sum(cols)  # 10368

        # 1. Update tblGrid
        grid = tbl.find(_w("tblGrid"))
        if grid is not None:
            grid.clear()
            for w in cols:
                etree.SubElement(grid, _w("gridCol"), {_w("w"): str(w)})

        # 2. Update tblW & jc
        tbl_pr = _get_or_insert_first(tbl, _w("tblPr"))
        tbl_w_el = _get_or_insert_first(tbl_pr, _w("tblW"))
        tbl_w_el.set(_w("w"), str(total_w))
        tbl_w_el.set(_w("type"), "dxa")

        jc_el = etree.Element(_w("jc"))
        jc_el.set(_w("val"), "center")
        _set_ordered_child(tbl_pr, _w("jc"), TBLPR_ORDER, jc_el)

        # 3. Row 0
        if len(first_tr_tcs) == 4:
            first_tr_tcs[0].find(_w("tcPr")).find(_w("tcW")).set(_w("w"), str(cols[0]))
            first_tr_tcs[1].find(_w("tcPr")).find(_w("tcW")).set(_w("w"), str(cols[1] + cols[2]))
            first_tr_tcs[2].find(_w("tcPr")).find(_w("tcW")).set(_w("w"), str(cols[3]))
            first_tr_tcs[3].find(_w("tcPr")).find(_w("tcW")).set(_w("w"), str(cols[4] + cols[5]))
        elif len(first_tr_tcs) == 6:
            for i, tc in enumerate(first_tr_tcs):
                tc.find(_w("tcPr")).find(_w("tcW")).set(_w("w"), str(cols[i]))

        # 4. Middle rows
        for r in trs[1:-1]:
            r_tcs = r.findall(_TC_TAG)
            if len(r_tcs) == 6:
                for i, tc in enumerate(r_tcs):
                    tc_pr = tc.find(_w("tcPr"))
                    if tc_pr is not None:
                        tc_w = tc_pr.find(_w("tcW"))
                        if tc_w is not None:
                            tc_w.set(_w("w"), str(cols[i]))

        # 5. Last row
        last_tc_pr = last_tcs[0].find(_w("tcPr"))
        if last_tc_pr is not None:
            last_tc_w = last_tc_pr.find(_w("tcW"))
            if last_tc_w is not None:
                last_tc_w.set(_w("w"), str(total_w))
            grid_span = last_tc_pr.find(_w("gridSpan"))
            if grid_span is not None:
                grid_span.set(_w("val"), "6")

        return True
    return False


def set_table_width_dxa(tbl: etree._Element, target_width: int) -> None:
    """Thu/phong DEU TY LE (proportional scale) toan bo w:gridCol va w:tcW cua 1 bang ve dung
    target_width (dxa), giu nguyen ty le tuong doi giua cac cot - khac balance_table_grid() o tren
    (chi nhan dien duoc DUNG 1 mau bang tong ket hoc ba co dinh 4/6 cot, tra ve False voi bat ky
    hinh dang bang nao khac). Ham nay ap dung duoc cho BAT KY so cot nao (vd cac bang khung the
    CCCD front/back/notary chi co 1-3 cot don gian, khong khop mau hoc ba) bang cach scale ty le
    thay vi ap mot bo do rong cot cung. Dat w:tblW type=dxa + w:jc=center. Bo qua (no-op) neu bang
    khong co w:tblGrid hoac tong do rong hien tai <= 0 (khong the tinh ty le)."""
    grid = tbl.find(_w("tblGrid"))
    if grid is None:
        return
    cols = grid.findall(_w("gridCol"))
    if not cols:
        return
    current_widths = [int(c.get(_w("w"), "0")) for c in cols]
    current_total = sum(current_widths)
    if current_total <= 0:
        return

    scale = target_width / current_total
    new_widths = [round(w * scale) for w in current_widths]
    new_widths[-1] += target_width - sum(new_widths)  # don rounding drift ve cot cuoi

    for col_el, new_w in zip(cols, new_widths):
        col_el.set(_w("w"), str(new_w))

    tbl_pr = _get_or_insert_first(tbl, _w("tblPr"))
    tbl_w_el = etree.Element(_w("tblW"))
    tbl_w_el.set(_w("w"), str(target_width))
    tbl_w_el.set(_w("type"), "dxa")
    _set_ordered_child(tbl_pr, _w("tblW"), TBLPR_ORDER, tbl_w_el)

    jc_el = etree.Element(_w("jc"))
    jc_el.set(_w("val"), "center")
    _set_ordered_child(tbl_pr, _w("jc"), TBLPR_ORDER, jc_el)

    for tr in tbl.findall(_TR_TAG):
        col_i = 0
        for tc in tr.findall(_TC_TAG):
            tc_pr = tc.find(_w("tcPr"))
            if tc_pr is None:
                col_i += 1
                continue
            grid_span_el = tc_pr.find(_w("gridSpan"))
            span = int(grid_span_el.get(_w("val"))) if grid_span_el is not None else 1
            span_width = sum(new_widths[col_i:col_i + span]) if col_i < len(new_widths) else 0
            tc_w = tc_pr.find(_w("tcW"))
            if tc_w is not None and span_width:
                tc_w.set(_w("w"), str(span_width))
            col_i += span


def strip_leading_empty_paragraphs(body: etree._Element) -> int:
    """Xoa cac w:p RONG (khong co bat ky w:t nao co noi dung, khong co w:drawing/w:pict/w:object)
    khi dung LAM CON DAU TIEN cua w:body (chi tinh cac doan van truoc bang dau tien - KHONG dung
    den bat ky doan van nao o giua hoac sau bang, vi do co the la khoang cach co chu dich), va khi
    dung LAM CON DAU TIEN cua bat ky w:tc nao co NHIEU HON 1 doan van truc tiep (giu lai neu do la
    doan van DUY NHAT cua o - 1 o luon can it nhat 1 w:p theo schema OOXML). Day la san pham thua
    dien hinh cua python-docx add_table()/add_paragraph() (xem CLAUDE.md goc, muc nhiem vu). Tra ve
    tong so doan van da xoa."""

    def _is_empty_paragraph(p: etree._Element) -> bool:
        if p.tag != _P_TAG:
            return False
        has_text = any((t.text or "").strip() for t in p.iter(_w("t")))
        has_object = next(p.iter(_w("drawing")), None) is not None or next(p.iter(_w("pict")), None) is not None
        return not has_text and not has_object

    removed = 0

    for child in list(body):
        if child.tag == _TBL_TAG:
            break
        if _is_empty_paragraph(child):
            body.remove(child)
            removed += 1
        else:
            break

    for tbl in body.findall(_TBL_TAG):
        for tc in tbl.iter(_TC_TAG):
            paragraphs = tc.findall(_P_TAG)
            while len(paragraphs) > 1 and _is_empty_paragraph(paragraphs[0]):
                tc.remove(paragraphs[0])
                removed += 1
                paragraphs = tc.findall(_P_TAG)

    return removed


def expand_transcript_table(tbl: etree._Element, row_height_dxa: int = 520) -> None:
    """Dat chieu cao toi thieu (w:trHeight atLeast) cho cac hang va can giua doc (w:vAlign center)
    cho cac o cua bang diem de lap day trang A4, tranh khoang trang trong o cuoi trang."""
    for tr in tbl.findall(_TR_TAG):
        tr_pr = _get_or_insert_first(tr, _w("trPr"))
        tr_height = etree.Element(_w("trHeight"), {_w("val"): str(row_height_dxa), _w("hRule"): "atLeast"})
        _set_ordered_child(tr_pr, _w("trHeight"), TRPR_ORDER, tr_height)
        for tc in tr.findall(_TC_TAG):
            tc_pr = _get_or_insert_first(tc, _w("tcPr"))
            v_align = etree.Element(_w("vAlign"), {_w("val"): "center"})
            _set_ordered_child(tc_pr, _w("vAlign"), TCPR_ORDER, v_align)


def box_visual_placeholders(body: etree._Element) -> int:
    """Enclose [ National Emblem ], [ Photo of Holder ], and fingerprint placeholders
    inside dedicated bordered mini-table boxes according to box_and_photo_rules.md.
    Returns number of cells modified.
    """
    def _make_tbl_border(color="000000", sz="4", val="single"):
        tbl_borders = etree.Element(_w("tblBorders"))
        for side in ("top", "left", "bottom", "right"):
            el = etree.SubElement(tbl_borders, _w(side))
            el.set(_w("val"), val)
            el.set(_w("sz"), sz)
            el.set(_w("space"), "0")
            el.set(_w("color"), color)
        for side in ("insideH", "insideV"):
            el = etree.SubElement(tbl_borders, _w(side))
            el.set(_w("val"), "none")
        return tbl_borders

    def _make_tc_border(color="000000", sz="4", val="single"):
        tc_borders = etree.Element(_w("tcBorders"))
        for side in ("top", "left", "bottom", "right"):
            el = etree.SubElement(tc_borders, _w(side))
            el.set(_w("val"), val)
            el.set(_w("sz"), sz)
            el.set(_w("space"), "0")
            el.set(_w("color"), color)
        return tc_borders

    def _set_tc_mar(tc_pr, top=30, bottom=30, left=20, right=20):
        mar = etree.SubElement(tc_pr, _w("tcMar"))
        for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
            el = etree.SubElement(mar, _w(side))
            el.set(_w("w"), str(val))
            el.set(_w("type"), "dxa")

    def _create_p(text, font_sz="14", italic=True, bold=False, before="40", after="40", align="center", subtext=None, subtext_sz="12"):
        p = etree.Element(_w("p"))
        p_pr = etree.SubElement(p, _w("pPr"))
        jc = etree.SubElement(p_pr, _w("jc"))
        jc.set(_w("val"), align)
        sp = etree.SubElement(p_pr, _w("spacing"))
        sp.set(_w("before"), before)
        sp.set(_w("after"), after)
        sp.set(_w("line"), "240")
        sp.set(_w("lineRule"), "auto")

        if text:
            r = etree.SubElement(p, _w("r"))
            r_pr = etree.SubElement(r, _w("rPr"))
            r_fonts = etree.SubElement(r_pr, _w("rFonts"))
            r_fonts.set(_w("ascii"), "Times New Roman")
            r_fonts.set(_w("hAnsi"), "Times New Roman")
            sz_el = etree.SubElement(r_pr, _w("sz"))
            sz_el.set(_w("val"), font_sz)
            if italic:
                etree.SubElement(r_pr, _w("i"))
            if bold:
                etree.SubElement(r_pr, _w("b"))
            t_el = etree.SubElement(r, _w("t"))
            t_el.text = text

        if subtext:
            r2 = etree.SubElement(p, _w("r"))
            etree.SubElement(r2, _w("br"))
            r2_pr = etree.SubElement(r2, _w("rPr"))
            r2_fonts = etree.SubElement(r2_pr, _w("rFonts"))
            r2_fonts.set(_w("ascii"), "Times New Roman")
            r2_fonts.set(_w("hAnsi"), "Times New Roman")
            sz2_el = etree.SubElement(r2_pr, _w("sz"))
            sz2_el.set(_w("val"), subtext_sz)
            if italic:
                etree.SubElement(r2_pr, _w("i"))
            if bold:
                etree.SubElement(r2_pr, _w("b"))
            t2_el = etree.SubElement(r2, _w("t"))
            t2_el.text = subtext
        return p

    modified_count = 0

    for tbl in body.findall(_TBL_TAG):
        for tc in tbl.iter(_TC_TAG):
            if tc.find(_TBL_TAG) is not None:
                continue

            tc_text = "".join((t.text or "") for t in tc.iter(_w("t")))

            # 1. Front card: National Emblem + Photo
            if "[ National Emblem ]" in tc_text or "[ Photo of Holder ]" in tc_text:
                expiry_text = ""
                for p in tc.findall(_P_TAG):
                    all_t = "".join((t.text or "") for t in p.iter(_w("t")))
                    if "expiry" in all_t.lower():
                        lines = [line.strip() for line in all_t.split("\n") if line.strip()]
                        if len(lines) > 1:
                            expiry_text = lines[-1]
                        elif ":" in all_t:
                            expiry_text = all_t.split(":")[-1].strip()

                tc_pr = tc.find(_w("tcPr"))
                for ch in list(tc):
                    if ch != tc_pr:
                        tc.remove(ch)

                # Emblem Box
                emblem_tbl = etree.SubElement(tc, _w("tbl"))
                e_tbl_pr = etree.SubElement(emblem_tbl, _w("tblPr"))
                e_tbl_w = etree.SubElement(e_tbl_pr, _w("tblW"))
                e_tbl_w.set(_w("w"), "1500")
                e_tbl_w.set(_w("type"), "dxa")
                e_jc = etree.SubElement(e_tbl_pr, _w("jc"))
                e_jc.set(_w("val"), "center")
                e_tbl_pr.append(_make_tbl_border(sz="4"))
                e_grid = etree.SubElement(emblem_tbl, _w("tblGrid"))
                e_col = etree.SubElement(e_grid, _w("gridCol"))
                e_col.set(_w("w"), "1500")
                e_tr = etree.SubElement(emblem_tbl, _w("tr"))
                e_tr_pr = etree.SubElement(e_tr, _w("trPr"))
                etree.SubElement(e_tr_pr, _w("cantSplit"))
                e_h = etree.SubElement(e_tr_pr, _w("trHeight"))
                e_h.set(_w("val"), "600")
                e_h.set(_w("hRule"), "atLeast")
                e_tc = etree.SubElement(e_tr, _w("tc"))
                e_tc_pr = etree.SubElement(e_tc, _w("tcPr"))
                e_tc_w = etree.SubElement(e_tc_pr, _w("tcW"))
                e_tc_w.set(_w("w"), "1500")
                e_tc_w.set(_w("type"), "dxa")
                _set_tc_mar(e_tc_pr, top=30, bottom=30, left=20, right=20)
                e_valign = etree.SubElement(e_tc_pr, _w("vAlign"))
                e_valign.set(_w("val"), "center")
                e_tc.append(_create_p("[ National Emblem ]", font_sz="14", italic=True, before="40", after="40"))

                # Spacer p
                tc.append(_create_p("", before="20", after="20"))

                # Photo Box
                photo_tbl = etree.SubElement(tc, _w("tbl"))
                p_tbl_pr = etree.SubElement(photo_tbl, _w("tblPr"))
                p_tbl_w = etree.SubElement(p_tbl_pr, _w("tblW"))
                p_tbl_w.set(_w("w"), "1872")
                p_tbl_w.set(_w("type"), "dxa")
                p_jc = etree.SubElement(p_tbl_pr, _w("jc"))
                p_jc.set(_w("val"), "center")
                p_tbl_pr.append(_make_tbl_border(sz="4"))
                p_grid = etree.SubElement(photo_tbl, _w("tblGrid"))
                p_col = etree.SubElement(p_grid, _w("gridCol"))
                p_col.set(_w("w"), "1872")
                p_tr = etree.SubElement(photo_tbl, _w("tr"))
                p_tr_pr = etree.SubElement(p_tr, _w("trPr"))
                etree.SubElement(p_tr_pr, _w("cantSplit"))
                p_h = etree.SubElement(p_tr_pr, _w("trHeight"))
                p_h.set(_w("val"), "2300")
                p_h.set(_w("hRule"), "atLeast")
                p_tc = etree.SubElement(p_tr, _w("tc"))
                p_tc_pr = etree.SubElement(p_tc, _w("tcPr"))
                p_tc_w = etree.SubElement(p_tc_pr, _w("tcW"))
                p_tc_w.set(_w("w"), "1872")
                p_tc_w.set(_w("type"), "dxa")
                _set_tc_mar(p_tc_pr, top=40, bottom=40, left=30, right=30)
                p_valign = etree.SubElement(p_tc_pr, _w("vAlign"))
                p_valign.set(_w("val"), "center")
                p_tc.append(_create_p("[ Photo of Holder ]", font_sz="15", italic=True, before="120", after="120"))

                # Expiry p (guarantees cell ends with w:p)
                p_exp = etree.SubElement(tc, _w("p"))
                p_exp_pr = etree.SubElement(p_exp, _w("pPr"))
                jc_exp = etree.SubElement(p_exp_pr, _w("jc"))
                jc_exp.set(_w("val"), "center")
                sp_exp = etree.SubElement(p_exp_pr, _w("spacing"))
                sp_exp.set(_w("before"), "40")
                sp_exp.set(_w("after"), "0")
                r_exp1 = etree.SubElement(p_exp, _w("r"))
                r_exp1_pr = etree.SubElement(r_exp1, _w("rPr"))
                rf1 = etree.SubElement(r_exp1_pr, _w("rFonts"))
                rf1.set(_w("ascii"), "Times New Roman")
                rf1.set(_w("hAnsi"), "Times New Roman")
                etree.SubElement(r_exp1_pr, _w("b"))
                sz1 = etree.SubElement(r_exp1_pr, _w("sz"))
                sz1.set(_w("val"), "15")
                t_exp1 = etree.SubElement(r_exp1, _w("t"))
                t_exp1.text = "Date of expiry:"

                if expiry_text:
                    r_exp2 = etree.SubElement(p_exp, _w("r"))
                    etree.SubElement(r_exp2, _w("br"))
                    r_exp2_pr = etree.SubElement(r_exp2, _w("rPr"))
                    rf2 = etree.SubElement(r_exp2_pr, _w("rFonts"))
                    rf2.set(_w("ascii"), "Times New Roman")
                    rf2.set(_w("hAnsi"), "Times New Roman")
                    etree.SubElement(r_exp2_pr, _w("b"))
                    sz2 = etree.SubElement(r_exp2_pr, _w("sz"))
                    sz2.set(_w("val"), "15")
                    t_exp2 = etree.SubElement(r_exp2, _w("t"))
                    t_exp2.text = expiry_text

                modified_count += 1

            # 2. Back card: Left & Right Fingerprints
            elif "fingerprint" in tc_text.lower() or ("left index finger" in tc_text.lower() and "right index finger" in tc_text.lower()):
                tc_pr = tc.find(_w("tcPr"))
                for ch in list(tc):
                    if ch != tc_pr:
                        tc.remove(ch)

                fp_tbl = etree.SubElement(tc, _w("tbl"))
                fp_tbl_pr = etree.SubElement(fp_tbl, _w("tblPr"))
                fp_tbl_w = etree.SubElement(fp_tbl_pr, _w("tblW"))
                fp_tbl_w.set(_w("w"), "2200")
                fp_tbl_w.set(_w("type"), "dxa")
                fp_jc = etree.SubElement(fp_tbl_pr, _w("jc"))
                fp_jc.set(_w("val"), "center")
                fp_sp = etree.SubElement(fp_tbl_pr, _w("tblCellSpacing"))
                fp_sp.set(_w("w"), "40")
                fp_sp.set(_w("type"), "dxa")

                fp_grid = etree.SubElement(fp_tbl, _w("tblGrid"))
                fp_c1 = etree.SubElement(fp_grid, _w("gridCol"))
                fp_c1.set(_w("w"), "1080")
                fp_c2 = etree.SubElement(fp_grid, _w("gridCol"))
                fp_c2.set(_w("w"), "1080")

                fp_tr = etree.SubElement(fp_tbl, _w("tr"))
                fp_tr_pr = etree.SubElement(fp_tr, _w("trPr"))
                etree.SubElement(fp_tr_pr, _w("cantSplit"))
                fp_h = etree.SubElement(fp_tr_pr, _w("trHeight"))
                fp_h.set(_w("val"), "1300")
                fp_h.set(_w("hRule"), "atLeast")

                # Left FP Cell
                l_tc = etree.SubElement(fp_tr, _w("tc"))
                l_tc_pr = etree.SubElement(l_tc, _w("tcPr"))
                l_tc_w = etree.SubElement(l_tc_pr, _w("tcW"))
                l_tc_w.set(_w("w"), "1080")
                l_tc_w.set(_w("type"), "dxa")
                l_tc_pr.append(_make_tc_border(sz="4"))
                _set_tc_mar(l_tc_pr, top=20, bottom=20, left=10, right=10)
                l_valign = etree.SubElement(l_tc_pr, _w("vAlign"))
                l_valign.set(_w("val"), "center")
                l_tc.append(_create_p("[ Left index finger ]", font_sz="12", italic=True, before="30", after="20", subtext="(Fingerprinted)", subtext_sz="12"))

                # Right FP Cell
                r_tc = etree.SubElement(fp_tr, _w("tc"))
                r_tc_pr = etree.SubElement(r_tc, _w("tcPr"))
                r_tc_w = etree.SubElement(r_tc_pr, _w("tcW"))
                r_tc_w.set(_w("w"), "1080")
                r_tc_w.set(_w("type"), "dxa")
                r_tc_pr.append(_make_tc_border(sz="4"))
                _set_tc_mar(r_tc_pr, top=20, bottom=20, left=10, right=10)
                r_valign = etree.SubElement(r_tc_pr, _w("vAlign"))
                r_valign.set(_w("val"), "center")
                r_tc.append(_create_p("[ Right index finger ]", font_sz="12", italic=True, before="30", after="20", subtext="(Fingerprinted)", subtext_sz="12"))

                # Trailing empty p in parent tc
                tc.append(_create_p("", before="0", after="0"))
                modified_count += 1

    return modified_count


# ---------------------------------------------------------------------------
# Public API: apply_geometry()
# ---------------------------------------------------------------------------
def apply_geometry(
    docx_path: Union[str, Path],
    page_size: Optional[str] = None,
    margins: Optional[str] = None,
    pagination: bool = False,
    borders_preset: Optional[str] = None,
    table_locator: Optional[ObjectLocator] = None,
    target_locator: Optional[ObjectLocator] = None,
    align: Optional[str] = None,
    balance: bool = False,
    expand_transcripts: bool = False,
    row_height_dxa: Optional[int] = None,
    table_width_dxa: Optional[int] = None,
    strip_leading_empty: bool = False,
    box_placeholders: bool = False,
    work_dir: Optional[Union[str, Path]] = None,
    allowed_roots: Optional[List[Union[str, Path]]] = None,
    actor: str = "word-engine",
    job_id: Optional[str] = None,
) -> PatchResult:
    """Diem vao cong khai duy nhat cua module (xem docstring nhiem vu Phase 2): ap dung tuan tu,
    hoan toan doc lap cac nhom Page Setup -> Pagination -> Borders -> Alignment -> Balance ->
    Table Width -> Strip Leading Empty -> Expand -> Box Placeholders, commit qua dung Commit Broker
    18 buoc da kiem toan va tra ve locators moi qua mot lan re-inspect.
    """
    docx_path = Path(docx_path)
    if (
        page_size is None
        and margins is None
        and not pagination
        and borders_preset is None
        and align is None
        and not balance
        and not expand_transcripts
        and table_width_dxa is None
        and not strip_leading_empty
        and not box_placeholders
    ):
        raise ValueError(
            "apply_geometry() can it nhat 1 trong page_size/margins/pagination/borders_preset/"
            "align/balance/expand_transcripts/table_width_dxa/strip_leading_empty/box_placeholders."
        )

    root = load_document_root(docx_path)
    body = root.find(_BODY_TAG)
    if body is None:
        raise ValueError(f"{docx_path}: word/document.xml khong co w:body.")

    if page_size is not None or margins is not None:
        sect_pr = _get_or_create_final_sect_pr(body)
        if page_size is not None:
            set_page_size(sect_pr, page_size)
        if margins is not None:
            set_page_margins(sect_pr, margins)

    if pagination or borders_preset is not None:
        target_table = _resolve_target_table(docx_path, root, table_locator)
        tables = [target_table] if target_table is not None else body.findall(_TBL_TAG)
        if not tables:
            raise ValueError(
                f"{docx_path}: khong co bang (w:tbl) nao truc tiep trong w:body de ap dung pagination/borders."
            )
        for tbl in tables:
            if pagination:
                set_pagination(tbl)
            if borders_preset is not None:
                set_borders(tbl, borders_preset)

    if align is not None:
        loc = target_locator or table_locator
        if loc is None:
            raise ValueError("set-geometry voi --align can --target-id de xac dinh doan van hoac bang can can le.")
        target_el = _resolve_target_element(docx_path, root, loc)
        if target_el is None:
            raise LocatorNotFoundError(loc.object_id)
        if target_el.tag == _P_TAG:
            set_paragraph_alignment(target_el, align)
        elif target_el.tag == _TBL_TAG or loc.kind == "table_cell":
            tbl = target_el
            while tbl is not None and tbl.tag != _TBL_TAG:
                tbl = tbl.getparent()
            if tbl is not None:
                tbl_pr = _get_or_insert_first(tbl, _w("tblPr"))
                jc_el = etree.Element(_w("jc"))
                jc_el.set(_w("val"), align)
                _set_ordered_child(tbl_pr, _w("jc"), TBLPR_ORDER, jc_el)
        else:
            raise ValueError(f"Khong the set alignment cho element: {target_el.tag}")

    if balance:
        loc = target_locator or table_locator
        if loc is not None:
            target_el = _resolve_target_element(docx_path, root, loc)
            if target_el is not None:
                tbl = target_el
                while tbl is not None and tbl.tag != _TBL_TAG:
                    tbl = tbl.getparent()
                if tbl is not None:
                    balance_table_grid(tbl)
        else:
            for tbl in body.findall(_TBL_TAG):
                balance_table_grid(tbl)

    if table_width_dxa is not None:
        loc = target_locator or table_locator
        if loc is not None:
            target_el = _resolve_target_element(docx_path, root, loc)
            if target_el is not None:
                tbl = target_el
                while tbl is not None and tbl.tag != _TBL_TAG:
                    tbl = tbl.getparent()
                if tbl is not None:
                    set_table_width_dxa(tbl, table_width_dxa)
        else:
            for tbl in body.findall(_TBL_TAG):
                set_table_width_dxa(tbl, table_width_dxa)

    if strip_leading_empty:
        strip_leading_empty_paragraphs(body)

    if expand_transcripts:
        loc = target_locator or table_locator
        h = row_height_dxa or 520
        if loc is not None:
            target_el = _resolve_target_element(docx_path, root, loc)
            if target_el is not None:
                tbl = target_el
                while tbl is not None and tbl.tag != _TBL_TAG:
                    tbl = tbl.getparent()
                if tbl is not None:
                    expand_transcript_table(tbl, h)
        else:
            for tbl in body.findall(_TBL_TAG):
                if len(tbl.findall(_TR_TAG)) >= 15:
                    expand_transcript_table(tbl, h)

    if box_placeholders:
        box_visual_placeholders(body)

    commit_result = _stage_and_commit(docx_path, root, work_dir, allowed_roots, actor, job_id)
    inspect_report = inspect_document(docx_path, work_dir=work_dir, job_id=job_id)
    return PatchResult(
        success=True,
        outcome=commit_result["outcome"],
        job_id=commit_result["job_id"],
        new_document_revision=inspect_report["document_revision"],
        new_locators=inspect_report["locators"],
        backup_path=commit_result.get("backup_path"),
        replace_mode=None,
    )
