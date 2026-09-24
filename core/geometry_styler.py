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
    work_dir: Optional[Union[str, Path]] = None,
    allowed_roots: Optional[List[Union[str, Path]]] = None,
    actor: str = "word-engine",
    job_id: Optional[str] = None,
) -> PatchResult:
    """Diem vao cong khai duy nhat cua module (xem docstring nhiem vu Phase 2): ap dung tuan tu,
    hoan toan doc lap 3 nhom Page Setup -> Pagination -> Borders -> Alignment -> Balance, commit qua dung Commit
    Broker 18 buoc da kiem toan (_stage_and_commit() tai su dung tu core/surgical_patcher.py, tu
    no da bao gom issue_job() chup baseline hash truoc luc dong goi candidate) va tra ve locators
    moi qua mot lan re-inspect.
    """
    docx_path = Path(docx_path)
    if (
        page_size is None
        and margins is None
        and not pagination
        and borders_preset is None
        and align is None
        and not balance
    ):
        raise ValueError(
            "apply_geometry() can it nhat 1 trong page_size/margins/pagination/borders_preset/align/balance."
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
