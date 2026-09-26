"""
tools/word-engine/core/text_resolver.py - Multi-Run Text Range Resolver cho tools/word-engine/.

Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 3.2 ("Multi-Run Text Range Resolver
(SAFE_TEXT_RUN, COMPLEX_SPAN & Edit Fidelity)"). Ham cong khai resolve_and_replace_text() nhan
mot phan tu w:p DA XAC DINH (Paragraph Boundary: khong bao gio tim/thay the xuyen qua nhieu doan
van, vi caller luon truyen dung 1 w:p - dung theo dung nghia cua tham so, khong can code rieng de
chan xuyen doan) va tim search_text trong toan bo van ban cua doan van do (noi ca text nam trong
w:hyperlink/w:sdt/tracked-changes, de con tinh dung offset ky tu truoc khi quyet dinh co fail-closed
hay khong).

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" (xem
core/__init__.py). WORD_NS duoc tai su dung tu core/inspector.py (nam trong CUNG du an, dung
nguyen tac "tai su dung noi bo dung du an con dang sua" - khong phai import cheo du an).

Quyet dinh thiet ke ve chuan hoa NFC (khong lap lai nguyen van de bai vi de bai chi noi "chuan hoa
truoc khi tim kiem", khong noi ro pham vi ghi lai): moi run duoc chuan hoa NFC MOT LAN o dau ham va
dung chinh chuoi da chuan hoa do lam "van ban lam viec" cho ca buoc tim kiem LAN buoc ghi lai (thay
vi tim tren ban NFC roi ghi lai tren ban tho chua chuan hoa) - vi 2 ban co the lech do dai o mep
ky tu to hop (NFD) neu chi chuan hoa rieng cho tim kiem, lam sai vi tri index luc ghi. Trong truong
hop pho bien (van ban .docx da la NFC san, la dinh dang Word mac dinh xuat ra), chuan hoa nay la
no-op nen phan van ban KHONG bi cham (prefix/suffix) van giu nguyen y het byte goc.

Quyet dinh thiet ke ve "do dai cho phep phan bo tuong ung" (Edit Fidelity, section 3.2): dieu kien
duy nhat de thu Edit Fidelity la len(replace_text) == do dai doan khop goc (matched span) - khi do
moi ky tu thay the co dung 1 vi tri tuyet doi tuong ung voi 1 ky tu bi thay trong span goc, nen co
the "rai" (distribute) replace_text vao dung ranh gioi run cu theo vi tri tuyet doi ma khong can
doan/suy ra cach chia nao khac "dung" hon - day la cach doc "cho phep phan bo tuong ung" it doan
bua nhat, khop voi vi du "L (normal) + oi (bold)" -> "N (normal) + oi (bold)" cua ke hoach (cung do
dai 3 ky tu, danh gioi run khong doi). Neu do dai khac nhau va co heterogeneous rPr, KHONG suy doan
cach phan bo nao khac - bao loi STYLE_BOUNDARY_CONFLICT tru khi co style_policy tuong minh.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from typing import List, Optional

from lxml import etree

from core.inspector import WORD_NS

XML_NS = "http://www.w3.org/XML/1998/namespace"
XML_SPACE_ATTR = f"{{{XML_NS}}}space"


def _w(local_name: str) -> str:
    return f"{{{WORD_NS}}}{local_name}"


RUN_TAG = _w("r")
TEXT_TAG = _w("t")
RPR_TAG = _w("rPr")

# Cac the container "bao boc" run - neu matched span giao thoa/nam trong bat ky the nao trong so
# nay (theo to tien cua bat ky run nao thuoc span), ABORT COMPLEX_SPAN ngay (section 3.2). Ghi chu
# ke hoach: w:drawing/w:object co the vua la container bao boc run vua la node con dac biet ben
# trong run - liet ke o day de bat ca truong hop bao boc; truong hop la node con da duoc
# _is_safe_text_run() bat rieng.
CONTAINER_TAGS = frozenset(
    {
        "hyperlink",
        "sdt",
        "sdtContent",
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "fldSimple",
        "object",
        "drawing",
    }
)


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------
class TextNotFoundError(RuntimeError):
    def __init__(self, search_text: str):
        super().__init__(f"Khong tim thay search_text trong pham vi doan van: {search_text!r}")
        self.search_text = search_text


class ComplexSpanError(RuntimeError):
    """ABORT: COMPLEX_SPAN - fail-closed tuyet doi, tu choi tu dong rewrite (section 3.2)."""

    def __init__(
        self,
        message: str = "COMPLEX_SPAN: Matched span touches complex inline elements or containers",
    ):
        super().__init__(message)


class StyleBoundaryConflictError(RuntimeError):
    def __init__(
        self,
        message: str = (
            "STYLE_BOUNDARY_CONFLICT: Replacement crosses heterogeneous style boundaries "
            "without explicit policy"
        ),
    ):
        super().__init__(message)


# ---------------------------------------------------------------------------
# Result model
# ---------------------------------------------------------------------------
@dataclass
class TextReplaceResult:
    search_text: str
    replace_text: str
    # "single_run" | "same_style_multi_run" | "edit_fidelity" | "collapsed_prefer_start" |
    # "collapsed_prefer_end"
    run_mode: str
    start_run_index: int
    end_run_index: int
    runs_removed: int
    style_policy_applied: Optional[str] = None


# ---------------------------------------------------------------------------
# Run helpers
# ---------------------------------------------------------------------------
def _run_text(run: etree._Element) -> str:
    raw = "".join(t.text or "" for t in run.findall(TEXT_TAG))
    return unicodedata.normalize("NFC", raw)


def _is_safe_text_run(run: etree._Element) -> bool:
    """SAFE_TEXT_RUN (section 3.2): run chi duoc phep chua w:rPr va w:t - bat ky the con nao
    khac (w:fldChar, w:instrText, w:drawing, w:object, w:br, w:tab, footnote/endnote/comment
    reference...) deu lam run nay KHONG an toan, khong can liet ke tung the rieng le."""
    for child in run:
        if not isinstance(child.tag, str):
            continue  # bo qua comment/processing-instruction node cua lxml
        if etree.QName(child).localname not in ("rPr", "t"):
            return False
    return True


def _ancestor_container_tag(run: etree._Element, stop_at: etree._Element) -> Optional[str]:
    """Di tu cha cua run len den truoc stop_at (w:p), tra ve ten the container dau tien gap phai
    neu co (xem CONTAINER_TAGS), nguoc lai None."""
    node = run.getparent()
    while node is not None and node is not stop_at:
        if isinstance(node.tag, str):
            local_name = etree.QName(node).localname
            if local_name in CONTAINER_TAGS:
                return local_name
        node = node.getparent()
    return None


def _rpr_key(run: etree._Element) -> bytes:
    """Khoa so sanh style: canonical XML (C14N) cua w:rPr, hoac chuoi rong neu khong co rPr / rPr
    rong (khong con nao) - 2 run deu "khong co style rieng" duoc coi la cung style."""
    rpr = run.find(RPR_TAG)
    if rpr is None or len(rpr) == 0:
        return b""
    return etree.tostring(rpr, method="c14n")


def _set_text_preserve_space(t_elem: etree._Element, text: str) -> None:
    t_elem.text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        t_elem.set(XML_SPACE_ATTR, "preserve")
    elif t_elem.get(XML_SPACE_ATTR) is not None:
        del t_elem.attrib[XML_SPACE_ATTR]


def _set_run_text(run: etree._Element, text: str) -> None:
    """Ghi lai noi dung van ban cua 1 SAFE_TEXT_RUN bang dung 1 w:t (xoa moi w:t cu - truong hop
    pho bien la dung 1 w:t/run, gop truong hop hiem co nhieu w:t/run thanh 1). w:rPr (neu co) luon
    dung nguyen vi tri truoc do, w:t moi duoc them vao sau cung - dung thu tu schema (rPr truoc
    t)."""
    for t in run.findall(TEXT_TAG):
        run.remove(t)
    t_elem = etree.SubElement(run, TEXT_TAG)
    _set_text_preserve_space(t_elem, text)


def _remove_run(run: etree._Element) -> None:
    parent = run.getparent()
    if parent is not None:
        parent.remove(run)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def resolve_and_replace_text(
    p_elem: etree._Element,
    search_text: str,
    replace_text: str,
    style_policy: Optional[str] = None,
) -> TextReplaceResult:
    """Tim search_text trong pham vi CHINH XAC 1 doan van p_elem va thay bang replace_text theo
    thuat toan Minimal Run Rewrite (section 3.2). Nem ComplexSpanError neu span khop cham/nam
    trong container phuc tap hoac cham run chua phan tu con dac biet (fail-closed tuyet doi -
    khong tu dong rewrite). Nem StyleBoundaryConflictError neu span vat qua cac run co w:rPr khac
    nhau ma khong the phan bo tuong duong va khong co style_policy tuong minh. Nem
    TextNotFoundError neu khong tim thay search_text."""
    if style_policy not in (None, "prefer-start", "prefer-end"):
        raise ValueError(f"style_policy khong hop le: {style_policy!r}")
    if not search_text:
        raise ValueError("search_text khong duoc rong")

    search_norm = unicodedata.normalize("NFC", search_text)
    replace_norm = unicodedata.normalize("NFC", replace_text)

    runs = list(p_elem.iter(RUN_TAG))
    run_texts = [_run_text(r) for r in runs]

    offsets: List[tuple[int, int]] = []
    pos = 0
    for text in run_texts:
        offsets.append((pos, pos + len(text)))
        pos += len(text)
    full_text = "".join(run_texts)

    idx = full_text.find(search_norm)
    if idx == -1:
        raise TextNotFoundError(search_text)
    end_idx = idx + len(search_norm)

    overlapping = [i for i, (s, e) in enumerate(offsets) if s < end_idx and e > idx]
    if not overlapping:
        raise TextNotFoundError(search_text)

    # --- COMPLEX_SPAN fail-closed (section 3.2) -----------------------------
    for i in overlapping:
        run = runs[i]
        if _ancestor_container_tag(run, p_elem) is not None:
            raise ComplexSpanError()
        if not _is_safe_text_run(run):
            raise ComplexSpanError()

    start_i, end_i = overlapping[0], overlapping[-1]
    start_run, end_run = runs[start_i], runs[end_i]

    # --- Single-run match ----------------------------------------------------
    if start_i == end_i:
        run_text = run_texts[start_i]
        run_start = offsets[start_i][0]
        local_start = idx - run_start
        local_end = end_idx - run_start
        new_text = run_text[:local_start] + replace_norm + run_text[local_end:]
        _set_run_text(start_run, new_text)
        return TextReplaceResult(
            search_text=search_text,
            replace_text=replace_text,
            run_mode="single_run",
            start_run_index=start_i,
            end_run_index=end_i,
            runs_removed=0,
        )

    # --- Multi-run match -------------------------------------------------
    rpr_keys = {_rpr_key(runs[i]) for i in overlapping}
    same_style = len(rpr_keys) == 1

    prefix = run_texts[start_i][: idx - offsets[start_i][0]]
    suffix = run_texts[end_i][end_idx - offsets[end_i][0] :]
    matched_len = end_idx - idx

    if same_style:
        _set_run_text(start_run, prefix + replace_norm)
        removed = 0
        for i in overlapping[1:-1]:
            _remove_run(runs[i])
            removed += 1
        if suffix:
            _set_run_text(end_run, suffix)
        else:
            _remove_run(end_run)
            removed += 1
        return TextReplaceResult(
            search_text=search_text,
            replace_text=replace_text,
            run_mode="same_style_multi_run",
            start_run_index=start_i,
            end_run_index=end_i,
            runs_removed=removed,
        )

    # Heterogeneous rPr across the span - try Edit Fidelity first (same length only, see module
    # docstring), else fall back to an explicit style_policy, else fail-closed.
    if len(replace_norm) == matched_len:
        for i in overlapping:
            run_text = run_texts[i]
            run_start, run_end = offsets[i]
            overlap_start = max(idx, run_start)
            overlap_end = min(end_idx, run_end)
            local_prefix = run_text[: overlap_start - run_start]
            local_suffix = run_text[overlap_end - run_start :]
            replace_slice = replace_norm[overlap_start - idx : overlap_end - idx]
            _set_run_text(runs[i], local_prefix + replace_slice + local_suffix)
        return TextReplaceResult(
            search_text=search_text,
            replace_text=replace_text,
            run_mode="edit_fidelity",
            start_run_index=start_i,
            end_run_index=end_i,
            runs_removed=0,
        )

    if style_policy is None:
        raise StyleBoundaryConflictError()

    style_source = start_run if style_policy == "prefer-start" else end_run
    _set_run_text(style_source, prefix + replace_norm + suffix)
    removed = 0
    for i in overlapping:
        if runs[i] is style_source:
            continue
        _remove_run(runs[i])
        removed += 1
    return TextReplaceResult(
        search_text=search_text,
        replace_text=replace_text,
        run_mode=f"collapsed_{style_policy.replace('-', '_')}",
        start_run_index=start_i,
        end_run_index=end_i,
        runs_removed=removed,
        style_policy_applied=style_policy,
    )
