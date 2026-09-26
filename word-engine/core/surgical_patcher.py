"""
tools/word-engine/core/surgical_patcher.py - OOXML Surgical Patcher (Profile 1) cho tools/word-engine/.

Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md muc 3.1 (state machine locator), muc 3.3
("Ho so 1: OOXML SURGICAL MODE") va muc 5.1 (bo tieu chuan nghiem thu Profile 1). Module nay la
lop dieu phoi (orchestration) GHEP LAI cac primitive da co san cua chinh tools/word-engine/ -
KHONG tu viet lai bat ky logic nao trong so do:
  - core/inspector.py: state machine locator (DocumentDriftError, resolve_in_document,
    load_document_root, inspect_document) va sha256_file/canonical_json_dumps dung lai tu
    safety_gateway.
  - core/text_resolver.py: Minimal Run Rewrite trong pham vi 1 doan van (SAFE_TEXT_RUN,
    COMPLEX_SPAN fail-closed, STYLE_BOUNDARY_CONFLICT).
  - core/safety_gateway.py: Commit Broker 18 buoc da duoc kiem toan (issue_job + request_commit) -
    dam nhan backup tu dong, optimistic lock, atomic replace, postcondition verification.

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" (xem
core/__init__.py va docs/plans muc 1.2).

Luong thuc thi (dung thu tu Profile 1, xem plan muc 3.3 va nhiem vu):
  1. Kiem tra state machine: sha256 dia hien tai cua docx_path phai == locator.revision, khong
     thi ABORT DocumentDriftError (_resolve_target()) - hoan toan khong tu "rebase" locator.
  2. resolve_locator (fast structural lookup, fallback context_sha256) qua resolve_in_document()
     cua inspector.py (da lam dung 2 nhanh cua state machine trong plan muc 3.1).
  3. Kiem tra noi dung node khop locator.expected_text (buoc phong thu - resolve_in_document() da
     tu bao dam dieu nay o ca 2 nhanh, xem module docstring cua no; kiem tra lai o day de lam
     tuong minh dung buoc 4 cua de bai va bat bat ky sai lech logic nao trong tuong lai).
  4. Phau thuat sua doi CHI phan tu muc tieu, hoan toan trong bo nho (khong dung ra dia) - patch_cell
     chi cham vao 1 w:p ben trong w:tc muc tieu (w:tcPr va cac w:tc khac khong bi dung toi);
     patch_text chi cham vao dung 1 w:p da xac dinh (qua locator hoac tim duy nhat trong toan tai
     lieu) - ca 2 deu uy quyen viec ghi lai run cho resolve_and_replace_text() cua text_resolver.py
     de thua huong nguyen ven SAFE_TEXT_RUN/COMPLEX_SPAN/STYLE_BOUNDARY_CONFLICT.
  5. Dong goi candidate: doc lai TOAN BO cac phan (part) khac cua goi OPC truc tiep tu docx_path
     tren dia (van la ban GOC, vi buoc 4 chi sua trong bo nho) va chi thay the rieng
     word/document.xml bang ban da sua - dam bao "Untouched OPC Parts... RAW SHA-256 IDENTICAL
     100%" cua Profile 1 mot cach tu nhien (khong can co gang "khong dung vao" thu gi, vi don gian
     la khong doc/ghi lai chung).
  6. Commit candidate qua dung Commit Broker da kiem toan: issue_job() (chup baseline hash ngay
     TRUOC luc dong goi candidate - thu hep further khe ho TOCTOU) roi request_commit() (18 buoc:
     backup tu dong o Buoc 10, optimistic lock lai o Buoc 8, atomic replace, postcondition
     verification).
  7. Tu dong re-inspect docx_path SAU khi commit thanh cong (inspect_document() voi CHINH job_id
     do caller truyen vao) va tra ve PatchResult voi new_document_revision/new_locators - day la
     co che "cap locator moi cho caller" duy nhat cua Profile 1 (plan muc 3.1: "Tu dong re-inspect
     -> Tra locator moi cho Caller"), khong phai rebase locator cu.

Neu bat ky buoc 1-4 ném loi (DocumentDriftError, ObjectContentMismatchError, TextNotFoundError,
ComplexSpanError, StyleBoundaryConflictError, AmbiguousTextMatchError...), ham nem thang loi do ra
ngoai va KHONG cham gi den dia - dung quy uoc exception-based cua toan bo core/*.py con lai trong
du an nay (khac voi mot PatchResult(success=False) rieng, de nhat quan voi DocumentDriftError/
TextNotFoundError/... da co).

Gioi han Phase 1 MVP (co chu dich, dung tinh than "khong bia dat du kien" - fail-closed thay vi
doan cach xu ly "hop ly nhat"):
  - patch_cell() chi ho tro o (w:tc) co DUNG 1 doan van (w:p) truc tiep - o co nhieu doan van nem
    CellSpansMultipleParagraphsError thay vi tu doan nen gop/tach the nao.
  - patch_text() khi KHONG truyen locator se tim search_text tren TOAN BO doan van cua tai lieu
    (ca trong than bai va trong o bang) va bat buoc dung 1 doan van khop - qua 1 doan van nem
    AmbiguousTextMatchError (khong tu doan chon dung cai nao, giong tinh than AmbiguousLocatorError
    cua inspector.py).
"""

from __future__ import annotations

import unicodedata
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from lxml import etree

from core.inspector import (
    WORD_NS,
    DocumentDriftError,
    ObjectLocator,
    inspect_document,
    load_document_root,
    resolve_in_document,
)
from core.safety_gateway import issue_job, request_commit, sha256_file
from core.text_resolver import TextNotFoundError, TextReplaceResult, resolve_and_replace_text

_NSMAP = {"w": WORD_NS}
_T_TAG = f"{{{WORD_NS}}}t"
_P_TAG = f"{{{WORD_NS}}}p"


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------
class ObjectContentMismatchError(RuntimeError):
    """Buoc phong thu (xem module docstring, buoc 3): noi dung node vua resolve duoc khac
    locator.expected_text. Ve nguyen tac khong nen bao gio xay ra khi document_revision da khop
    (resolve_in_document() da tu bao dam dieu nay), giu lai de bat som bat ky sai lech logic nao
    trong tuong lai thay vi am tham ghi de sai noi dung."""

    def __init__(self, object_id: str, expected_text: str, actual_text: str):
        super().__init__(
            f"OBJECT_CONTENT_MISMATCH: doi tuong {object_id} co noi dung khac locator.expected_text "
            f"(mong doi {expected_text!r}, thuc te {actual_text!r})."
        )
        self.object_id = object_id
        self.expected_text = expected_text
        self.actual_text = actual_text


class AmbiguousTextMatchError(RuntimeError):
    """ABORT: patch_text() khong truyen locator va tim thay nhieu hon 1 doan van chua search_text -
    tu choi doan xem doan nao la dung (xem module docstring)."""

    def __init__(self, search_text: str, match_count: int):
        super().__init__(
            f"AMBIGUOUS_TEXT_MATCH: tim thay {match_count} doan van chua {search_text!r} trong pham vi "
            "tim kiem - can truyen locator de thu hep pham vi truoc khi patch (khong tu doan doan nao)."
        )
        self.search_text = search_text
        self.match_count = match_count


class CellSpansMultipleParagraphsError(RuntimeError):
    """ABORT: patch_cell() Phase 1 MVP chi ho tro o co dung 1 doan van (xem module docstring)."""

    def __init__(self, object_id: str, paragraph_count: int):
        super().__init__(
            f"CELL_SPANS_MULTIPLE_PARAGRAPHS: o {object_id} co {paragraph_count} doan van - patch_cell() "
            "Phase 1 MVP chi ho tro o co dung 1 doan van truc tiep."
        )
        self.object_id = object_id
        self.paragraph_count = paragraph_count


# ---------------------------------------------------------------------------
# Result model
# ---------------------------------------------------------------------------
@dataclass
class PatchResult:
    success: bool
    outcome: str
    job_id: str
    new_document_revision: str
    new_locators: List[Dict[str, Any]]
    backup_path: Optional[str] = None
    replace_mode: Optional[str] = None


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------
def _element_text(element: etree._Element) -> str:
    return unicodedata.normalize("NFC", "".join(t.text or "" for t in element.iter(_T_TAG)))


def _resolve_target(docx_path: Path, locator: ObjectLocator) -> Tuple[etree._Element, etree._Element]:
    """Buoc 1-3 cua module docstring: state machine drift check + resolve + kiem tra noi dung."""
    current_revision = f"sha256:{sha256_file(docx_path)}"
    if current_revision != locator.revision:
        raise DocumentDriftError(locator.revision, current_revision)
    root = load_document_root(docx_path)
    element = resolve_in_document(root, asdict(locator))
    actual_text = _element_text(element)
    if actual_text != locator.expected_text:
        raise ObjectContentMismatchError(locator.object_id, locator.expected_text, actual_text)
    return root, element


def _repack_document_xml(docx_path: Path, new_root: etree._Element, candidate_path: Path) -> None:
    """Dong goi lai toan bo goi OPC cua docx_path vao candidate_path: moi part khac ngoai
    word/document.xml duoc doc THANG tu docx_path tren dia (van la ban goc, chua bi dung toi) va
    ghi lai y het - day chinh la co che dam bao "Untouched OPC Parts RAW SHA-256 IDENTICAL 100%"
    cua Profile 1 (xem module docstring buoc 5), khong can xu ly gi them."""
    new_xml_bytes = etree.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    candidate_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "r") as src, zipfile.ZipFile(candidate_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = new_xml_bytes if item.filename == "word/document.xml" else src.read(item.filename)
            dst.writestr(item, data)


def _stage_and_commit(
    docx_path: Path,
    root: etree._Element,
    work_dir: Union[str, Path],
    allowed_roots: List[Union[str, Path]],
    actor: str,
    job_id: str,
) -> Dict[str, Any]:
    """Buoc 6 cua module docstring: issue_job() (Commit Broker chup baseline hash ngay truoc luc
    dong goi candidate) -> dong goi candidate ngay tai job_work_dir do Broker cap -> request_commit()
    (18 buoc day du, xem core/safety_gateway.py). KHONG tu viet lai bat ky phan nao cua 18 buoc."""
    ticket = issue_job(
        docx_path,
        actor=actor,
        task_id=job_id,
        allowed_roots=[str(r) for r in allowed_roots],
        work_dir=work_dir,
    )
    candidate_path = Path(ticket.work_dir) / f"candidate{docx_path.suffix}"
    _repack_document_xml(docx_path, root, candidate_path)
    commit_result = dict(request_commit(ticket.job_id, work_dir=work_dir))
    commit_result["job_id"] = ticket.job_id
    return commit_result


def _finalize(
    docx_path: Path,
    commit_result: Dict[str, Any],
    replace_result: TextReplaceResult,
    work_dir: Union[str, Path],
    job_id: str,
) -> PatchResult:
    """Buoc 7 cua module docstring: tu dong re-inspect SAU commit, cap document_revision/locators
    moi cho caller - khong bao gio tu "rebase" locator cu (plan muc 3.1)."""
    inspect_report = inspect_document(docx_path, work_dir=work_dir, job_id=job_id)
    return PatchResult(
        success=True,
        outcome=commit_result["outcome"],
        job_id=commit_result["job_id"],
        new_document_revision=inspect_report["document_revision"],
        new_locators=inspect_report["locators"],
        backup_path=commit_result.get("backup_path"),
        replace_mode=replace_result.run_mode,
    )


# ---------------------------------------------------------------------------
# Public API: patch_cell()
# ---------------------------------------------------------------------------
def patch_cell(
    docx_path: Union[str, Path],
    locator: ObjectLocator,
    new_text: str,
    work_dir: Union[str, Path],
    allowed_roots: List[Union[str, Path]],
    actor: str,
    job_id: str,
) -> PatchResult:
    """Thay toan bo noi dung van ban cua 1 o bang (locator.kind == "table_cell") bang new_text,
    giu nguyen w:tcPr va moi w:tc khac trong tai lieu (xem module docstring). Nem ValueError neu
    locator khong phai table_cell, CellSpansMultipleParagraphsError neu o co nhieu hon 1 doan van
    (Phase 1 MVP)."""
    docx_path = Path(docx_path)
    if locator.kind != "table_cell":
        raise ValueError(f"patch_cell() chi nhan locator.kind == 'table_cell', nhan duoc {locator.kind!r}")

    root, tc_element = _resolve_target(docx_path, locator)

    paragraphs = tc_element.findall(_P_TAG)
    if len(paragraphs) != 1:
        raise CellSpansMultipleParagraphsError(locator.object_id, len(paragraphs))

    replace_result = resolve_and_replace_text(paragraphs[0], locator.expected_text, new_text, None)

    commit_result = _stage_and_commit(docx_path, root, work_dir, allowed_roots, actor, job_id)
    return _finalize(docx_path, commit_result, replace_result, work_dir, job_id)


# ---------------------------------------------------------------------------
# Public API: patch_text()
# ---------------------------------------------------------------------------
def patch_text(
    docx_path: Union[str, Path],
    search_text: str,
    replace_text: str,
    locator: Optional[ObjectLocator],
    style_policy: Optional[str],
    work_dir: Union[str, Path],
    allowed_roots: List[Union[str, Path]],
    actor: str,
    job_id: str,
) -> PatchResult:
    """Tim search_text va thay bang replace_text. Neu locator duoc truyen, thu hep pham vi tim
    kiem vao dung 1 doan van (locator.kind == "paragraph") hoac cac doan van truc tiep ben trong 1
    o bang (locator.kind == "table_cell"); neu khong truyen locator, tim tren TOAN BO doan van cua
    tai lieu (ca than bai lan trong o bang) va bat buoc dung 1 doan van khop
    (AmbiguousTextMatchError neu qua 1, TextNotFoundError neu khong co). style_policy duoc chuyen
    thang cho resolve_and_replace_text() cua text_resolver.py khi span vat qua nhieu run khac
    style (xem core/text_resolver.py)."""
    docx_path = Path(docx_path)

    if locator is not None:
        root, target_element = _resolve_target(docx_path, locator)
        if locator.kind == "paragraph":
            candidate_paragraphs = [target_element]
        elif locator.kind == "table_cell":
            candidate_paragraphs = target_element.findall(_P_TAG)
        else:
            raise ValueError(f"locator.kind khong hop le: {locator.kind!r}")
    else:
        root = load_document_root(docx_path)
        body = root.find("w:body", _NSMAP)
        candidate_paragraphs = list(body.iter(_P_TAG)) if body is not None else []

    search_norm = unicodedata.normalize("NFC", search_text)
    matches = [p for p in candidate_paragraphs if search_norm in _element_text(p)]
    if not matches:
        raise TextNotFoundError(search_text)
    if len(matches) > 1:
        raise AmbiguousTextMatchError(search_text, len(matches))

    replace_result = resolve_and_replace_text(matches[0], search_text, replace_text, style_policy)

    commit_result = _stage_and_commit(docx_path, root, work_dir, allowed_roots, actor, job_id)
    return _finalize(docx_path, commit_result, replace_result, work_dir, job_id)
