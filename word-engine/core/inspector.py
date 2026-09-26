"""
tools/word-engine/core/inspector.py - Revision-bound Object Locators cho tools/word-engine/.

Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 3.1 ("Revision-bound Object Locators
& State Machine Chuan"): quet toan bo w:body/w:p va w:body/w:tbl cua word/document.xml, sinh cho
moi doi tuong mot locator opaque (object_id) gan voi document_revision (sha256 toan bo file
.docx luc quet) va context_sha256 (sha256 tren mot canonical payload mo ta noi dung + ngu canh
lan can cua doi tuong do). Locator duoc luu vao sidecar `.jarvis/work/<job-id>/document-index.json`
- KHONG bao gio ghi object_id/context_sha256 vao chinh file .docx.

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" (xem
core/__init__.py va docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 1.2). Tai su dung noi
bo dung du an: sha256_file, atomic_write_bytes va canonical_json_dumps lay tu
core/safety_gateway.py cua chinh tools/word-engine/ thay vi viet lai.

Quyet dinh thiet ke ve context_sha256 (khong lap lai nguyen van vi du JSON minh hoa trong ke
hoach, vi vi du do dung trung 2 cach doc "1-based structural index" va "0-based noi bo" cho cung
mot con so nen khong the dung lam fixture dung nghia): payload dung de bam KHONG duoc chua bat ky
chi so vi tri nao co the tu dich chuyen khi mot "giao dich noi bo" (vd mot patch khac chen/xoa mot
doan van hoac mot hang bang o noi khac trong cung tai lieu) lam lech structural_path - neu khong
fallback re-scan se khong bao gio khop lai duoc dung doi tuong da dich chuyen. Cu the:
  - paragraph: kind + van ban chinh no + van ban doan truoc/sau no trong danh sach w:p cua body
    (KHONG chua paragraph index).
  - table_cell: kind + table_index (0-based, on dinh tuong doi vi cell khong doi bang) + van ban
    chinh no + van ban o truoc/sau trong CUNG HANG (KHONG chua row_index/col_index, vi day chinh
    la nhung gia tri se dich chuyen khi mot hang khac trong cung bang bi chen/xoa).
Han che da biet: 2 doi tuong co noi dung + ngu canh lan can giong het nhau se cho cung mot
context_sha256 - fallback se coi day la AMBIGUOUS_LOCATOR va tu choi doan bua (dung tinh than
"khong bia dat du kien" cua repo), khong phai loi.

Pham vi Phase 1 MVP: chi doc paragraphs/tables la CON TRUC TIEP cua w:body (dung nhu de bai section
3.1 - "w:body/w:p" va "w:body/w:tbl"), chua xu ly bang long trong o (nested table) hay doan van
long trong o cua bang.
"""

from __future__ import annotations

import json
import re
import unicodedata
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from lxml import etree

from core.safety_gateway import atomic_write_bytes, canonical_json_dumps, sha256_file

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
_NSMAP = {"w": WORD_NS}

DOCUMENT_INDEX_SCHEMA = "word-engine-document-index.v1"

_PATH_SEGMENT_RE = re.compile(r"^w:([A-Za-z0-9]+)\[(\d+)\]$")


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------
class DocumentDriftError(RuntimeError):
    """ABORT: DOCUMENT_DRIFT - hash dia hien tai khac voi document_revision da ghi trong locator
    (nguoi dung/tien trinh khac da sua file .docx ke tu luc inspect_document() quet)."""

    def __init__(self, expected_revision: str, actual_revision: str):
        super().__init__(
            f"DOCUMENT_DRIFT: locator revision {expected_revision} != current disk revision {actual_revision}"
        )
        self.expected_revision = expected_revision
        self.actual_revision = actual_revision


class LocatorNotFoundError(RuntimeError):
    def __init__(self, object_id: str):
        super().__init__(f"Khong tim thay doi tuong nao cho object_id {object_id}")
        self.object_id = object_id


class AmbiguousLocatorError(RuntimeError):
    """ABORT: AMBIGUOUS_LOCATOR - fallback re-scan tim thay khong dung 1 ket qua."""

    def __init__(self, object_id: str, match_count: int):
        super().__init__(
            f"AMBIGUOUS_LOCATOR: {match_count} doi tuong khop object_id {object_id}, can dung 1"
        )
        self.object_id = object_id
        self.match_count = match_count


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
@dataclass
class ObjectLocator:
    object_id: str
    revision: str
    kind: str  # "paragraph" | "table_cell"
    structural_path: str
    expected_text: str
    context_sha256: str
    w14_para_id: Optional[str] = None
    table_index: Optional[int] = None  # 0-based noi bo, chi co o kind == "table_cell"
    row_index: Optional[int] = None  # 0-based noi bo, chi co o kind == "table_cell"
    col_index: Optional[int] = None  # 0-based noi bo, chi co o kind == "table_cell"


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------
def load_document_root(docx_path: Union[str, Path]) -> etree._Element:
    """Giai nen word/document.xml trong bo nho (khong ghi ra dia) va parse bang lxml."""
    with zipfile.ZipFile(Path(docx_path), "r") as zf:
        xml_bytes = zf.read("word/document.xml")
    return etree.fromstring(xml_bytes)


def _text_of(element: etree._Element) -> str:
    parts = [t.text or "" for t in element.iter(f"{{{WORD_NS}}}t")]
    return unicodedata.normalize("NFC", "".join(parts))


def _context_sha256(payload: Dict[str, Any]) -> str:
    return "sha256:" + sha256(canonical_json_dumps(payload).encode("utf-8")).hexdigest()


def _lookup_structural_path(root: etree._Element, structural_path: str) -> Optional[etree._Element]:
    """Fast structural lookup: di theo dung cac chi so 1-based trong structural_path (chuan
    XPath OOXML), khong lam gi ngoai dieu huong theo vi tri - khong tu doan/sua neu lech."""
    segments = [s for s in structural_path.split("/") if s]
    if not segments or segments[0] != "w:document":
        return None
    node = root
    for segment in segments[1:]:
        if segment == "w:body":
            node = node.find("w:body", _NSMAP)
            if node is None:
                return None
            continue
        match = _PATH_SEGMENT_RE.match(segment)
        if not match:
            return None
        tag, index_str = match.group(1), match.group(2)
        index = int(index_str)
        children = node.findall(f"w:{tag}", _NSMAP)
        if index < 1 or index > len(children):
            return None
        node = children[index - 1]
    return node


# ---------------------------------------------------------------------------
# Scan: sinh toan bo locators cho 1 document root da parse
# ---------------------------------------------------------------------------
def _scan(root: etree._Element, document_revision: str) -> List[ObjectLocator]:
    body = root.find("w:body", _NSMAP)
    if body is None:
        return []

    locators: List[ObjectLocator] = []

    paragraphs = body.findall("w:p", _NSMAP)
    paragraph_texts = [_text_of(p) for p in paragraphs]
    for i, p in enumerate(paragraphs, start=1):
        text = paragraph_texts[i - 1]
        prev_text = paragraph_texts[i - 2] if i - 2 >= 0 else None
        next_text = paragraph_texts[i] if i < len(paragraphs) else None
        payload = {"kind": "paragraph", "text": text, "prev_text": prev_text, "next_text": next_text}
        locators.append(
            ObjectLocator(
                object_id=f"para_{i}",
                revision=document_revision,
                kind="paragraph",
                structural_path=f"/w:document/w:body/w:p[{i}]",
                expected_text=text,
                context_sha256=_context_sha256(payload),
                w14_para_id=p.get(f"{{{W14_NS}}}paraId"),
            )
        )

    tables = body.findall("w:tbl", _NSMAP)
    for ti, tbl in enumerate(tables, start=1):
        rows = tbl.findall("w:tr", _NSMAP)
        rows_cells = [row.findall("w:tc", _NSMAP) for row in rows]
        rows_texts = [[_text_of(tc) for tc in cells] for cells in rows_cells]

        for rj, row in enumerate(rows, start=1):
            cells = rows_cells[rj - 1]
            texts = rows_texts[rj - 1]
            w14_para_id = row.get(f"{{{W14_NS}}}paraId")
            for ck, tc in enumerate(cells, start=1):
                text = texts[ck - 1]
                prev_cell_text = texts[ck - 2] if ck - 2 >= 0 else None
                next_cell_text = texts[ck] if ck < len(texts) else None
                payload = {
                    "kind": "table_cell",
                    "table_index": ti - 1,
                    "text": text,
                    "prev_cell_text": prev_cell_text,
                    "next_cell_text": next_cell_text,
                }
                locators.append(
                    ObjectLocator(
                        object_id=f"cell_t{ti}_r{rj}_c{ck}",
                        revision=document_revision,
                        kind="table_cell",
                        structural_path=f"/w:document/w:body/w:tbl[{ti}]/w:tr[{rj}]/w:tc[{ck}]",
                        expected_text=text,
                        context_sha256=_context_sha256(payload),
                        w14_para_id=w14_para_id,
                        table_index=ti - 1,
                        row_index=rj - 1,
                        col_index=ck - 1,
                    )
                )

    return locators


# ---------------------------------------------------------------------------
# Public API: inspect_document()
# ---------------------------------------------------------------------------
def inspect_document(docx_path: Union[str, Path], work_dir: Union[str, Path], job_id: str) -> Dict[str, Any]:
    """Quet docx_path, sinh Revision-bound Object Locators va ghi sidecar index vao
    `<work_dir>/.jarvis/work/<job_id>/document-index.json`. Tra ve InspectionReport dang dict:
    {document_revision, locators, index_file_path}."""
    docx_path = Path(docx_path)
    document_revision = f"sha256:{sha256_file(docx_path)}"
    root = load_document_root(docx_path)
    locators = _scan(root, document_revision)
    locator_dicts = [asdict(loc) for loc in locators]

    index_dir = Path(work_dir).resolve() / ".jarvis" / "work" / job_id
    index_dir.mkdir(parents=True, exist_ok=True)
    index_path = index_dir / "document-index.json"
    index_payload = {
        "schema": DOCUMENT_INDEX_SCHEMA,
        "job_id": job_id,
        "docx_path": str(docx_path.resolve()),
        "document_revision": document_revision,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "locators": locator_dicts,
    }
    atomic_write_bytes(index_path, (json.dumps(index_payload, indent=2, ensure_ascii=False) + "\n").encode("utf-8"))

    return {
        "document_revision": document_revision,
        "locators": locator_dicts,
        "index_file_path": str(index_path),
    }


# ---------------------------------------------------------------------------
# Public API: resolve locators (fast structural lookup + fallback re-scan)
# ---------------------------------------------------------------------------
def resolve_in_document(root: etree._Element, locator: Dict[str, Any]) -> etree._Element:
    """Tim lai element trong MOT root da parse san, dung dung nhanh "Thu structural_path (Fast
    structural lookup)" -> "Fallback re-scan theo object_id & context_sha256" cua state machine
    section 3.1 - KHONG kiem tra document_revision (caller da/se tu kiem tra o buoc rieng, vd
    resolve_locator() ben duoi khi lam viec truc tiep tren dia)."""
    element = _lookup_structural_path(root, locator["structural_path"])
    if element is not None and _text_of(element) == locator["expected_text"]:
        return element

    candidates = [loc for loc in _scan(root, locator["revision"]) if loc.context_sha256 == locator["context_sha256"]]
    if len(candidates) == 1:
        resolved = _lookup_structural_path(root, candidates[0].structural_path)
        if resolved is not None:
            return resolved
        raise LocatorNotFoundError(locator["object_id"])
    if not candidates:
        raise LocatorNotFoundError(locator["object_id"])
    raise AmbiguousLocatorError(locator["object_id"], len(candidates))


def resolve_locator(docx_path: Union[str, Path], locator: Dict[str, Any]) -> etree._Element:
    """Diem vao day du cua state machine section 3.1, tinh tu dau: kiem tra
    CURRENT_DISK_REVISION == LOCATOR_REVISION (ABORT: DOCUMENT_DRIFT neu khac), roi moi goi
    resolve_in_document() tren mot ban parse tuoi cua chinh docx_path do."""
    docx_path = Path(docx_path)
    current_revision = f"sha256:{sha256_file(docx_path)}"
    if current_revision != locator["revision"]:
        raise DocumentDriftError(locator["revision"], current_revision)
    root = load_document_root(docx_path)
    return resolve_in_document(root, locator)
