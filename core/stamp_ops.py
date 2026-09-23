"""
tools/word-engine/core/stamp_ops.py - Quan tri Con dau & Chu ky (Phase 3) cho tools/word-engine/.

Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md muc 3.6 (cong thuc mm->points cho Word COM,
Phase 4 - KHONG dung o day) va muc 4 ("Phase 3: Quan tri Con dau & Chu ky" - "Quan ly Audit Trail
tai `.jarvis/audit/stamp_provenance.jsonl` va whitelist hash anh"). Module nay chi xu ly nhanh
OOXML SURGICAL MODE (chen truc tiep DrawingML vao goi OPC qua Commit Broker) - KHONG dung Word COM
Shapes.AddPicture (do la Phase 4, core/com_supervisor.py).

Module nay la lop dieu phoi (orchestration) GHEP LAI cac primitive da co san cua chinh
tools/word-engine/ - khong tu viet lai Commit Broker/locator state machine:
  - core/inspector.py + core/surgical_patcher.py: `_resolve_target()` (state machine drift check +
    resolve_in_document() + kiem tra noi dung khop locator.expected_text) tai su dung nguyen ven
    tu core/surgical_patcher.py thay vi viet lai (nac 2 cua CLAUDE.md goc: tai su dung noi bo dung
    du an con dang sua truoc khi viet moi).
  - core/safety_gateway.py: issue_job()/request_commit() (Commit Broker 18 buoc da kiem toan),
    sha256_bytes, atomic_write_bytes, MicroLock, _jarvis_root (neo `.jarvis/` theo work_dir).

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" (xem
core/__init__.py va docs/plans muc 1.2).

Module nay chiu trach nhiem 3 phan doc lap:

1. **Uy quyen (Authorization)** - `is_asset_authorized()`/`register_stamp_asset()`: whitelist tai
   `.jarvis/security/stamp_allowlist.json` (neo theo work_dir, cung quy uoc voi backups/state/jobs/
   audit cua safety_gateway.py) luu danh sach SHA-256 cua cac anh dau/chu ky da duoc dang ky. SHA-256
   cua `--asset-path` KHONG co trong whitelist -> nem UnauthorizedStampAssetError, tru khi caller
   truyen `allow_unregistered=True` (CLI: `--allow-unregistered`) de bo qua co chu dich (vd luc dang
   ky lan dau/test). `register_stamp_asset()` la helper cong khai de dang ky 1 asset (dung trong
   test va cho nguoi van hanh dang ky truoc mot con dau/chu ky moi).

2. **So kiem toan xuat xu (Provenance Ledger)** - `_append_stamp_provenance()`: moi lan chen anh
   thanh cong, ghi them 1 dong JSON vao `.jarvis/audit/stamp_provenance.jsonl` (schema dung dung
   nhu de bai: timestamp, document_path, asset_sha256, authorized_by, inserted_dimensions, job_id) -
   tap tin JSONL rieng, KHONG dung chung `audit/<YYYY-MM>.jsonl` cua Commit Broker (do la audit
   trail cua BAN THAN Commit Broker - STARTED/STAGED/COMMITTED/FAILED - khac muc dich voi so kiem
   toan xuat xu con dau nay).

3. **Chen anh vao goi DOCX (DrawingML)** - `apply_stamp()`: dung dung `resolve_in_document()`/state
   machine cua inspector.py de xac dinh DUNG 1 doan van muc tieu (locator.kind == "paragraph"),
   them 1 w:r moi chua w:drawing (wp:inline chuan DrawingML) vao CUOI doan van do, dong thoi:
     - Them file anh vao `word/media/<ten-file-moi>` (ten khong trung voi media da co san).
     - Them 1 `<Relationship>` kieu image vao `word/_rels/document.xml.rels`.
     - Dam bao `[Content_Types].xml` co `<Default Extension="png|jpg|jpeg" ContentType="..."/>`
       tuong ung (chi them neu chua co, khong tao trung).
   KHAC voi core/surgical_patcher.py/core/geometry_styler.py (chi thay the DUY NHAT
   word/document.xml, moi part khac giu RAW SHA-256 IDENTICAL): thao tac chen anh BAT BUOC phai
   sua ca 3 part (document.xml, _rels/document.xml.rels, [Content_Types].xml) va THEM 1 part hoan
   toan moi (word/media/...), nen module nay tu viet ham dong goi rieng (`_repack_with_stamp()`)
   thay vi tai su dung `_stage_and_commit()`/`_repack_document_xml()` (2 ham do chi ho tro thay the
   dung 1 part co san la word/document.xml) - van goi thang `issue_job()`/`request_commit()` cua
   Commit Broker cho buoc commit, khong viet lai bat ky phan nao cua 18 buoc.

Gioi han Phase 3 (co chu dich, dung tinh than "khong bia dat du kien" cua repo):
  - Chi ho tro locator.kind == "paragraph" (dung nhu de bai: "Chen vao paragraph duoc chi dinh qua
    locator") - locator.kind == "table_cell" nem ValueError, khong tu doan chen vao o bang nao.
  - Chi ho tro dinh dang anh png/jpg/jpeg (Content-Type map co dinh _CONTENT_TYPE_BY_EXTENSION) -
    dinh dang khac nem ValueError, khong doan Content-Type.
  - KHONG doc pixel-size that su cua anh (vd qua Pillow) de suy ra ty le khung hinh - width_mm/
    height_mm la tham so BAT BUOC do caller truyen tuong minh, dung dung 2 gia tri do quy doi EMU
    (1 mm = 36000 EMU) cho ca wp:extent lan a:xfrm/a:ext. Neu caller truyen sai ty le so voi anh
    that, anh se bi keo dan/bop meo trong Word - day la trach nhiem cua caller, khong phai loi cua
    module nay tu suy doan sai.
"""

from __future__ import annotations

import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union

from lxml import etree

from core.inspector import ObjectLocator, WORD_NS, inspect_document
from core.safety_gateway import (
    MicroLock,
    atomic_write_bytes,
    issue_job,
    request_commit,
    sha256_bytes,
    _jarvis_root,
)
from core.surgical_patcher import PatchResult, _resolve_target

WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_IMAGE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

_DOCUMENT_PART = "word/document.xml"
_RELS_PART = "word/_rels/document.xml.rels"
_CONTENT_TYPES_PART = "[Content_Types].xml"

_CONTENT_TYPE_BY_EXTENSION = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}
_MM_TO_EMU = 36000

_ALLOWLIST_SCHEMA = "word-engine-stamp-allowlist.v1"
_PROVENANCE_LOG_NAME = "stamp_provenance.jsonl"


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------
class UnauthorizedStampAssetError(RuntimeError):
    """Nem boi apply_stamp() khi sha256 cua --asset-path khong co trong
    `.jarvis/security/stamp_allowlist.json` va caller khong truyen allow_unregistered=True."""

    def __init__(self, asset_sha256: str):
        super().__init__(
            f"UNAUTHORIZED_STAMP_ASSET: Asset sha256 is not registered in allowlist ({asset_sha256})"
        )
        self.asset_sha256 = asset_sha256


# ---------------------------------------------------------------------------
# Public API: Authorization (whitelist SHA-256)
# ---------------------------------------------------------------------------
def _allowlist_path(work_dir: Optional[Union[str, Path]]) -> Path:
    return _jarvis_root(work_dir) / "security" / "stamp_allowlist.json"


def _load_allowlist(work_dir: Optional[Union[str, Path]]) -> Dict[str, Any]:
    path = _allowlist_path(work_dir)
    if not path.exists():
        return {"schema": _ALLOWLIST_SCHEMA, "entries": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"schema": _ALLOWLIST_SCHEMA, "entries": []}
    return {"schema": _ALLOWLIST_SCHEMA, "entries": list(data.get("entries", []))}


def is_asset_authorized(asset_sha256: str, work_dir: Optional[Union[str, Path]] = None) -> bool:
    """True neu asset_sha256 da duoc dang ky trong `.jarvis/security/stamp_allowlist.json`."""
    allowlist = _load_allowlist(work_dir)
    return any(entry.get("sha256") == asset_sha256 for entry in allowlist["entries"])


def register_stamp_asset(
    asset_path: Union[str, Path],
    work_dir: Optional[Union[str, Path]] = None,
    label: Optional[str] = None,
    registered_by: str = "system",
) -> str:
    """Helper dang ky 1 anh dau/chu ky vao whitelist (dung trong test va cho nguoi van hanh dang
    ky truoc 1 con dau/chu ky moi qua CLI/script rieng) - tinh SHA-256 cua asset_path, them vao
    `.jarvis/security/stamp_allowlist.json` neu chua co (idempotent), tra ve sha256 da tinh."""
    asset_sha256 = sha256_bytes(Path(asset_path).read_bytes())
    path = _allowlist_path(work_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    lock_path = path.with_name(path.name + ".lock")
    with MicroLock(lock_path, timeout=5.0):
        allowlist = _load_allowlist(work_dir)
        if not any(entry.get("sha256") == asset_sha256 for entry in allowlist["entries"]):
            allowlist["entries"].append({
                "sha256": asset_sha256,
                "label": label,
                "registered_at": datetime.now(timezone.utc).isoformat(),
                "registered_by": registered_by,
            })
        atomic_write_bytes(path, (json.dumps(allowlist, indent=2, ensure_ascii=False) + "\n").encode("utf-8"))
    return asset_sha256


# ---------------------------------------------------------------------------
# Public API: Provenance Ledger
# ---------------------------------------------------------------------------
def _append_stamp_provenance(work_dir: Optional[Union[str, Path]], entry: Dict[str, Any]) -> Path:
    audit_dir = _jarvis_root(work_dir) / "audit"
    audit_dir.mkdir(parents=True, exist_ok=True)
    log_path = audit_dir / _PROVENANCE_LOG_NAME
    lock_path = audit_dir / f"{_PROVENANCE_LOG_NAME}.lock"
    with MicroLock(lock_path, timeout=5.0):
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return log_path


# ---------------------------------------------------------------------------
# Internal helpers: OPC part read/patch (rels, content-types, media)
# ---------------------------------------------------------------------------
def _extract_part(docx_path: Path, part_name: str) -> bytes:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return zf.read(part_name)


def _existing_media_names(docx_path: Path) -> Set[str]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return {name.rsplit("/", 1)[-1] for name in zf.namelist() if name.startswith("word/media/")}


def _next_media_filename(existing_names: Set[str], ext: str) -> str:
    n = 1
    while True:
        candidate = f"stamp_image{n}.{ext}"
        if candidate not in existing_names:
            return candidate
        n += 1


def _next_rel_id(rels_root: etree._Element) -> str:
    max_n = 0
    for rel in rels_root.findall(f"{{{_REL_NS}}}Relationship"):
        match = re.match(r"^rId(\d+)$", rel.get("Id") or "")
        if match:
            max_n = max(max_n, int(match.group(1)))
    return f"rId{max_n + 1}"


def _add_image_relationship(rels_root: etree._Element, target: str) -> str:
    rid = _next_rel_id(rels_root)
    rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", rid)
    rel.set("Type", _REL_IMAGE_TYPE)
    rel.set("Target", target)
    return rid


def _ensure_default_extension(content_types_root: etree._Element, extension: str, content_type: str) -> None:
    for default in content_types_root.findall(f"{{{_CT_NS}}}Default"):
        if (default.get("Extension") or "").lower() == extension:
            return
    default_el = etree.SubElement(content_types_root, f"{{{_CT_NS}}}Default")
    default_el.set("Extension", extension)
    default_el.set("ContentType", content_type)


def _next_docpr_id(document_root: etree._Element) -> int:
    max_id = 0
    for doc_pr in document_root.iter(f"{{{WP_NS}}}docPr"):
        try:
            max_id = max(max_id, int(doc_pr.get("id", "0")))
        except ValueError:
            continue
    return max_id + 1


# ---------------------------------------------------------------------------
# Internal helpers: DrawingML element (wp:inline -> a:graphic -> pic:pic)
# ---------------------------------------------------------------------------
def _build_drawing_element(rid: str, width_emu: int, height_emu: int, docpr_id: int, name: str) -> etree._Element:
    """Xay dung 1 w:drawing (wp:inline) chuan DrawingML cho 1 anh tinh, dung cau truc de bai yeu
    cau: wp:inline -> a:graphic -> a:graphicData -> pic:pic (pic:nvPicPr, pic:blipFill, pic:spPr).
    Kich thuoc width_emu/height_emu duoc dat dong thoi tren wp:extent VA a:xfrm/a:ext (2 noi Word
    doc kich thuoc hien thi, phai khop nhau)."""
    drawing = etree.Element(f"{{{WORD_NS}}}drawing")

    inline = etree.SubElement(drawing, f"{{{WP_NS}}}inline")
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")

    extent = etree.SubElement(inline, f"{{{WP_NS}}}extent")
    extent.set("cx", str(width_emu))
    extent.set("cy", str(height_emu))

    effect_extent = etree.SubElement(inline, f"{{{WP_NS}}}effectExtent")
    for attr in ("l", "t", "r", "b"):
        effect_extent.set(attr, "0")

    doc_pr = etree.SubElement(inline, f"{{{WP_NS}}}docPr")
    doc_pr.set("id", str(docpr_id))
    doc_pr.set("name", name)

    cnv_graphic_frame_pr = etree.SubElement(inline, f"{{{WP_NS}}}cNvGraphicFramePr")
    graphic_frame_locks = etree.SubElement(cnv_graphic_frame_pr, f"{{{A_NS}}}graphicFrameLocks")
    graphic_frame_locks.set("noChangeAspect", "1")

    graphic = etree.SubElement(inline, f"{{{A_NS}}}graphic")
    graphic_data = etree.SubElement(graphic, f"{{{A_NS}}}graphicData")
    graphic_data.set("uri", PIC_NS)

    pic = etree.SubElement(graphic_data, f"{{{PIC_NS}}}pic")

    nv_pic_pr = etree.SubElement(pic, f"{{{PIC_NS}}}nvPicPr")
    cnv_pr = etree.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPr")
    cnv_pr.set("id", "0")
    cnv_pr.set("name", name)
    etree.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPicPr")

    blip_fill = etree.SubElement(pic, f"{{{PIC_NS}}}blipFill")
    blip = etree.SubElement(blip_fill, f"{{{A_NS}}}blip")
    blip.set(f"{{{R_NS}}}embed", rid)
    stretch = etree.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    etree.SubElement(stretch, f"{{{A_NS}}}fillRect")

    sp_pr = etree.SubElement(pic, f"{{{PIC_NS}}}spPr")
    xfrm = etree.SubElement(sp_pr, f"{{{A_NS}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", "0")
    off.set("y", "0")
    ext_el = etree.SubElement(xfrm, f"{{{A_NS}}}ext")
    ext_el.set("cx", str(width_emu))
    ext_el.set("cy", str(height_emu))
    prst_geom = etree.SubElement(sp_pr, f"{{{A_NS}}}prstGeom")
    prst_geom.set("prst", "rect")
    etree.SubElement(prst_geom, f"{{{A_NS}}}avLst")

    return drawing


# ---------------------------------------------------------------------------
# Internal helper: dong goi lai OPC voi 3 part sua + 1 part moi (media)
# ---------------------------------------------------------------------------
def _repack_with_stamp(
    docx_path: Path,
    new_document_root: etree._Element,
    new_rels_bytes: bytes,
    new_content_types_bytes: bytes,
    media_filename: str,
    media_bytes: bytes,
    candidate_path: Path,
) -> None:
    """Dong goi lai toan bo goi OPC cua docx_path vao candidate_path: 3 part duoc THAY THE
    (word/document.xml, word/_rels/document.xml.rels, [Content_Types].xml), 1 part MOI duoc THEM
    (word/media/<media_filename>), moi part khac doc THANG tu docx_path tren dia va ghi lai y het -
    khac voi core/surgical_patcher.py::_repack_document_xml() (chi thay the DUY NHAT
    word/document.xml, khong them part nao), vi chen anh bat buoc phai sua ca rels/content-types
    va them 1 media part hoan toan moi (xem module docstring)."""
    new_document_bytes = etree.tostring(new_document_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    overrides = {
        _DOCUMENT_PART: new_document_bytes,
        _RELS_PART: new_rels_bytes,
        _CONTENT_TYPES_PART: new_content_types_bytes,
    }
    candidate_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "r") as src, zipfile.ZipFile(candidate_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = overrides.get(item.filename)
            if data is None:
                data = src.read(item.filename)
            dst.writestr(item, data)
        dst.writestr(f"word/media/{media_filename}", media_bytes)


# ---------------------------------------------------------------------------
# Public API: apply_stamp()
# ---------------------------------------------------------------------------
def apply_stamp(
    docx_path: Union[str, Path],
    asset_path: Union[str, Path],
    locator: ObjectLocator,
    width_mm: float,
    height_mm: float,
    work_dir: Union[str, Path],
    allowed_roots: List[Union[str, Path]],
    actor: str,
    job_id: str,
    allow_unregistered: bool = False,
) -> PatchResult:
    """Chen anh dau/chu ky (`asset_path`) vao CUOI doan van (`locator.kind == "paragraph"`) duoc
    chi dinh, sau khi xac thuc SHA-256 cua asset qua whitelist (xem module docstring "1. Uy
    quyen"). Ghi 1 dong vao so kiem toan xuat xu SAU khi commit thanh cong (xem "2. So kiem toan
    xuat xu"). Nem UnauthorizedStampAssetError neu asset chua duoc dang ky va allow_unregistered
    la False; nem ValueError neu locator.kind != "paragraph", dinh dang anh khong ho tro, hoac
    width_mm/height_mm <= 0; nem DocumentDriftError/ObjectContentMismatchError qua _resolve_target()
    (tai su dung tu core/surgical_patcher.py) neu tai lieu da bi sua tren dia ke tu luc locator
    duoc quet."""
    docx_path = Path(docx_path)
    asset_path = Path(asset_path)

    if locator.kind != "paragraph":
        raise ValueError(f"apply_stamp() chi nhan locator.kind == 'paragraph', nhan duoc {locator.kind!r}")
    if width_mm <= 0 or height_mm <= 0:
        raise ValueError("width_mm va height_mm phai > 0.")

    ext = asset_path.suffix.lower().lstrip(".")
    if ext not in _CONTENT_TYPE_BY_EXTENSION:
        raise ValueError(f"Dinh dang anh khong ho tro: {asset_path.suffix!r} (chi nhan png/jpg/jpeg).")

    asset_bytes = asset_path.read_bytes()
    asset_sha256 = sha256_bytes(asset_bytes)
    if not allow_unregistered and not is_asset_authorized(asset_sha256, work_dir):
        raise UnauthorizedStampAssetError(asset_sha256)

    # Buoc 1-3 cua state machine (xem core/surgical_patcher.py::_resolve_target module docstring):
    # DOCUMENT_DRIFT check + resolve_in_document() + kiem tra noi dung khop expected_text.
    root, paragraph = _resolve_target(docx_path, locator)

    rels_root = etree.fromstring(_extract_part(docx_path, _RELS_PART))
    content_types_root = etree.fromstring(_extract_part(docx_path, _CONTENT_TYPES_PART))

    media_filename = _next_media_filename(_existing_media_names(docx_path), ext)
    rid = _add_image_relationship(rels_root, f"media/{media_filename}")
    _ensure_default_extension(content_types_root, ext, _CONTENT_TYPE_BY_EXTENSION[ext])

    width_emu = round(width_mm * _MM_TO_EMU)
    height_emu = round(height_mm * _MM_TO_EMU)
    docpr_id = _next_docpr_id(root)
    drawing = _build_drawing_element(rid, width_emu, height_emu, docpr_id, asset_path.stem or "Stamp")

    run = etree.SubElement(paragraph, f"{{{WORD_NS}}}r")
    run.append(drawing)

    new_rels_bytes = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_content_types_bytes = etree.tostring(content_types_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    ticket = issue_job(
        docx_path,
        actor=actor,
        task_id=job_id,
        allowed_roots=[str(r) for r in allowed_roots],
        work_dir=work_dir,
    )
    candidate_path = Path(ticket.work_dir) / f"candidate{docx_path.suffix}"
    _repack_with_stamp(docx_path, root, new_rels_bytes, new_content_types_bytes, media_filename, asset_bytes, candidate_path)
    commit_result = dict(request_commit(ticket.job_id, work_dir=work_dir))
    commit_result["job_id"] = ticket.job_id

    _append_stamp_provenance(work_dir, {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "document_path": str(docx_path),
        "asset_sha256": asset_sha256,
        "authorized_by": actor,
        "inserted_dimensions": {"width_mm": width_mm, "height_mm": height_mm},
        "job_id": commit_result["job_id"],
    })

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
