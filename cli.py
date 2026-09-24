"""
tools/word-engine/cli.py - Entry-point duy nhat cua Word Engine (AI va TypeScript subprocess
consumers), xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md muc 1.2 ("Co che goi tu
TypeScript/Node") va muc 2 (cau truc thu muc, cli.py la "entrypoint hỗ trợ --json stdout
contract"). Module nay CHI la lop dieu phoi (orchestration) argparse GHEP LAI cac ham cong khai da
co san trong core/*.py - khong tu viet lai bat ky logic patch/inspect/backup/restore nao.

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" (xem
core/__init__.py va docs/plans muc 1.2).

Hop dong JSON stdout (khi co --json):
  - Thanh cong: exit code 0, DUY NHAT 1 dong JSON {"success": true, "outcome": ..., ...}.
  - Loi: exit code 1, DUY NHAT 1 dong JSON {"success": false, "error": "<TenException>",
    "reason": "<str(exception)>"}.
  - Khong in log/debug nao khac ra stdout khi co --json (moi thong bao khac, neu can, di stderr).
  - json.dumps() dung ensure_ascii=True (mac dinh) thay vi False nhu cac file khac trong core/ -
    quyet dinh co chu dich rieng cho DAY LA DUONG BIEN gui qua stdout/subprocess: cac core/*.py
    khac ghi thang ra file da mo voi encoding="utf-8" tuong minh nen an toan giu dau tieng Viet,
    nhung stdout cua 1 tien trinh con tren Windows co the dung code page khac utf-8
    (subprocess.run(..., text=True) mac dinh giai ma bang locale.getpreferredencoding(), thuong
    KHONG phai utf-8 tren Windows) - escape ve \\uXXXX thuan ASCII loai bo hoan toan rui ro
    UnicodeEncodeError/garbled text ma khong lam JSON kem chinh xac hon (moi JSON parser doc
    \\uXXXX dung nhu ky tu goc).

Quy uoc mac dinh khi CLI khong duoc truyen tuong minh (khong co flag --allowed-roots trong de
bai, nen day la quyet dinh mac dinh CO CHU DICH cua module nay, khong phai tham so nguoi dung
chinh):
  - work_dir: thu muc cha cua docx_path (".jarvis/" cua Commit Broker duoc neo ngay canh ho so
    dang sua, dung tinh than "du lieu ho so ... la du lieu cua mot case cu the" o CLAUDE.md goc).
  - allowed_roots cho issue_job(): [thu muc cha cua docx_path, REPO_ROOT cua chinh repo nay] -
    dung nhu de bai "Tu dong lay allowed_roots mac dinh la thu muc cha cua docx_path hoac
    REPO_ROOT".
  - job_id: uuid4 hex ngau nhien neu khong truyen --job-id.
  - actor: "word-engine-cli" neu khong truyen --actor.
"""

from __future__ import annotations

import argparse
import json
import sys
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

_SELF_DIR = Path(__file__).resolve().parent
if str(_SELF_DIR) not in sys.path:
    sys.path.insert(0, str(_SELF_DIR))

from core.inspector import (  # noqa: E402
    DocumentDriftError,
    LocatorNotFoundError,
    ObjectLocator,
    inspect_document,
)
from core.com_supervisor import export_pdf as com_export_pdf  # noqa: E402
from core.com_supervisor import inspect_vml_signature  # noqa: E402
from core.geometry_styler import apply_geometry  # noqa: E402
from core.safety_gateway import list_backups, restore_backup  # noqa: E402
from core.stamp_ops import apply_stamp  # noqa: E402
from core.surgical_patcher import patch_cell, patch_text  # noqa: E402

DEFAULT_ACTOR = "word-engine-cli"
REPO_ROOT = _SELF_DIR.parent.parent

# Raw argv duoc dung de parse trong lan goi main() gan nhat - _JsonAwareArgumentParser.error() doc
# bien module-level nay (khong doc sys.argv truc tiep) de hanh vi dung nhat quan bat ke main()
# duoc goi voi argv=None (mac dinh sys.argv[1:]) hay voi 1 list argv tuong minh (vd goi noi bo tu
# test). CLI la tien trinh don luong, khong co rui ro concurrency o day.
_RAW_ARGV: List[str] = []


class _JsonAwareArgumentParser(argparse.ArgumentParser):
    """Ghi de error(): argparse mac dinh in usage/message ra stderr roi sys.exit(2) khi cu phap
    sai (thieu tham so bat buoc, sai subcommand...). Khi --json co mat trong argv goc, in DUNG 1
    dong JSON loi ra STDOUT va exit 1 thay vi vay - giu dung hop dong JSON stdout cua toan bo CLI
    nay (xem module docstring) ngay ca khi loi xay ra TRUOC khi bat ky command handler nao duoc
    goi toi (vd argparse tu choi ngay luc parse, chua kip vao try/except cua main()).
    add_subparsers() tu dong dung parser_class=type(self) cho moi subparser (hanh vi mac dinh cua
    argparse), nen moi subparser sinh ra tu parser nay deu ke thua dung override nay - khong can
    truyen tuong minh."""

    def error(self, message: str) -> None:  # type: ignore[override]
        if "--json" in _RAW_ARGV:
            print(json.dumps({"success": False, "error": "ArgumentError", "reason": message}))
            sys.exit(1)
        super().error(message)


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------
def _default_work_dir(docx_path: Path, explicit: Optional[str]) -> Path:
    return Path(explicit).resolve() if explicit else docx_path.parent


def _default_allowed_roots(docx_path: Path) -> List[str]:
    return [str(docx_path.parent), str(REPO_ROOT)]


def _resolve_locator_or_raise(
    docx_path: Path, work_dir: Path, job_id: str, target_id: str, expected_revision: str
) -> ObjectLocator:
    """Quet lai docx_path (inspect_document() tu no da tu kiem tra state machine bang cach so
    document_revision moi voi noi dung dia hien tai - luon dung), roi doi chieu document_revision
    vua quet voi expected_revision nguoi goi CLI truyen vao: khac nhau nghia la dia da bi sua ke tu
    lan inspect truoc do (DOCUMENT_DRIFT). Tim object_id == target_id trong ket qua quet moi nay -
    khong doc lai sidecar index cu, vi mot ban quet moi tren dung noi dung hien tai la nguon su
    that duy nhat can thiet o day."""
    report = inspect_document(docx_path, work_dir=work_dir, job_id=job_id)
    if report["document_revision"] != expected_revision:
        raise DocumentDriftError(expected_revision, report["document_revision"])
    locator_dict = next((loc for loc in report["locators"] if loc["object_id"] == target_id), None)
    if locator_dict is None:
        raise LocatorNotFoundError(target_id)
    return ObjectLocator(**locator_dict)


# ---------------------------------------------------------------------------
# Command handlers - moi ham tra ve payload thanh cong (dict) hoac nem exception
# ---------------------------------------------------------------------------
def _cmd_inspect(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    work_dir = _default_work_dir(docx_path, args.work_dir)
    job_id = args.job_id or uuid.uuid4().hex

    report = inspect_document(docx_path, work_dir=work_dir, job_id=job_id)
    return {
        "success": True,
        "outcome": "inspected",
        "document_revision": report["document_revision"],
        "locators": report["locators"],
        "job_id": job_id,
        "index_file_path": report["index_file_path"],
    }


def _cmd_patch_cell(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    work_dir = _default_work_dir(docx_path, args.work_dir)
    job_id = args.job_id or uuid.uuid4().hex
    actor = args.actor or DEFAULT_ACTOR

    locator = _resolve_locator_or_raise(docx_path, work_dir, job_id, args.target_id, args.expected_revision)
    result = patch_cell(
        docx_path,
        locator,
        args.new_text,
        work_dir=work_dir,
        allowed_roots=_default_allowed_roots(docx_path),
        actor=actor,
        job_id=job_id,
    )
    return {
        "success": True,
        "outcome": result.outcome,
        "document_revision": result.new_document_revision,
        "locators": result.new_locators,
        "job_id": result.job_id,
        "backup_path": result.backup_path,
        "replace_mode": result.replace_mode,
    }


def _cmd_patch_text(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    work_dir = _default_work_dir(docx_path, args.work_dir)
    job_id = args.job_id or uuid.uuid4().hex
    actor = args.actor or DEFAULT_ACTOR

    if bool(args.target_id) != bool(args.expected_revision):
        raise ValueError("--target-id va --expected-revision phai duoc truyen cung nhau hoac deu bo trong.")

    locator: Optional[ObjectLocator] = None
    if args.target_id:
        locator = _resolve_locator_or_raise(docx_path, work_dir, job_id, args.target_id, args.expected_revision)

    result = patch_text(
        docx_path,
        args.search,
        args.replace,
        locator,
        args.style_policy,
        work_dir=work_dir,
        allowed_roots=_default_allowed_roots(docx_path),
        actor=actor,
        job_id=job_id,
    )
    return {
        "success": True,
        "outcome": result.outcome,
        "document_revision": result.new_document_revision,
        "locators": result.new_locators,
        "job_id": result.job_id,
        "backup_path": result.backup_path,
        "replace_mode": result.replace_mode,
    }


def _cmd_backups(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    work_dir = _default_work_dir(docx_path, args.work_dir)

    backups = list_backups(docx_path, work_dir=work_dir)
    return {"success": True, "outcome": "listed", "backups": backups}


def _cmd_restore(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    work_dir = _default_work_dir(docx_path, args.work_dir)
    actor = args.actor or DEFAULT_ACTOR

    result = restore_backup(
        docx_path,
        args.backup_id,
        work_dir=work_dir,
        allowed_roots=_default_allowed_roots(docx_path),
        actor=actor,
    )
    return {
        "success": True,
        "outcome": result["outcome"],
        "document_revision": f"sha256:{result['sha256']}",
        "job_id": result["job_id"],
        "backup_path": result.get("backup_path"),
        "restored_from_backup_id": result.get("restored_from_backup_id"),
    }


def _cmd_set_geometry(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    work_dir = _default_work_dir(docx_path, args.work_dir)
    job_id = args.job_id or uuid.uuid4().hex
    actor = args.actor or DEFAULT_ACTOR

    target_locator: Optional[ObjectLocator] = None
    target_id = getattr(args, "target_id", None) or getattr(args, "table_id", None)
    expected_revision = getattr(args, "expected_revision", None)
    if target_id:
        if expected_revision:
            target_locator = _resolve_locator_or_raise(docx_path, work_dir, job_id, target_id, expected_revision)
        else:
            report = inspect_document(docx_path, work_dir=work_dir, job_id=job_id)
            locator_dict = next((loc for loc in report["locators"] if loc["object_id"] == target_id), None)
            if locator_dict is None:
                raise LocatorNotFoundError(target_id)
            target_locator = ObjectLocator(**locator_dict)

    result = apply_geometry(
        docx_path,
        page_size=args.page_size,
        margins=args.margins,
        pagination=args.pagination,
        borders_preset=args.borders,
        table_locator=target_locator if (target_locator and target_locator.kind == "table_cell") else None,
        target_locator=target_locator,
        align=getattr(args, "align", None),
        balance=getattr(args, "balance", False),
        expand_transcripts=getattr(args, "expand_transcripts", False),
        row_height_dxa=getattr(args, "row_height_dxa", None),
        work_dir=work_dir,
        allowed_roots=_default_allowed_roots(docx_path),
        actor=actor,
        job_id=job_id,
    )
    return {
        "success": True,
        "outcome": result.outcome,
        "document_revision": result.new_document_revision,
        "locators": result.new_locators,
        "job_id": result.job_id,
        "backup_path": result.backup_path,
    }


def _cmd_stamp_ops(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    work_dir = _default_work_dir(docx_path, args.work_dir)
    job_id = args.job_id or uuid.uuid4().hex
    actor = args.actor or DEFAULT_ACTOR

    locator = _resolve_locator_or_raise(docx_path, work_dir, job_id, args.target_id, args.expected_revision)
    result = apply_stamp(
        docx_path,
        Path(args.asset_path).resolve(),
        locator,
        args.width_mm,
        args.height_mm,
        work_dir=work_dir,
        allowed_roots=_default_allowed_roots(docx_path),
        actor=actor,
        job_id=job_id,
        allow_unregistered=args.allow_unregistered,
    )
    return {
        "success": True,
        "outcome": result.outcome,
        "document_revision": result.new_document_revision,
        "locators": result.new_locators,
        "job_id": result.job_id,
        "backup_path": result.backup_path,
    }


def _cmd_export_pdf(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    pdf_path = Path(args.output).resolve() if args.output else docx_path.with_suffix(".pdf")

    result_path = com_export_pdf(docx_path, pdf_path)
    return {"success": True, "outcome": "exported_pdf", "pdf_path": str(result_path)}


def _cmd_validate_vml(args: argparse.Namespace) -> Dict[str, Any]:
    docx_path = Path(args.docx_path).resolve()
    signature = inspect_vml_signature(docx_path)
    return {"success": True, "outcome": "vml_signature", **signature}


_COMMANDS = {
    "inspect": _cmd_inspect,
    "patch-cell": _cmd_patch_cell,
    "patch-text": _cmd_patch_text,
    "backups": _cmd_backups,
    "restore": _cmd_restore,
    "set-geometry": _cmd_set_geometry,
    "stamp-ops": _cmd_stamp_ops,
    "export-pdf": _cmd_export_pdf,
    "validate-vml": _cmd_validate_vml,
}


# ---------------------------------------------------------------------------
# argparse wiring
# ---------------------------------------------------------------------------
def _build_parser() -> argparse.ArgumentParser:
    parser = _JsonAwareArgumentParser(
        prog="word-engine",
        description="Runtime/CLI chuyen dung xu ly Word (tools/word-engine/).",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_inspect = sub.add_parser("inspect", help="Quet document, sinh locators.")
    p_inspect.add_argument("docx_path")
    p_inspect.add_argument("--json", action="store_true")
    p_inspect.add_argument("--job-id")
    p_inspect.add_argument("--work-dir")

    p_cell = sub.add_parser("patch-cell", help="Sua noi dung 1 o bang (table_cell).")
    p_cell.add_argument("docx_path")
    p_cell.add_argument("--target-id", required=True)
    p_cell.add_argument("--new-text", required=True)
    p_cell.add_argument("--expected-revision", required=True)
    p_cell.add_argument("--json", action="store_true")
    p_cell.add_argument("--job-id")
    p_cell.add_argument("--work-dir")
    p_cell.add_argument("--actor")

    p_text = sub.add_parser("patch-text", help="Tim va thay the van ban.")
    p_text.add_argument("docx_path")
    p_text.add_argument("--search", required=True)
    p_text.add_argument("--replace", required=True)
    p_text.add_argument("--target-id")
    p_text.add_argument("--expected-revision")
    p_text.add_argument("--style-policy", choices=["prefer-start", "prefer-end"])
    p_text.add_argument("--json", action="store_true")
    p_text.add_argument("--job-id")
    p_text.add_argument("--work-dir")
    p_text.add_argument("--actor")

    p_backups = sub.add_parser("backups", help="Liet ke cac ban sao luu.")
    p_backups.add_argument("docx_path")
    p_backups.add_argument("--json", action="store_true")
    p_backups.add_argument("--work-dir")

    p_restore = sub.add_parser("restore", help="Phuc hoi tu 1 ban sao luu.")
    p_restore.add_argument("docx_path")
    p_restore.add_argument("--backup-id", required=True)
    p_restore.add_argument("--json", action="store_true")
    p_restore.add_argument("--work-dir")
    p_restore.add_argument("--actor")

    p_geom = sub.add_parser("set-geometry", help="Chuan hoa kho giay/le/dan trang/vien bang.")
    p_geom.add_argument("docx_path")
    p_geom.add_argument("--page-size", choices=["A4", "A5"])
    p_geom.add_argument("--margins", choices=["notary", "standard", "compact"])
    p_geom.add_argument("--pagination", action="store_true")
    p_geom.add_argument("--borders", choices=["all", "none", "horizontal-only", "notary-standard"])
    p_geom.add_argument("--align", choices=["left", "center", "right", "both"])
    p_geom.add_argument("--balance", action="store_true", help="Can chinh lai luoi cot va chieu rong cac o cua bang bi lech.")
    p_geom.add_argument("--expand-transcripts", action="store_true", help="Keo dan chieu cao hang de full trang A4.")
    p_geom.add_argument("--row-height-dxa", type=int, help="Chieu cao hang toi thieu (dxa).")
    p_geom.add_argument("--target-id")
    p_geom.add_argument("--expected-revision")
    p_geom.add_argument("--table-id")
    p_geom.add_argument("--json", action="store_true")
    p_geom.add_argument("--work-dir")
    p_geom.add_argument("--job-id")
    p_geom.add_argument("--actor")

    p_stamp = sub.add_parser(
        "stamp-ops",
        help="Chen anh dau/chu ky vao 1 doan van (DrawingML), qua whitelist SHA-256 & so kiem toan xuat xu.",
    )
    p_stamp.add_argument("docx_path")
    p_stamp.add_argument("--asset-path", required=True)
    # --target-id + --expected-revision deu bat buoc (khac vi du rut gon trong nhiem vu goc) de
    # giu dung bat bien "Revision-bound Object Locators" ma moi subcommand ghi khac (patch-cell/
    # patch-text/set-geometry) da cuong che: khong cho phep 1 duong ghi khong qua kiem tra
    # DOCUMENT_DRIFT nao rieng cho stamp-ops.
    p_stamp.add_argument("--target-id", required=True)
    p_stamp.add_argument("--expected-revision", required=True)
    p_stamp.add_argument("--width-mm", type=float, default=30.0)
    p_stamp.add_argument("--height-mm", type=float, default=30.0)
    p_stamp.add_argument("--allow-unregistered", action="store_true")
    p_stamp.add_argument("--json", action="store_true")
    p_stamp.add_argument("--job-id")
    p_stamp.add_argument("--work-dir")
    p_stamp.add_argument("--actor")

    p_export_pdf = sub.add_parser(
        "export-pdf", help="Xuat DOCX sang PDF qua Word COM native ExportAsFixedFormat (Phase 4)."
    )
    p_export_pdf.add_argument("docx_path")
    p_export_pdf.add_argument("--output")
    p_export_pdf.add_argument("--json", action="store_true")

    p_validate_vml = sub.add_parser(
        "validate-vml", help="Doc chu ky cau truc VML (w:pict/v:shape/v:imagedata) cua 1 docx."
    )
    p_validate_vml.add_argument("docx_path")
    p_validate_vml.add_argument("--json", action="store_true")

    return parser


def _print_human(command: str, payload: Dict[str, Any]) -> None:
    print(f"[{command}] outcome={payload.get('outcome')}")
    for key, value in payload.items():
        if key in ("outcome", "success"):
            continue
        print(f"  {key}: {value}")


def main(argv: Optional[List[str]] = None) -> int:
    global _RAW_ARGV
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    _RAW_ARGV = list(argv) if argv is not None else sys.argv[1:]

    parser = _build_parser()
    args = parser.parse_args(_RAW_ARGV)
    handler = _COMMANDS[args.command]

    try:
        payload = handler(args)
    except Exception as exc:  # noqa: BLE001 - CLI boundary: moi loi deu phai hoa thanh JSON contract
        error_payload = {"success": False, "error": type(exc).__name__, "reason": str(exc)}
        if args.json:
            print(json.dumps(error_payload))
        else:
            print(f"ERROR [{error_payload['error']}]: {error_payload['reason']}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(payload))
    else:
        _print_human(args.command, payload)
    return 0


if __name__ == "__main__":
    sys.exit(main())
