"""
tools/excel-engine/cli.py - Entry-point duy nhat cua Excel Engine (AI va subprocess consumers).
Module nay CHI la lop dieu phoi argparse ghep cac ham cong khai trong core/*.py.

KHONG import bat ky module nao tu word-engine/, dich-thuat/, phan-tich/ hay "nhan vien ho so/".

Hop dong JSON stdout (khi co --json), giong word-engine:
  - Thanh cong: exit code 0, DUY NHAT 1 dong JSON {"success": true, "outcome": ..., ...}.
  - Loi: exit code 1, DUY NHAT 1 dong JSON {"success": false, "error": "<TenException>",
    "reason": "<str(exception)>"}.
  - json.dumps() dung ensure_ascii=True (mac dinh): stdout cua tien trinh con tren Windows co the
    khong phai utf-8, \\uXXXX thuan ASCII tranh UnicodeEncodeError/garbled text.

Quy uoc mac dinh: work_dir = thu muc cha cua file (".jarvis/" duoc neo canh file dang sua);
allowed_roots = [thu muc cha cua file]; actor = "excel-engine-cli".
Lenh doc (inspect/read-sheet/read-cell) tuyet doi khong ghi file nao.
Cac lenh patch-* BAT BUOC --expected-revision (document_revision lay tu inspect/read-*).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

_SELF_DIR = Path(__file__).resolve().parent
if str(_SELF_DIR) not in sys.path:
    sys.path.insert(0, str(_SELF_DIR))

from core.inspector import inspect_workbook, read_cell, read_sheet  # noqa: E402
from core.safety_gateway import list_backups, restore_backup  # noqa: E402
from core.surgical_patcher import (  # noqa: E402
    CellOp,
    InvalidValueError,
    batch_patch,
    op_from_dict,
    patch_range,
)

DEFAULT_ACTOR = "excel-engine-cli"
_RAW_ARGV: List[str] = []


class _JsonAwareArgumentParser(argparse.ArgumentParser):
    """Khi --json co mat trong argv, loi cu phap cung tra ve dung 1 dong JSON tren stdout (exit 1)."""

    def error(self, message: str) -> None:  # type: ignore[override]
        if "--json" in _RAW_ARGV:
            print(json.dumps({"success": False, "error": "ArgumentError", "reason": message}))
            sys.exit(1)
        super().error(message)


def _work_dir(xlsx_path: Path, explicit: Optional[str]) -> Path:
    return Path(explicit).resolve() if explicit else xlsx_path.parent


def _roots(xlsx_path: Path) -> List[str]:
    return [str(xlsx_path.parent)]


def _typed_value(raw: Optional[str], kind: str) -> Any:
    if kind == "blank":
        return None
    if raw is None:
        raise InvalidValueError("--value la bat buoc tru khi --type blank.")
    if kind == "number":
        try:
            return int(raw) if raw.lstrip("-").isdigit() else float(raw)
        except ValueError:
            raise InvalidValueError(f"Khong doc duoc so: {raw!r}") from None
    if kind == "bool":
        if raw.lower() not in ("true", "false"):
            raise InvalidValueError(f"bool phai la true|false: {raw!r}")
        return raw.lower() == "true"
    return raw  # text | formula


def _patch_payload(result: Any) -> Dict[str, Any]:
    return {
        "success": True,
        "outcome": result.outcome,
        "document_revision": result.new_document_revision,
        "job_id": result.job_id,
        "backup_path": result.backup_path,
        "ops_applied": result.ops_applied,
        "modified_parts": result.modified_parts,
        "removed_parts": result.removed_parts,
        "calc_chain_removed": result.calc_chain_removed,
        "unchanged_part_count": result.unchanged_part_count,
    }


def _patch_kwargs(args: argparse.Namespace, path: Path) -> Dict[str, Any]:
    return {
        "work_dir": _work_dir(path, args.work_dir),
        "allowed_roots": _roots(path),
        "actor": args.actor or DEFAULT_ACTOR,
        "truncate": args.truncate,
    }


# ---------------------------------------------------------------------------
# Command handlers
# ---------------------------------------------------------------------------
def _cmd_inspect(args: argparse.Namespace) -> Dict[str, Any]:
    return {"success": True, "outcome": "inspected", **inspect_workbook(Path(args.xlsx_path).resolve(), args.limit)}


def _cmd_read_sheet(args: argparse.Namespace) -> Dict[str, Any]:
    return {"success": True, "outcome": "read_sheet", **read_sheet(Path(args.xlsx_path).resolve(), args.sheet, args.range)}


def _cmd_read_cell(args: argparse.Namespace) -> Dict[str, Any]:
    return {"success": True, "outcome": "read_cell", **read_cell(Path(args.xlsx_path).resolve(), args.sheet, args.cell)}


def _cmd_patch_cell(args: argparse.Namespace) -> Dict[str, Any]:
    path = Path(args.xlsx_path).resolve()
    if bool(args.target_id) == bool(args.sheet or args.cell):
        raise InvalidValueError("Chi dinh dung 1 trong: --target-id, hoac --sheet + --cell.")
    if not args.target_id and not (args.sheet and args.cell):
        raise InvalidValueError("Can ca --sheet va --cell.")
    op = CellOp(args.type, _typed_value(args.value, args.type), args.sheet, args.cell, args.target_id)
    return _patch_payload(batch_patch(path, [op], args.expected_revision, **_patch_kwargs(args, path)))


def _cmd_patch_range(args: argparse.Namespace) -> Dict[str, Any]:
    path = Path(args.xlsx_path).resolve()
    try:
        values = json.loads(args.values_json)
    except json.JSONDecodeError as exc:
        raise InvalidValueError(f"--values-json khong phai JSON hop le: {exc}") from None
    return _patch_payload(patch_range(path, args.sheet, args.range, values, args.expected_revision, **_patch_kwargs(args, path)))


def _cmd_batch_patch(args: argparse.Namespace) -> Dict[str, Any]:
    path = Path(args.xlsx_path).resolve()
    if bool(args.ops_file) == bool(args.ops_json):
        raise InvalidValueError("Chi dinh dung 1 trong --ops-file, --ops-json.")
    raw = Path(args.ops_file).read_text(encoding="utf-8") if args.ops_file else args.ops_json
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise InvalidValueError(f"Ops khong phai JSON hop le: {exc}") from None
    if isinstance(parsed, dict) and "ops" in parsed:
        parsed = parsed["ops"]
    if not isinstance(parsed, list):
        raise InvalidValueError("Ops phai la mang JSON (hoac object co khoa 'ops').")
    ops = [op_from_dict(item) for item in parsed]
    return _patch_payload(batch_patch(path, ops, args.expected_revision, **_patch_kwargs(args, path)))


def _cmd_backups(args: argparse.Namespace) -> Dict[str, Any]:
    path = Path(args.xlsx_path).resolve()
    return {"success": True, "outcome": "listed", "backups": list_backups(path, work_dir=_work_dir(path, args.work_dir))}


def _cmd_restore(args: argparse.Namespace) -> Dict[str, Any]:
    path = Path(args.xlsx_path).resolve()
    result = restore_backup(
        path, args.backup_id, work_dir=_work_dir(path, args.work_dir),
        allowed_roots=_roots(path), actor=args.actor or DEFAULT_ACTOR,
    )
    return {
        "success": True,
        "outcome": result["outcome"],
        "document_revision": f"sha256:{result['sha256']}",
        "job_id": result["job_id"],
        "backup_path": result.get("backup_path"),
        "restored_from_backup_id": result.get("restored_from_backup_id"),
    }


_COMMANDS = {
    "inspect": _cmd_inspect,
    "read-sheet": _cmd_read_sheet,
    "read-cell": _cmd_read_cell,
    "patch-cell": _cmd_patch_cell,
    "patch-range": _cmd_patch_range,
    "batch-patch": _cmd_batch_patch,
    "backups": _cmd_backups,
    "restore": _cmd_restore,
}


# ---------------------------------------------------------------------------
# argparse wiring
# ---------------------------------------------------------------------------
def _add_write_flags(p: argparse.ArgumentParser) -> None:
    p.add_argument("--expected-revision", required=True)
    p.add_argument("--truncate", action="store_true", help="Cat chuoi > 32767 ky tu thay vi tu choi.")
    p.add_argument("--json", action="store_true")
    p.add_argument("--work-dir")
    p.add_argument("--actor")


def _build_parser() -> argparse.ArgumentParser:
    parser = _JsonAwareArgumentParser(prog="excel-engine", description="Runtime/CLI chuyen dung xu ly Excel (tools/excel-engine/).")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("inspect", help="Quet workbook: sheet, dimension, merged, Table, locator o (chi doc).")
    p.add_argument("xlsx_path")
    p.add_argument("--limit", type=int, default=5000, help="So locator toi da tra ve.")
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("read-sheet", help="Doc cac o cua 1 sheet (chi doc).")
    p.add_argument("xlsx_path")
    p.add_argument("--sheet", required=True)
    p.add_argument("--range")
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("read-cell", help="Doc 1 o (chi doc).")
    p.add_argument("xlsx_path")
    p.add_argument("--sheet", required=True)
    p.add_argument("--cell", required=True)
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("patch-cell", help="Sua 1 o.")
    p.add_argument("xlsx_path")
    p.add_argument("--sheet")
    p.add_argument("--cell")
    p.add_argument("--target-id", help="Locator cell_s{idx}_r{row}_c{col}, thay cho --sheet/--cell.")
    p.add_argument("--value")
    p.add_argument("--type", choices=["text", "number", "bool", "formula", "blank"], default="text")
    _add_write_flags(p)

    p = sub.add_parser("patch-range", help="Ghi ma tran 2D vao 1 vung (1 lo atomic).")
    p.add_argument("xlsx_path")
    p.add_argument("--sheet", required=True)
    p.add_argument("--range", required=True)
    p.add_argument("--values-json", required=True, help='Vd: [["a",1],[true,{"formula":"SUM(B1:B1)"}]]')
    _add_write_flags(p)

    p = sub.add_parser("batch-patch", help="Ap dung nhieu op trong 1 lo atomic.")
    p.add_argument("xlsx_path")
    p.add_argument("--ops-file")
    p.add_argument("--ops-json")
    _add_write_flags(p)

    p = sub.add_parser("backups", help="Liet ke cac ban sao luu.")
    p.add_argument("xlsx_path")
    p.add_argument("--json", action="store_true")
    p.add_argument("--work-dir")

    p = sub.add_parser("restore", help="Phuc hoi tu 1 ban sao luu.")
    p.add_argument("xlsx_path")
    p.add_argument("--backup-id", required=True)
    p.add_argument("--json", action="store_true")
    p.add_argument("--work-dir")
    p.add_argument("--actor")

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
    args = _build_parser().parse_args(_RAW_ARGV)

    try:
        payload = _COMMANDS[args.command](args)
    except Exception as exc:  # noqa: BLE001 - CLI boundary: moi loi deu thanh JSON contract
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
