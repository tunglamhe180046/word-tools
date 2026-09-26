"""Chung minh moi thao tac chi ghi vao tmp_path: khong dung vao cay ma nguon, `.jarvis/` that cua repo
hay thu muc cwd goc; khong de lai lock/tmp/job tam."""
import os
from pathlib import Path

from conftest import WORKSPACE_ROOT, file_revision

from core.inspector import inspect_workbook, read_cell
from core.safety_gateway import list_backups, restore_backup
from core.surgical_patcher import CellOp, batch_patch

REPO_ROOT = WORKSPACE_ROOT.parent.parent
_SKIP_DIRS = {"__pycache__", ".pytest_cache"}


def _snapshot(root: Path):
    snap = {}
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in _SKIP_DIRS]
        for name in filenames:
            p = Path(dirpath) / name
            st = p.stat()
            snap[str(p)] = (st.st_size, st.st_mtime_ns)
    return snap


def test_operations_write_only_inside_tmp_path(xlsx_path, tmp_path):
    engine_before = _snapshot(WORKSPACE_ROOT)
    repo_jarvis = REPO_ROOT / ".jarvis"
    jarvis_before = _snapshot(repo_jarvis) if repo_jarvis.exists() else None
    original_cwd = Path.cwd()
    assert original_cwd == tmp_path.resolve() or original_cwd == tmp_path  # conftest da chdir

    inspect_workbook(xlsx_path)
    read_cell(xlsx_path, "Data", "A1")
    batch_patch(xlsx_path, [CellOp("text", "iso", "Data", "A2"), CellOp("formula", "1+1", "Data", "F1")], file_revision(xlsx_path))
    restore_backup(xlsx_path, list_backups(xlsx_path)[0]["backup_id"], work_dir=tmp_path, allowed_roots=[str(tmp_path)])

    assert _snapshot(WORKSPACE_ROOT) == engine_before  # cay tools/excel-engine khong doi
    assert (_snapshot(repo_jarvis) if repo_jarvis.exists() else None) == jarvis_before
    assert (tmp_path / ".jarvis").is_dir()  # .jarvis nam trong tmp_path


def test_no_temp_debris_left_after_patch_and_restore(xlsx_path, tmp_path):
    batch_patch(xlsx_path, [CellOp("text", "x", "Data", "A2")], file_revision(xlsx_path))
    restore_backup(xlsx_path, list_backups(xlsx_path)[0]["backup_id"], work_dir=tmp_path, allowed_roots=[str(tmp_path)])
    names = [p.name for p in tmp_path.rglob("*")]
    assert not [n for n in names if n.endswith(".lock") or ".tmp-" in n or n.startswith("staged") or n.startswith("candidate")]
    assert list((tmp_path / ".jarvis" / "work").glob("*")) == []  # thu muc job tam da duoc don
    top = {p.name for p in tmp_path.iterdir()}
    assert top == {"book.xlsx", ".jarvis"}


def test_failed_patch_leaves_no_job_dirs_or_backups(xlsx_path, tmp_path):
    try:
        batch_patch(xlsx_path, [CellOp("text", "x", "Data", "B6")], file_revision(xlsx_path))
    except Exception:
        pass
    assert {p.name for p in tmp_path.iterdir()} == {"book.xlsx"}  # loi som -> khong tao gi ca
