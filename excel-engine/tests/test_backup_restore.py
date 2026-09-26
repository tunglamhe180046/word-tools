import hashlib
import os
import time
from pathlib import Path

import pytest
from conftest import file_revision

from core.inspector import read_cell
from core.safety_gateway import BackupNotFoundError, RestoreConflictError, list_backups, restore_backup
from core.surgical_patcher import patch_cell


def test_no_backups_before_any_patch(xlsx_path):
    assert list_backups(xlsx_path) == []


def test_patch_creates_backup_of_previous_bytes(xlsx_path):
    original = xlsx_path.read_bytes()
    result = patch_cell(xlsx_path, "Data", "A2", "changed", file_revision(xlsx_path))
    backups = list_backups(xlsx_path)
    assert len(backups) == 1
    entry = backups[0]
    assert entry["sha256"] == hashlib.sha256(original).hexdigest()
    assert Path(entry["backup_file"]).read_bytes() == original
    assert result.backup_path == entry["backup_file"]
    assert ".jarvis" in Path(entry["backup_file"]).parts  # tap trung, khong nam canh file goc


def test_restore_returns_exact_original_and_backs_up_current(xlsx_path):
    original = xlsx_path.read_bytes()
    patch_cell(xlsx_path, "Data", "A2", "changed", file_revision(xlsx_path))
    patched = xlsx_path.read_bytes()
    backup_id = list_backups(xlsx_path)[0]["backup_id"]

    result = restore_backup(xlsx_path, backup_id, work_dir=xlsx_path.parent, allowed_roots=[str(xlsx_path.parent)])

    assert result["outcome"] == "restored_from_backup" and result["restored_from_backup_id"] == backup_id
    assert xlsx_path.read_bytes() == original
    assert read_cell(xlsx_path, "Data", "A2")["cell"]["value"] == "Alice"
    backups = list_backups(xlsx_path)
    assert len(backups) == 2  # them 1 ban sao luu cua trang thai da patch
    assert any(Path(b["backup_file"]).read_bytes() == patched for b in backups)


def test_restore_refuses_when_user_edited_by_hand(xlsx_path):
    patch_cell(xlsx_path, "Data", "A2", "changed", file_revision(xlsx_path))
    backup_id = list_backups(xlsx_path)[0]["backup_id"]
    manual = xlsx_path.read_bytes() + b"\x00manual-edit"  # nguoi dung sua tay ngoai engine
    xlsx_path.write_bytes(manual)
    with pytest.raises(RestoreConflictError):
        restore_backup(xlsx_path, backup_id, work_dir=xlsx_path.parent, allowed_roots=[str(xlsx_path.parent)])
    assert xlsx_path.read_bytes() == manual


def test_restore_unknown_backup_id(xlsx_path):
    patch_cell(xlsx_path, "Data", "A2", "changed", file_revision(xlsx_path))
    with pytest.raises(BackupNotFoundError):
        restore_backup(xlsx_path, "does-not-exist", work_dir=xlsx_path.parent, allowed_roots=[str(xlsx_path.parent)])


def test_restore_untracked_file_is_refused(xlsx_path):
    # file chua tung di qua engine -> chua co baseline -> khong the chung minh khong co sua tay
    with pytest.raises(RestoreConflictError):
        restore_backup(xlsx_path, "x", work_dir=xlsx_path.parent, allowed_roots=[str(xlsx_path.parent)])


def test_backups_sorted_newest_first(xlsx_path):
    patch_cell(xlsx_path, "Data", "A2", "one", file_revision(xlsx_path))
    time.sleep(0.01)
    patch_cell(xlsx_path, "Data", "A2", "two", file_revision(xlsx_path))
    backups = list_backups(xlsx_path)
    assert len(backups) == 2
    assert backups[0]["created_at"] >= backups[1]["created_at"]
    assert Path(backups[1]["backup_file"]).read_bytes() != Path(backups[0]["backup_file"]).read_bytes()


def test_backups_are_keyed_by_canonical_path(xlsx_path, tmp_path):
    patch_cell(xlsx_path, "Data", "A2", "one", file_revision(xlsx_path))
    dotted = tmp_path / "sub" / ".." / "book.xlsx"  # alias bang dot-segment, chay tren moi HDH
    (tmp_path / "sub").mkdir()
    assert len(list_backups(dotted)) == 1
    if os.name == "nt":  # NTFS khong phan biet hoa/thuong
        assert len(list_backups(Path(str(xlsx_path).upper()))) == 1


def test_restore_refuses_corrupted_backup(xlsx_path):
    from core.safety_gateway import BackupCorruptedError

    patch_cell(xlsx_path, "Data", "A2", "changed", file_revision(xlsx_path))
    entry = list_backups(xlsx_path)[0]
    Path(entry["backup_file"]).write_bytes(Path(entry["backup_file"]).read_bytes() + b"tamper")
    current = xlsx_path.read_bytes()
    with pytest.raises(BackupCorruptedError):
        restore_backup(xlsx_path, entry["backup_id"], work_dir=xlsx_path.parent, allowed_roots=[str(xlsx_path.parent)])
    assert xlsx_path.read_bytes() == current
