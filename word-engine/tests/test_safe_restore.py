"""
Tests for core/safety_gateway.py's Safe Restore Mutation Transaction (backup creation,
list_backups, restore_backup) - see docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 3.4
and the module docstring in core/safety_gateway.py for the design (centralized
`.jarvis/backups/<target_key>/`, restore_backup() reusing the ported 18-step commit broker
rather than its own ad-hoc overwrite).

No mocking, following this project's existing test style (dich-thuat/tests/test_docx_guard.py):
real files on disk under tmp_path, real SHA-256 hashes, real JSON on disk. Every test passes an
explicit `work_dir=tmp_path` so `.jarvis/` is rooted inside the test's own tmp_path and never
touches the real repo's `.jarvis/` (see core/safety_gateway.py's module docstring, deviation 1).

Targets use a plain `.txt` extension rather than `.docx` on purpose: Step 9 of the commit
pipeline (DOCX Zip Package Integrity Check) only runs for `.docx` targets, and these tests are
about the generic backup/restore transaction, not DOCX-specific validity - a fixture that needs
to be a real minimal zip would only add unrelated complexity.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from core.safety_gateway import (
    BackupNotFoundError,
    RestoreConflictError,
    issue_job,
    list_backups,
    request_commit,
    restore_backup,
    sha256_bytes,
)


def _commit(work_dir: Path, target_path: Path, content: bytes, actor: str = "tester") -> dict:
    # Mirrors real usage: the destination folder (a dossier/case folder) already exists before
    # Word Engine writes into it. Neither issue_job() nor request_commit() create target_path's
    # parent directory (ported as-is from dich-thuat/engine/file_safety/commit_broker.py, which
    # makes the same assumption) - that is the caller's responsibility.
    target_path.parent.mkdir(parents=True, exist_ok=True)
    ticket = issue_job(target_path, actor=actor, allowed_roots=[str(work_dir)], work_dir=work_dir)
    candidate_path = Path(ticket.work_dir) / f"candidate{target_path.suffix}"
    candidate_path.write_bytes(content)
    return request_commit(ticket.job_id, work_dir=work_dir)


# ---------------------------------------------------------------------------
# Backup creation & list_backups
# ---------------------------------------------------------------------------
def test_no_backup_on_first_write(tmp_path):
    target = tmp_path / "production" / "sample.txt"
    result = _commit(tmp_path, target, b"v1 content")

    assert result["success"] is True
    assert target.read_bytes() == b"v1 content"
    # Nothing existed before the first write - Step 10 has nothing to back up.
    assert list_backups(target, work_dir=tmp_path) == []


def test_backup_created_on_overwrite_and_listed(tmp_path):
    target = tmp_path / "production" / "sample.txt"
    _commit(tmp_path, target, b"v1 content")
    _commit(tmp_path, target, b"v2 content")

    backups = list_backups(target, work_dir=tmp_path)
    assert len(backups) == 1

    entry = backups[0]
    assert entry["target_path"] == str(target.resolve())
    assert entry["sha256"] == sha256_bytes(b"v1 content")
    assert "backup_id" in entry and entry["backup_id"]

    backup_file = Path(entry["backup_file"])
    assert backup_file.exists()
    assert backup_file.read_bytes() == b"v1 content"


def test_list_backups_orders_newest_first(tmp_path):
    target = tmp_path / "production" / "sample.txt"
    _commit(tmp_path, target, b"v1")
    _commit(tmp_path, target, b"v2")  # backs up v1
    _commit(tmp_path, target, b"v3")  # backs up v2

    backups = list_backups(target, work_dir=tmp_path)
    assert len(backups) == 2
    assert backups[0]["sha256"] == sha256_bytes(b"v2")
    assert backups[1]["sha256"] == sha256_bytes(b"v1")


# ---------------------------------------------------------------------------
# restore_backup: success path
# ---------------------------------------------------------------------------
def test_restore_backup_succeeds_and_reverts_content(tmp_path):
    target = tmp_path / "production" / "sample.txt"
    _commit(tmp_path, target, b"v1 content")
    _commit(tmp_path, target, b"v2 content")  # backs up v1, target is now v2 (clean baseline)

    backups_before = list_backups(target, work_dir=tmp_path)
    assert len(backups_before) == 1
    v1_backup_id = backups_before[0]["backup_id"]

    result = restore_backup(
        target, v1_backup_id, work_dir=tmp_path, allowed_roots=[str(tmp_path)], actor="tester"
    )

    assert result["success"] is True
    assert result["outcome"] == "restored_from_backup"
    assert result["restored_from_backup_id"] == v1_backup_id
    assert target.read_bytes() == b"v1 content"

    # The restore itself is a commit (Step 10 backs up the pre-restore v2 state).
    backups_after = list_backups(target, work_dir=tmp_path)
    assert len(backups_after) == 2
    assert any(b["sha256"] == sha256_bytes(b"v2 content") for b in backups_after)


def test_restore_backup_raises_when_backup_id_unknown(tmp_path):
    target = tmp_path / "production" / "sample.txt"
    _commit(tmp_path, target, b"v1 content")
    _commit(tmp_path, target, b"v2 content")

    with pytest.raises(BackupNotFoundError):
        restore_backup(
            target, "does-not-exist", work_dir=tmp_path, allowed_roots=[str(tmp_path)], actor="tester"
        )


# ---------------------------------------------------------------------------
# restore_backup: RestoreConflictError on a disk state the system doesn't recognize
# ---------------------------------------------------------------------------
def test_restore_backup_raises_conflict_on_hand_edit(tmp_path):
    target = tmp_path / "production" / "sample.txt"
    _commit(tmp_path, target, b"v1 content")
    _commit(tmp_path, target, b"v2 content")  # backs up v1, baseline now = v2

    backups = list_backups(target, work_dir=tmp_path)
    v1_backup_id = backups[0]["backup_id"]

    # Simulate a phong ho so staff hand-edit in Word/Notepad: mutate the file directly on disk,
    # bypassing the broker entirely, so it no longer matches the last recorded baseline.
    target.write_bytes(b"hand-edited content, never went through the broker")

    with pytest.raises(RestoreConflictError):
        restore_backup(
            target, v1_backup_id, work_dir=tmp_path, allowed_roots=[str(tmp_path)], actor="tester"
        )

    # Refusal must be total: the hand-edited content on disk is left exactly as found.
    assert target.read_bytes() == b"hand-edited content, never went through the broker"


def test_restore_backup_raises_conflict_when_target_never_tracked(tmp_path):
    # A file that exists on disk but was never written through issue_job/request_commit has no
    # baseline manifest at all ("untracked") - safety_gateway cannot prove it wasn't hand-authored,
    # so restore must refuse exactly like the "modified" case rather than assume it's safe.
    target = tmp_path / "production" / "untracked.txt"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(b"never touched the broker")

    with pytest.raises(RestoreConflictError):
        restore_backup(
            target, "irrelevant-backup-id", work_dir=tmp_path, allowed_roots=[str(tmp_path)], actor="tester"
        )
