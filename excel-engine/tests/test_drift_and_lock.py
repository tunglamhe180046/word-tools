import json
import zipfile

import pytest
from conftest import build_minimal_xlsx, file_revision

import core.surgical_patcher as sp
from core.inspector import read_cell
from core.safety_gateway import (
    DocumentDriftError,
    DocumentLockedError,
    DocumentModifiedError,
    JobAlreadyCommittedError,
    JobConcurrentCommitError,
    JobTicketTamperedError,
    assert_revision,
    compute_document_revision,
    issue_job,
    request_commit,
)
from core.surgical_patcher import patch_cell


def _no_debris(tmp_path):
    stray = [p.name for p in tmp_path.rglob("*") if p.name.endswith((".write.lock", ".commit.lock")) or ".tmp-" in p.name]
    assert stray == []


def test_wrong_expected_revision_is_drift_and_writes_nothing(xlsx_path):
    original = xlsx_path.read_bytes()
    with pytest.raises(DocumentDriftError) as exc:
        patch_cell(xlsx_path, "Data", "A2", "x", "sha256:" + "0" * 64)
    assert exc.value.actual_revision == file_revision(xlsx_path)
    assert xlsx_path.read_bytes() == original


def test_stale_revision_after_previous_patch_is_drift(xlsx_path):
    old_rev = file_revision(xlsx_path)
    patch_cell(xlsx_path, "Data", "A2", "first", old_rev)
    with pytest.raises(DocumentDriftError):
        patch_cell(xlsx_path, "Data", "A3", "second", old_rev)
    assert read_cell(xlsx_path, "Data", "A3")["cell"]["value"] == "Bob"


def test_revision_helpers():
    assert compute_document_revision(b"abc").startswith("sha256:ba7816bf")
    assert assert_revision(b"abc", compute_document_revision(b"abc")) == compute_document_revision(b"abc")
    with pytest.raises(DocumentDriftError):
        assert_revision(b"abc", "sha256:00")


def test_manual_edit_between_read_and_issue_job_is_caught(xlsx_path, monkeypatch, tmp_path):
    """Nguoi dung luu file ngay sau khi engine da doc, truoc luc issue_job: baseline cua ticket khac."""
    rev = file_revision(xlsx_path)
    real_issue = sp.issue_job

    def racing_issue(*args, **kwargs):
        build_minimal_xlsx(xlsx_path, merge=False)  # "sua tay": doi noi dung file
        return real_issue(*args, **kwargs)

    monkeypatch.setattr(sp, "issue_job", racing_issue)
    with pytest.raises(DocumentDriftError):
        patch_cell(xlsx_path, "Data", "A2", "lost?", rev)
    edited = xlsx_path.read_bytes()
    assert b'<mergeCell ' not in zipfile.ZipFile(xlsx_path).read("xl/worksheets/sheet1.xml")
    assert b"lost?" not in edited
    _no_debris(tmp_path)
    assert not list((tmp_path / ".jarvis" / "work").glob("*"))  # job tam da duoc don


def test_manual_edit_between_issue_job_and_commit_hits_optimistic_lock(xlsx_path, monkeypatch, tmp_path):
    """Buoc 8 cua Broker: file doi sau luc issue_job -> DocumentModifiedError, khong ghi de."""
    rev = file_revision(xlsx_path)
    real_commit = sp.request_commit

    def racing_commit(job_id, work_dir=None):
        build_minimal_xlsx(xlsx_path, table=False)
        return real_commit(job_id, work_dir=work_dir)

    monkeypatch.setattr(sp, "request_commit", racing_commit)
    with pytest.raises(DocumentModifiedError):
        patch_cell(xlsx_path, "Data", "A2", "lost?", rev)
    assert b"lost?" not in xlsx_path.read_bytes()
    assert "xl/tables/table1.xml" not in zipfile.ZipFile(xlsx_path).namelist()  # ban sua tay con nguyen
    _no_debris(tmp_path)


def test_excel_owner_file_blocks_commit(xlsx_path, tmp_path):
    (tmp_path / "~$book.xlsx").write_bytes(b"lock")
    original = xlsx_path.read_bytes()
    with pytest.raises(DocumentLockedError):
        patch_cell(xlsx_path, "Data", "A2", "x", file_revision(xlsx_path))
    assert xlsx_path.read_bytes() == original
    _no_debris(tmp_path)


def test_excel_owner_file_variant_without_first_two_chars(xlsx_path, tmp_path):
    (tmp_path / "~$ok.xlsx").write_bytes(b"lock")  # ~$ + book.xlsx[2:]
    with pytest.raises(DocumentLockedError):
        patch_cell(xlsx_path, "Data", "A2", "x", file_revision(xlsx_path))


def test_target_outside_allowed_roots_is_refused(xlsx_path, tmp_path):
    other = tmp_path / "elsewhere"
    other.mkdir()
    original = xlsx_path.read_bytes()
    with pytest.raises(ValueError, match="allowed_roots"):
        patch_cell(xlsx_path, "Data", "A2", "x", file_revision(xlsx_path), allowed_roots=[str(other)])
    assert xlsx_path.read_bytes() == original


def _ticket_with_candidate(xlsx_path, tmp_path):
    ticket = issue_job(xlsx_path, actor="t", allowed_roots=[str(tmp_path)], work_dir=tmp_path)
    (tmp_path / ".jarvis" / "work" / ticket.job_id / "candidate.xlsx").write_bytes(build_minimal_xlsx(tmp_path / "cand.xlsx").read_bytes())
    return ticket


def test_broker_concurrent_commit_lease_lock(xlsx_path, tmp_path):
    ticket = _ticket_with_candidate(xlsx_path, tmp_path)
    (tmp_path / ".jarvis" / "jobs" / f"{ticket.job_id}.commit.lock").write_text("held")
    with pytest.raises(JobConcurrentCommitError):
        request_commit(ticket.job_id, work_dir=tmp_path)


def test_broker_rejects_tampered_ticket(xlsx_path, tmp_path):
    ticket = _ticket_with_candidate(xlsx_path, tmp_path)
    ticket_file = tmp_path / ".jarvis" / "jobs" / f"{ticket.job_id}.json"
    data = json.loads(ticket_file.read_text(encoding="utf-8"))
    data["target_path"] = str(tmp_path / "other.xlsx")
    ticket_file.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(JobTicketTamperedError):
        request_commit(ticket.job_id, work_dir=tmp_path)


def test_broker_ticket_is_single_use(xlsx_path, tmp_path):
    ticket = _ticket_with_candidate(xlsx_path, tmp_path)
    request_commit(ticket.job_id, work_dir=tmp_path)
    with pytest.raises(JobAlreadyCommittedError):
        request_commit(ticket.job_id, work_dir=tmp_path)


def test_broker_rejects_corrupt_xlsx_candidate(xlsx_path, tmp_path):
    original = xlsx_path.read_bytes()
    ticket = issue_job(xlsx_path, actor="t", allowed_roots=[str(tmp_path)], work_dir=tmp_path)
    (tmp_path / ".jarvis" / "work" / ticket.job_id / "candidate.xlsx").write_bytes(b"PK\x03\x04 garbage")
    with pytest.raises(ValueError, match="XLSX package integrity"):
        request_commit(ticket.job_id, work_dir=tmp_path)
    assert xlsx_path.read_bytes() == original
    _no_debris(tmp_path)
