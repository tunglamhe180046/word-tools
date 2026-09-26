"""
Tests for core/stamp_ops.py (Phase 3 - Quan tri Con dau & Chu ky) va cli.py's `stamp-ops`
subcommand.

Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md muc 3.6/muc 4 ("Phase 3") va
core/stamp_ops.py's module docstring cho thiet ke. No mocking, real files on disk under tmp_path,
real SHA-256/XML/zip comparisons - same style as tests/test_geometry.py va tests/test_cli.py.
Asset anh dung trong test la byte gia (khong phai PNG that) - apply_stamp() khong doc pixel-data
cua anh (xem module docstring "Gioi han Phase 3"), chi doc extension + noi dung bytes de tinh
SHA-256 va nhung thang vao goi OPC, nen mot file .png voi noi dung bat ky la du de kiem thu day
du hanh vi cua module. Moi test truyen work_dir=tmp_path tuong minh de `.jarvis/` khong bao gio
dung vao `.jarvis/` that cua repo.
"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from hashlib import sha256
from pathlib import Path

import pytest
from docx import Document

from core.inspector import ObjectLocator, inspect_document, load_document_root
from core.stamp_ops import (
    WP_NS,
    UnauthorizedStampAssetError,
    apply_stamp,
    is_asset_authorized,
    register_stamp_asset,
)

CLI_PATH = Path(__file__).resolve().parent.parent / "cli.py"
COMMON_KWARGS = dict(actor="tester", job_id="job-stamp")


# ---------------------------------------------------------------------------
# Fixtures & small helpers
# ---------------------------------------------------------------------------
def _build_sample_docx(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("Nguoi ky: Nguyen Van A")
    doc.add_paragraph("Cho dong dau tai day.")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def _write_stamp_asset(path: Path, content: bytes = b"fake-png-bytes-for-test") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)


def _target_paragraph_locator(docx_path: Path, tmp_path: Path, object_id: str = "para_2") -> ObjectLocator:
    report = inspect_document(docx_path, work_dir=tmp_path, job_id="job-inspect")
    locator_dict = next(loc for loc in report["locators"] if loc["object_id"] == object_id)
    return ObjectLocator(**locator_dict)


def _run_cli(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(CLI_PATH), *args], capture_output=True, text=True)


# ---------------------------------------------------------------------------
# Chen anh hop le (co trong allowlist)
# ---------------------------------------------------------------------------
def test_apply_stamp_inserts_image_when_asset_authorized(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)

    locator = _target_paragraph_locator(docx_path, tmp_path)

    result = apply_stamp(
        docx_path, asset_path, locator, width_mm=30.0, height_mm=30.0,
        work_dir=tmp_path, allowed_roots=[str(tmp_path)], **COMMON_KWARGS,
    )
    assert result.success is True
    assert result.outcome == "committed_clean"

    # File van mo lai binh thuong bang python-docx sau khi chen.
    doc = Document(str(docx_path))
    assert len(doc.paragraphs) == 2

    with zipfile.ZipFile(docx_path, "r") as zf:
        names = zf.namelist()
        media_files = [n for n in names if n.startswith("word/media/")]
        assert len(media_files) == 1
        assert zf.read(media_files[0]) == asset_path.read_bytes()

        rels_xml = zf.read("word/_rels/document.xml.rels")
        assert b"relationships/image" in rels_xml
        assert media_files[0].rsplit("/", 1)[-1].encode() in rels_xml

        content_types_xml = zf.read("[Content_Types].xml")
        assert b'Extension="png"' in content_types_xml

    root = load_document_root(docx_path)
    drawings = root.findall(f".//{{{WP_NS}}}inline")
    assert len(drawings) == 1


def test_apply_stamp_leaves_untouched_opc_parts_raw_identical_except_the_four_touched(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)
    locator = _target_paragraph_locator(docx_path, tmp_path)

    with zipfile.ZipFile(docx_path, "r") as zf:
        before_names = set(zf.namelist())

    apply_stamp(
        docx_path, asset_path, locator, width_mm=20.0, height_mm=20.0,
        work_dir=tmp_path, allowed_roots=[str(tmp_path)], **COMMON_KWARGS,
    )

    with zipfile.ZipFile(docx_path, "r") as zf:
        after_names = set(zf.namelist())

    new_names = after_names - before_names
    assert len(new_names) == 1
    assert next(iter(new_names)).startswith("word/media/")


# ---------------------------------------------------------------------------
# Tu choi anh chua duoc dang ky (khong co trong allowlist)
# ---------------------------------------------------------------------------
def test_apply_stamp_rejects_unregistered_asset(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    # KHONG goi register_stamp_asset() - asset chua duoc dang ky.
    locator = _target_paragraph_locator(docx_path, tmp_path)
    before_bytes = docx_path.read_bytes()

    with pytest.raises(UnauthorizedStampAssetError):
        apply_stamp(
            docx_path, asset_path, locator, width_mm=30.0, height_mm=30.0,
            work_dir=tmp_path, allowed_roots=[str(tmp_path)], **COMMON_KWARGS,
        )

    assert docx_path.read_bytes() == before_bytes
    assert is_asset_authorized(sha256(asset_path.read_bytes()).hexdigest(), work_dir=tmp_path) is False


def test_apply_stamp_allow_unregistered_flag_bypasses_whitelist(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    locator = _target_paragraph_locator(docx_path, tmp_path)

    result = apply_stamp(
        docx_path, asset_path, locator, width_mm=30.0, height_mm=30.0,
        work_dir=tmp_path, allowed_roots=[str(tmp_path)], allow_unregistered=True, **COMMON_KWARGS,
    )
    assert result.success is True


def test_apply_stamp_rejects_locator_kind_table_cell(tmp_path):
    docx_path = tmp_path / "sample.docx"
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "O 1"
    doc.save(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)

    report = inspect_document(docx_path, work_dir=tmp_path, job_id="job-inspect")
    locator = ObjectLocator(**next(loc for loc in report["locators"] if loc["object_id"] == "cell_t1_r1_c1"))

    with pytest.raises(ValueError):
        apply_stamp(
            docx_path, asset_path, locator, width_mm=30.0, height_mm=30.0,
            work_dir=tmp_path, allowed_roots=[str(tmp_path)], **COMMON_KWARGS,
        )


# ---------------------------------------------------------------------------
# So kiem toan xuat xu (Provenance Ledger)
# ---------------------------------------------------------------------------
def test_apply_stamp_writes_provenance_ledger_entry(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    asset_sha256 = register_stamp_asset(asset_path, work_dir=tmp_path, registered_by="tester")
    locator = _target_paragraph_locator(docx_path, tmp_path)

    result = apply_stamp(
        docx_path, asset_path, locator, width_mm=25.0, height_mm=18.0,
        work_dir=tmp_path, allowed_roots=[str(tmp_path)], actor="signer-1", job_id="job-prov",
    )

    ledger_path = tmp_path / ".jarvis" / "audit" / "stamp_provenance.jsonl"
    assert ledger_path.exists()
    lines = ledger_path.read_text(encoding="utf-8").strip().splitlines()
    assert len(lines) == 1
    entry = json.loads(lines[0])
    assert entry["document_path"] == str(docx_path)
    assert entry["asset_sha256"] == asset_sha256
    assert entry["authorized_by"] == "signer-1"
    assert entry["inserted_dimensions"] == {"width_mm": 25.0, "height_mm": 18.0}
    assert entry["job_id"] == result.job_id
    assert entry["timestamp"]


def test_apply_stamp_appends_multiple_entries_across_calls(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)

    locator1 = _target_paragraph_locator(docx_path, tmp_path)
    apply_stamp(
        docx_path, asset_path, locator1, width_mm=10.0, height_mm=10.0,
        work_dir=tmp_path, allowed_roots=[str(tmp_path)], actor="signer-1", job_id="job-1",
    )
    locator2 = _target_paragraph_locator(docx_path, tmp_path)
    apply_stamp(
        docx_path, asset_path, locator2, width_mm=10.0, height_mm=10.0,
        work_dir=tmp_path, allowed_roots=[str(tmp_path)], actor="signer-2", job_id="job-2",
    )

    ledger_path = tmp_path / ".jarvis" / "audit" / "stamp_provenance.jsonl"
    lines = ledger_path.read_text(encoding="utf-8").strip().splitlines()
    assert len(lines) == 2


# ---------------------------------------------------------------------------
# Kich thuoc EMU chuan (width_mm * 36000, height_mm * 36000)
# ---------------------------------------------------------------------------
def test_apply_stamp_emu_dimensions_match_mm_conversion(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)
    locator = _target_paragraph_locator(docx_path, tmp_path)

    apply_stamp(
        docx_path, asset_path, locator, width_mm=40.0, height_mm=15.0,
        work_dir=tmp_path, allowed_roots=[str(tmp_path)], **COMMON_KWARGS,
    )

    root = load_document_root(docx_path)
    extents = root.findall(f".//{{{WP_NS}}}extent")
    assert len(extents) == 1
    assert extents[0].get("cx") == str(round(40.0 * 36000))
    assert extents[0].get("cy") == str(round(15.0 * 36000))

    xfrm_exts = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
    assert len(xfrm_exts) == 1
    assert xfrm_exts[0].get("cx") == str(round(40.0 * 36000))
    assert xfrm_exts[0].get("cy") == str(round(15.0 * 36000))


def test_apply_stamp_rejects_non_positive_dimensions(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)
    locator = _target_paragraph_locator(docx_path, tmp_path)

    with pytest.raises(ValueError):
        apply_stamp(
            docx_path, asset_path, locator, width_mm=0, height_mm=10.0,
            work_dir=tmp_path, allowed_roots=[str(tmp_path)], **COMMON_KWARGS,
        )


def test_apply_stamp_rejects_unsupported_extension(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.bmp"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)
    locator = _target_paragraph_locator(docx_path, tmp_path)

    with pytest.raises(ValueError):
        apply_stamp(
            docx_path, asset_path, locator, width_mm=10.0, height_mm=10.0,
            work_dir=tmp_path, allowed_roots=[str(tmp_path)], **COMMON_KWARGS,
        )


# ---------------------------------------------------------------------------
# CLI: stamp-ops --json contract
# ---------------------------------------------------------------------------
def test_cli_stamp_ops_json_contract(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)

    inspect_proc = _run_cli("inspect", str(docx_path), "--json", "--work-dir", str(tmp_path))
    assert inspect_proc.returncode == 0, inspect_proc.stderr
    revision = json.loads(inspect_proc.stdout)["document_revision"]

    proc = _run_cli(
        "stamp-ops", str(docx_path),
        "--asset-path", str(asset_path),
        "--target-id", "para_2",
        "--expected-revision", revision,
        "--width-mm", "30", "--height-mm", "30",
        "--json", "--work-dir", str(tmp_path),
    )

    assert proc.returncode == 0, proc.stderr
    assert proc.stdout.count("\n") == 1
    payload = json.loads(proc.stdout)
    assert payload["success"] is True
    assert payload["outcome"] == "committed_clean"
    assert payload["document_revision"].startswith("sha256:")

    with zipfile.ZipFile(docx_path, "r") as zf:
        media_files = [n for n in zf.namelist() if n.startswith("word/media/")]
        assert len(media_files) == 1


def test_cli_stamp_ops_unauthorized_asset_fails_cleanly(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    # KHONG dang ky asset.

    inspect_proc = _run_cli("inspect", str(docx_path), "--json", "--work-dir", str(tmp_path))
    revision = json.loads(inspect_proc.stdout)["document_revision"]

    proc = _run_cli(
        "stamp-ops", str(docx_path),
        "--asset-path", str(asset_path),
        "--target-id", "para_2",
        "--expected-revision", revision,
        "--json", "--work-dir", str(tmp_path),
    )

    assert proc.returncode == 1
    payload = json.loads(proc.stdout)
    assert payload["success"] is False
    assert payload["error"] == "UnauthorizedStampAssetError"
    assert "UNAUTHORIZED_STAMP_ASSET" in payload["reason"]


def test_cli_stamp_ops_allow_unregistered_flag(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)

    inspect_proc = _run_cli("inspect", str(docx_path), "--json", "--work-dir", str(tmp_path))
    revision = json.loads(inspect_proc.stdout)["document_revision"]

    proc = _run_cli(
        "stamp-ops", str(docx_path),
        "--asset-path", str(asset_path),
        "--target-id", "para_2",
        "--expected-revision", revision,
        "--allow-unregistered",
        "--json", "--work-dir", str(tmp_path),
    )

    assert proc.returncode == 0, proc.stderr
    payload = json.loads(proc.stdout)
    assert payload["success"] is True


def test_cli_stamp_ops_wrong_expected_revision_fails_cleanly_and_leaves_disk_untouched(tmp_path):
    docx_path = tmp_path / "sample.docx"
    _build_sample_docx(docx_path)
    asset_path = tmp_path / "stamp.png"
    _write_stamp_asset(asset_path)
    register_stamp_asset(asset_path, work_dir=tmp_path)
    before_bytes = docx_path.read_bytes()

    proc = _run_cli(
        "stamp-ops", str(docx_path),
        "--asset-path", str(asset_path),
        "--target-id", "para_2",
        "--expected-revision", "sha256:0000000000000000000000000000000000000000000000000000000000000000",
        "--json", "--work-dir", str(tmp_path),
    )

    assert proc.returncode == 1
    payload = json.loads(proc.stdout)
    assert payload["success"] is False
    assert payload["error"] == "DocumentDriftError"
    assert docx_path.read_bytes() == before_bytes
