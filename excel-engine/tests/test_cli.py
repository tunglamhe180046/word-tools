import json
import os
import subprocess
import sys

from conftest import WORKSPACE_ROOT, build_minimal_xlsx, file_revision

import cli


def run(capsys, *argv):
    try:
        code = cli.main(list(argv))
    except SystemExit as exc:  # loi cu phap argparse: exit 1 sau khi in dong JSON
        code = exc.code
    out = capsys.readouterr().out
    lines = [ln for ln in out.splitlines() if ln.strip()]
    assert len(lines) == 1, out  # hop dong: DUNG 1 dong JSON tren stdout
    return code, json.loads(lines[0])


def test_inspect_json(capsys, xlsx_path):
    code, out = run(capsys, "inspect", str(xlsx_path), "--json")
    assert code == 0 and out["success"] is True and out["outcome"] == "inspected"
    assert out["document_revision"] == file_revision(xlsx_path)
    assert [s["name"] for s in out["sheets"]] == ["Data", "Second"]


def test_read_sheet_and_read_cell_json(capsys, xlsx_path):
    code, out = run(capsys, "read-sheet", str(xlsx_path), "--sheet", "Data", "--range", "A1:B2", "--json")
    assert code == 0 and sorted(c["ref"] for c in out["cells"]) == ["A1", "A2", "B1", "B2"]
    code, out = run(capsys, "read-cell", str(xlsx_path), "--sheet", "Data", "--cell", "B4", "--json")
    assert code == 0 and out["cell"]["formula"] == "SUM(B2:B3)" and out["cell"]["cached_value"] == 12


def test_patch_cell_flow(capsys, xlsx_path):
    rev = file_revision(xlsx_path)
    code, out = run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--value", "Xin chào",
                    "--expected-revision", rev, "--json")
    assert code == 0 and out["outcome"] == "patched" and out["document_revision"] == file_revision(xlsx_path)
    assert out["backup_path"] and out["ops_applied"] == 1
    code, out = run(capsys, "read-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--json")
    assert out["cell"]["value"] == "Xin chào"


def test_patch_cell_types_and_locator(capsys, xlsx_path):
    def go(*extra):
        code, out = run(capsys, "patch-cell", str(xlsx_path), *extra, "--expected-revision", file_revision(xlsx_path), "--json")
        assert code == 0, out
    go("--sheet", "Data", "--cell", "B3", "--value", "-12.5", "--type", "number")
    go("--sheet", "Data", "--cell", "B2", "--value", "true", "--type", "bool")
    go("--sheet", "Data", "--cell", "C3", "--value", "=B2", "--type", "formula")
    go("--target-id", "cell_s1_r2_c1", "--type", "blank")
    cells = {c["ref"]: c for c in run(capsys, "read-sheet", str(xlsx_path), "--sheet", "Data", "--json")[1]["cells"]}
    assert cells["B3"]["value"] == -12.5 and cells["B2"]["value"] is True
    assert cells["C3"]["formula"] == "B2" and cells["A2"]["type"] == "blank"


def test_patch_requires_expected_revision(capsys, xlsx_path):
    code, out = run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--value", "x", "--json")
    assert code == 1 and out["success"] is False and out["error"] == "ArgumentError"
    assert "expected-revision" in out["reason"]
    code, out = run(capsys, "batch-patch", str(xlsx_path), "--ops-json", "[]", "--json")
    assert code == 1 and out["error"] == "ArgumentError"
    code, out = run(capsys, "patch-range", str(xlsx_path), "--sheet", "Data", "--range", "A1", "--values-json", "[[1]]", "--json")
    assert code == 1 and out["error"] == "ArgumentError"


def test_drift_error_contract(capsys, xlsx_path):
    original = xlsx_path.read_bytes()
    code, out = run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--value", "x",
                    "--expected-revision", "sha256:deadbeef", "--json")
    assert code == 1 and out == {"success": False, "error": "DocumentDriftError", "reason": out["reason"]}
    assert "DOCUMENT_DRIFT" in out["reason"] and xlsx_path.read_bytes() == original


def test_fail_closed_errors_via_cli(capsys, xlsx_path):
    original = xlsx_path.read_bytes()
    rev = file_revision(xlsx_path)
    for extra, err in ((["--sheet", "Data", "--cell", "B6"], "MergedCellError"),
                       (["--sheet", "Data", "--cell", "A1"], "TableHeaderError"),
                       (["--sheet", "Data", "--cell", "D8"], "ArrayFormulaError"),
                       (["--sheet", "Nope", "--cell", "A1"], "SheetNotFoundError")):
        code, out = run(capsys, "patch-cell", str(xlsx_path), *extra, "--value", "x", "--expected-revision", rev, "--json")
        assert code == 1 and out["error"] == err, out
    assert xlsx_path.read_bytes() == original


def test_patch_cell_argument_validation(capsys, xlsx_path):
    rev = file_revision(xlsx_path)
    for extra in (["--sheet", "Data"], ["--target-id", "cell_s1_r1_c1", "--sheet", "Data", "--cell", "A2"], []):
        code, out = run(capsys, "patch-cell", str(xlsx_path), *extra, "--value", "x", "--expected-revision", rev, "--json")
        assert code == 1 and out["error"] == "InvalidValueError"
    code, out = run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--value", "abc", "--type", "number",
                    "--expected-revision", rev, "--json")
    assert code == 1 and out["error"] == "InvalidValueError"
    code, out = run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--type", "text",
                    "--expected-revision", rev, "--json")
    assert code == 1 and out["error"] == "InvalidValueError"


def test_patch_range_and_batch_patch(capsys, xlsx_path, tmp_path):
    code, out = run(capsys, "patch-range", str(xlsx_path), "--sheet", "Data", "--range", "A10:B11",
                    "--values-json", '[["a",1],["b",2]]', "--expected-revision", file_revision(xlsx_path), "--json")
    assert code == 0 and out["ops_applied"] == 4

    ops = [{"sheet": "Data", "cell": "A2", "value": "batch"}, {"sheet": "Second", "cell": "C1", "formula": "=A1*2"}]
    code, out = run(capsys, "batch-patch", str(xlsx_path), "--ops-json", json.dumps(ops), "--expected-revision",
                    file_revision(xlsx_path), "--json")
    assert code == 0 and out["ops_applied"] == 2

    ops_file = tmp_path / "ops.json"
    ops_file.write_text(json.dumps({"ops": [{"sheet": "Data", "cell": "A3", "value": 9}]}), encoding="utf-8")
    code, out = run(capsys, "batch-patch", str(xlsx_path), "--ops-file", str(ops_file), "--expected-revision",
                    file_revision(xlsx_path), "--json")
    assert code == 0 and out["ops_applied"] == 1


def test_batch_patch_bad_inputs(capsys, xlsx_path):
    rev = file_revision(xlsx_path)
    for extra in (["--ops-json", "{not json"], ["--ops-json", '{"a":1}'], [], ["--ops-json", "[1]"]):
        code, out = run(capsys, "batch-patch", str(xlsx_path), *extra, "--expected-revision", rev, "--json")
        assert code == 1 and out["error"] == "InvalidValueError", (extra, out)
    code, out = run(capsys, "patch-range", str(xlsx_path), "--sheet", "Data", "--range", "A1", "--values-json", "nope",
                    "--expected-revision", rev, "--json")
    assert code == 1 and out["error"] == "InvalidValueError"


def test_backups_and_restore(capsys, xlsx_path):
    original = xlsx_path.read_bytes()
    run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--value", "z",
        "--expected-revision", file_revision(xlsx_path), "--json")
    code, out = run(capsys, "backups", str(xlsx_path), "--json")
    assert code == 0 and len(out["backups"]) == 1
    code, out = run(capsys, "restore", str(xlsx_path), "--backup-id", out["backups"][0]["backup_id"], "--json")
    assert code == 0 and out["outcome"] == "restored_from_backup"
    assert xlsx_path.read_bytes() == original and out["document_revision"] == file_revision(xlsx_path)


def test_unsupported_and_missing_files(capsys, tmp_path):
    xls = tmp_path / "old.xls"
    xls.write_bytes(b"x")
    code, out = run(capsys, "inspect", str(xls), "--json")
    assert code == 1 and out["error"] == "UnsupportedWorkbookError"
    code, out = run(capsys, "inspect", str(tmp_path / "missing.xlsx"), "--json")
    assert code == 1 and out["error"] == "FileNotFoundError"


def test_human_output_and_error_go_to_right_streams(capsys, xlsx_path):
    assert cli.main(["read-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2"]) == 0
    assert "[read-cell] outcome=read_cell" in capsys.readouterr().out
    assert cli.main(["read-cell", str(xlsx_path), "--sheet", "Nope", "--cell", "A2"]) == 1
    captured = capsys.readouterr()
    assert captured.out == "" and "SheetNotFoundError" in captured.err


def test_real_subprocess_json_contract(xlsx_path, tmp_path):
    env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}
    proc = subprocess.run([sys.executable, str(WORKSPACE_ROOT / "cli.py"), "inspect", str(xlsx_path), "--json"],
                          capture_output=True, text=True, env=env, cwd=tmp_path)
    assert proc.returncode == 0
    assert json.loads(proc.stdout)["outcome"] == "inspected" and proc.stdout.count("\n") == 1
    bad = subprocess.run([sys.executable, str(WORKSPACE_ROOT / "cli.py"), "patch-cell", str(xlsx_path), "--json"],
                         capture_output=True, text=True, env=env, cwd=tmp_path)
    assert bad.returncode == 1 and json.loads(bad.stdout)["error"] == "ArgumentError"


def test_truncate_flag_via_cli(capsys, xlsx_path):
    code, out = run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--value", "q" * 40000,
                    "--expected-revision", file_revision(xlsx_path), "--json")
    assert code == 1 and out["error"] == "TextTooLongError"
    code, out = run(capsys, "patch-cell", str(xlsx_path), "--sheet", "Data", "--cell", "A2", "--value", "q" * 40000,
                    "--truncate", "--expected-revision", file_revision(xlsx_path), "--json")
    assert code == 0


def test_cli_rejects_xlsm_patch_but_reads(capsys, tmp_path):
    macro = build_minimal_xlsx(tmp_path / "m.xlsm")
    code, out = run(capsys, "patch-cell", str(macro), "--sheet", "Data", "--cell", "A2", "--value", "x",
                    "--expected-revision", file_revision(macro), "--json")
    assert code == 1 and out["error"] == "UnsupportedWorkbookError"
    code, out = run(capsys, "inspect", str(macro), "--json")
    assert code == 0
