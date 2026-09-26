/**
 * tools/word-engine/tests/test_client.test.ts - Kiem thu adapter TypeScript
 * (adapters/word_engine_client.ts) qua dung process boundary that (subprocess goi cli.py) - cung
 * tinh than voi tests/test_cli.py (Python): khong mock, khong goi thang core/*.py, chi dung cong
 * khai (inspectDocument/patchCell/...) nhu mot consumer TypeScript that (phan-tich/, "nhan vien
 * ho so/") se lam.
 *
 * Chay bang Node native (khong can cai them dependency nao - xem nac 1/nac 6 cua CLAUDE.md goc:
 * Node >= 22.6 tren may nay da ho tro strip type annotation cho .ts/import ".ts" truc tiep, da
 * kiem chung truoc khi viet file nay):
 *   node tools/word-engine/tests/test_client.test.ts
 *
 * Khong dung pytest/jest - tu viet mot test runner toi gian (mang cac async test + assert cua
 * chinh node:assert/strict) vi day la CACH DUY NHAT chay .ts khong can them dependency nao trong
 * tools/word-engine/ (thu muc nay co chu dich khong co package.json/node_modules rieng - xem
 * CLAUDE.md goc muc "Cau truc goc").
 */
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import {
  inspectDocument,
  listBackups,
  patchCell,
  patchText,
  restoreBackup,
  WordEngineError,
} from "../adapters/word_engine_client.ts";

const PYTHON = process.env.WORD_ENGINE_PYTHON ?? "python";

function buildSampleDocx(docxPath: string): void {
  const script = `
import sys
from docx import Document
doc = Document()
doc.add_paragraph("Doan van thu nhat.")
doc.add_paragraph("Doan van thu hai.")
table = doc.add_table(rows=1, cols=2)
table.cell(0, 0).text = "O 1"
table.cell(0, 1).text = "8.5"
doc.save(sys.argv[1])
`;
  execFileSync(PYTHON, ["-c", script, docxPath]);
}

function makeTmpDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), "word-engine-ts-client-"));
}

type TestFn = () => Promise<void>;
const tests: { name: string; fn: TestFn }[] = [];
function test(name: string, fn: TestFn): void {
  tests.push({ name, fn });
}

// ---------------------------------------------------------------------------
// inspectDocument
// ---------------------------------------------------------------------------
test("inspectDocument returns camelCase locators with a sha256 revision", async () => {
  const workDir = makeTmpDir();
  const docxPath = path.join(workDir, "sample.docx");
  buildSampleDocx(docxPath);

  const result = await inspectDocument(docxPath, { workDir });

  assert.ok(result.documentRevision.startsWith("sha256:"));
  const objectIds = result.locators.map((loc) => loc.objectId);
  assert.ok(objectIds.includes("para_1"));
  assert.ok(objectIds.includes("cell_t1_r1_c2"));
  const cellLocator = result.locators.find((loc) => loc.objectId === "cell_t1_r1_c2");
  assert.equal(cellLocator?.expectedText, "8.5");
  assert.equal(cellLocator?.kind, "table_cell");
});

// ---------------------------------------------------------------------------
// patchCell
// ---------------------------------------------------------------------------
test("patchCell updates the target cell and returns a new revision", async () => {
  const workDir = makeTmpDir();
  const docxPath = path.join(workDir, "sample.docx");
  buildSampleDocx(docxPath);
  const before = await inspectDocument(docxPath, { workDir });

  const result = await patchCell(docxPath, "cell_t1_r1_c2", "9.0", before.documentRevision, {
    workDir,
    jobId: "job-cell",
  });

  assert.equal(result.outcome, "committed_clean");
  assert.notEqual(result.documentRevision, before.documentRevision);

  const after = await inspectDocument(docxPath, { workDir });
  const cellLocator = after.locators.find((loc) => loc.objectId === "cell_t1_r1_c2");
  assert.equal(cellLocator?.expectedText, "9.0");
});

test("patchCell with a stale expectedRevision throws WordEngineError(DocumentDriftError)", async () => {
  const workDir = makeTmpDir();
  const docxPath = path.join(workDir, "sample.docx");
  buildSampleDocx(docxPath);

  await assert.rejects(
    () =>
      patchCell(
        docxPath,
        "cell_t1_r1_c2",
        "9.0",
        "sha256:0000000000000000000000000000000000000000000000000000000000000000",
        { workDir },
      ),
    (err: unknown) => {
      assert.ok(err instanceof WordEngineError);
      assert.equal(err.code, "DocumentDriftError");
      return true;
    },
  );
});

// ---------------------------------------------------------------------------
// patchText
// ---------------------------------------------------------------------------
test("patchText without a locator replaces the matching paragraph", async () => {
  const workDir = makeTmpDir();
  const docxPath = path.join(workDir, "sample.docx");
  buildSampleDocx(docxPath);

  const result = await patchText(docxPath, "Doan van thu nhat.", "Doan van da duoc sua.", {
    workDir,
  });

  assert.equal(result.outcome, "committed_clean");
  const after = await inspectDocument(docxPath, { workDir });
  const paragraph = after.locators.find((loc) => loc.objectId === "para_1");
  assert.equal(paragraph?.expectedText, "Doan van da duoc sua.");
});

// ---------------------------------------------------------------------------
// listBackups / restoreBackup
// ---------------------------------------------------------------------------
test("listBackups is empty before any write and gains an entry after patchCell", async () => {
  const workDir = makeTmpDir();
  const docxPath = path.join(workDir, "sample.docx");
  buildSampleDocx(docxPath);

  const emptyBackups = await listBackups(docxPath, { workDir });
  assert.deepEqual(emptyBackups, []);

  const before = await inspectDocument(docxPath, { workDir });
  await patchCell(docxPath, "cell_t1_r1_c2", "9.0", before.documentRevision, { workDir });

  const backups = await listBackups(docxPath, { workDir });
  assert.equal(backups.length, 1);
  assert.ok(backups[0].sha256);
  assert.ok(backups[0].backupId);
});

test("restoreBackup reverts the document to the backed-up content", async () => {
  const workDir = makeTmpDir();
  const docxPath = path.join(workDir, "sample.docx");
  buildSampleDocx(docxPath);
  const before = await inspectDocument(docxPath, { workDir });
  await patchCell(docxPath, "cell_t1_r1_c2", "9.0", before.documentRevision, { workDir });

  const backups = await listBackups(docxPath, { workDir });
  const result = await restoreBackup(docxPath, backups[0].backupId, { workDir });

  assert.equal(result.outcome, "restored_from_backup");
  const after = await inspectDocument(docxPath, { workDir });
  const cellLocator = after.locators.find((loc) => loc.objectId === "cell_t1_r1_c2");
  assert.equal(cellLocator?.expectedText, "8.5");
});

// ---------------------------------------------------------------------------
// Runner
// ---------------------------------------------------------------------------
async function main(): Promise<void> {
  let failures = 0;
  for (const { name, fn } of tests) {
    try {
      await fn();
      console.log(`PASS ${name}`);
    } catch (err) {
      failures += 1;
      console.error(`FAIL ${name}`);
      console.error(err);
    }
  }
  console.log(`\n${tests.length - failures}/${tests.length} tests passed.`);
  if (failures > 0) {
    process.exitCode = 1;
  }
}

await main();
