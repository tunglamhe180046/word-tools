/**
 * tools/word-engine/adapters/word_engine_client.ts - Adapter TypeScript dung chung cho cac phan
 * he Node/TypeScript (phan-tich/, "nhan vien ho so/") goi tools/word-engine/ qua subprocess CLI,
 * dung dung co che duoc quy dinh o docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md muc 1.2 ("Co che
 * goi tu TypeScript/Node"): child_process.execFile("python", ["tools/word-engine/cli.py",
 * "--json", ...]) - khong import ngon ngu cheo.
 *
 * File nay CHI la lop wrapper typed goi dung cli.py qua process boundary that (khong doc
 * core/*.py truc tiep) - moi logic patch/inspect/backup/restore van nam nguyen ven trong
 * tools/word-engine/core/*.py (xem cli.py module docstring). Chay bang Node native (khong can
 * bien dich/tsx): Node >= 22.6 ho tro strip type annotation cho .ts truc tiep, xem
 * tools/word-engine/tests/test_client.test.ts.
 *
 * Hop dong JSON stdout cua cli.py (--json): thanh cong -> exit 0 + {"success": true, ...}; loi ->
 * exit 1 + {"success": false, "error": "<TenException>", "reason": "<str>"}. Adapter nay chuyen
 * loi thanh WordEngineError (voi .code = TenException tu Python) thay vi de caller tu parse JSON
 * loi; khi tien trinh con “that bai” theo nghia khac (khong tim thay python, crash truoc khi kip
 * in JSON...) van nem WordEngineError voi code "WordEngineProcessError" de caller luon chi can
 * bat 1 loai loi duy nhat.
 *
 * Cac key JSON tra ve tu cli.py deu la snake_case (quy uoc Python cua repo nay) - adapter chuyen
 * sang camelCase dung quy uoc TypeScript truoc khi tra ve cho caller, qua ham toCamel() dung
 * chung cho moi lenh thay vi anh xa tung field thu cong (nac 4 cua CLAUDE.md goc: ham nho, cung
 * file, khi tai su dung tai cho khong du).
 */

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import path from "node:path";

const execFileAsync = promisify(execFile);

const CLI_PATH = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "cli.py",
);
const DEFAULT_PYTHON = process.env.WORD_ENGINE_PYTHON ?? "python";

// ---------------------------------------------------------------------------
// Error type
// ---------------------------------------------------------------------------
export class WordEngineError extends Error {
  /** Ten exception Python (vd "DocumentDriftError", "TextNotFoundError") hoac
   * "WordEngineProcessError" khi loi xay ra o tang subprocess (khong phai tu logic cli.py). */
  readonly code: string;

  constructor(code: string, reason: string) {
    super(reason);
    this.name = "WordEngineError";
    this.code = code;
  }
}

// ---------------------------------------------------------------------------
// Shared option types
// ---------------------------------------------------------------------------
export interface WordEngineCommonOptions {
  workDir?: string;
  jobId?: string;
  actor?: string;
  /** Duong dan/ten executable Python, mac dinh doc tu bien moi truong WORD_ENGINE_PYTHON hoac
   * "python" neu khong dat. */
  pythonPath?: string;
}

export interface PatchTextOptions extends WordEngineCommonOptions {
  targetId?: string;
  expectedRevision?: string;
  stylePolicy?: "prefer-start" | "prefer-end";
}

export interface GeometryOptions extends WordEngineCommonOptions {
  pageSize?: "A4" | "A5";
  margins?: "notary" | "standard" | "compact";
  pagination?: boolean;
  borders?: "all" | "none" | "horizontal-only" | "notary-standard";
  tableId?: string;
}

export interface StampOptions extends WordEngineCommonOptions {
  widthMm?: number;
  heightMm?: number;
  allowUnregistered?: boolean;
}

export interface ListBackupsOptions {
  workDir?: string;
  pythonPath?: string;
}

export interface RestoreOptions {
  workDir?: string;
  actor?: string;
  pythonPath?: string;
}

// ---------------------------------------------------------------------------
// Result types (camelCase - xem doc dau file ve toCamel())
// ---------------------------------------------------------------------------
export interface ObjectLocator {
  objectId: string;
  revision: string;
  kind: "paragraph" | "table_cell";
  structuralPath: string;
  expectedText: string;
  contextSha256: string;
  w14ParaId: string | null;
  tableIndex: number | null;
  rowIndex: number | null;
  colIndex: number | null;
}

export interface InspectionResult {
  documentRevision: string;
  locators: ObjectLocator[];
  jobId: string;
  indexFilePath: string;
}

export interface PatchResult {
  outcome: string;
  documentRevision: string;
  locators: ObjectLocator[];
  jobId: string;
  backupPath: string | null;
  replaceMode: string | null;
}

export interface BackupEntry {
  backupId: string;
  backupFile: string;
  targetPath: string;
  sha256: string;
  createdAt?: string;
  [key: string]: unknown;
}

export interface RestoreResult {
  outcome: string;
  documentRevision: string;
  jobId: string;
  backupPath: string | null;
  restoredFromBackupId: string | null;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------
function snakeToCamelKey(key: string): string {
  return key.replace(/_([a-z0-9])/g, (_match, c: string) => c.toUpperCase());
}

function toCamel(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(toCamel);
  }
  if (value !== null && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [key, val] of Object.entries(value as Record<string, unknown>)) {
      out[snakeToCamelKey(key)] = toCamel(val);
    }
    return out;
  }
  return value;
}

interface CliErrorPayload {
  success: false;
  error?: string;
  reason?: string;
}

async function runCli(args: string[], pythonPath: string): Promise<Record<string, unknown>> {
  let stdout: string;
  try {
    const result = await execFileAsync(pythonPath, [CLI_PATH, ...args], {
      encoding: "utf8",
      maxBuffer: 16 * 1024 * 1024,
    });
    stdout = result.stdout;
  } catch (err) {
    const execErr = err as NodeJS.ErrnoException & { stdout?: string; stderr?: string };
    if (typeof execErr.stdout === "string" && execErr.stdout.trim().length > 0) {
      let payload: CliErrorPayload | undefined;
      try {
        payload = JSON.parse(execErr.stdout) as CliErrorPayload;
      } catch {
        payload = undefined;
      }
      if (payload && payload.success === false) {
        throw new WordEngineError(
          payload.error ?? "UnknownError",
          payload.reason ?? execErr.message,
        );
      }
    }
    throw new WordEngineError(
      "WordEngineProcessError",
      execErr.stderr || execErr.message || String(err),
    );
  }

  const payload = JSON.parse(stdout) as { success: boolean; error?: string; reason?: string } & Record<
    string,
    unknown
  >;
  if (!payload.success) {
    throw new WordEngineError(
      payload.error ?? "UnknownError",
      payload.reason ?? "word-engine CLI reported failure without a reason.",
    );
  }
  return payload;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------
export async function inspectDocument(
  docxPath: string,
  options?: WordEngineCommonOptions,
): Promise<InspectionResult> {
  const args = ["inspect", docxPath, "--json"];
  if (options?.workDir) args.push("--work-dir", options.workDir);
  if (options?.jobId) args.push("--job-id", options.jobId);
  const payload = await runCli(args, options?.pythonPath ?? DEFAULT_PYTHON);
  return toCamel(payload) as unknown as InspectionResult;
}

export async function patchCell(
  docxPath: string,
  targetId: string,
  newText: string,
  expectedRevision: string,
  options?: WordEngineCommonOptions,
): Promise<PatchResult> {
  const args = [
    "patch-cell",
    docxPath,
    "--json",
    "--target-id",
    targetId,
    "--new-text",
    newText,
    "--expected-revision",
    expectedRevision,
  ];
  if (options?.workDir) args.push("--work-dir", options.workDir);
  if (options?.jobId) args.push("--job-id", options.jobId);
  if (options?.actor) args.push("--actor", options.actor);
  const payload = await runCli(args, options?.pythonPath ?? DEFAULT_PYTHON);
  return toCamel(payload) as unknown as PatchResult;
}

export async function patchText(
  docxPath: string,
  search: string,
  replace: string,
  options?: PatchTextOptions,
): Promise<PatchResult> {
  const args = ["patch-text", docxPath, "--json", "--search", search, "--replace", replace];
  if (options?.targetId) args.push("--target-id", options.targetId);
  if (options?.expectedRevision) args.push("--expected-revision", options.expectedRevision);
  if (options?.stylePolicy) args.push("--style-policy", options.stylePolicy);
  if (options?.workDir) args.push("--work-dir", options.workDir);
  if (options?.jobId) args.push("--job-id", options.jobId);
  if (options?.actor) args.push("--actor", options.actor);
  const payload = await runCli(args, options?.pythonPath ?? DEFAULT_PYTHON);
  return toCamel(payload) as unknown as PatchResult;
}

export async function setGeometry(docxPath: string, options: GeometryOptions): Promise<PatchResult> {
  const args = ["set-geometry", docxPath, "--json"];
  if (options.pageSize) args.push("--page-size", options.pageSize);
  if (options.margins) args.push("--margins", options.margins);
  if (options.pagination) args.push("--pagination");
  if (options.borders) args.push("--borders", options.borders);
  if (options.tableId) args.push("--table-id", options.tableId);
  if (options.workDir) args.push("--work-dir", options.workDir);
  if (options.jobId) args.push("--job-id", options.jobId);
  if (options.actor) args.push("--actor", options.actor);
  const payload = await runCli(args, options.pythonPath ?? DEFAULT_PYTHON);
  return toCamel(payload) as unknown as PatchResult;
}

export async function stampOps(
  docxPath: string,
  assetPath: string,
  targetId: string,
  expectedRevision: string,
  options?: StampOptions,
): Promise<PatchResult> {
  const args = [
    "stamp-ops",
    docxPath,
    "--json",
    "--asset-path",
    assetPath,
    "--target-id",
    targetId,
    "--expected-revision",
    expectedRevision,
  ];
  if (options?.widthMm !== undefined) args.push("--width-mm", String(options.widthMm));
  if (options?.heightMm !== undefined) args.push("--height-mm", String(options.heightMm));
  if (options?.allowUnregistered) args.push("--allow-unregistered");
  if (options?.workDir) args.push("--work-dir", options.workDir);
  if (options?.jobId) args.push("--job-id", options.jobId);
  if (options?.actor) args.push("--actor", options.actor);
  const payload = await runCli(args, options?.pythonPath ?? DEFAULT_PYTHON);
  return toCamel(payload) as unknown as PatchResult;
}

export async function listBackups(
  docxPath: string,
  options?: ListBackupsOptions,
): Promise<BackupEntry[]> {
  const args = ["backups", docxPath, "--json"];
  if (options?.workDir) args.push("--work-dir", options.workDir);
  const payload = await runCli(args, options?.pythonPath ?? DEFAULT_PYTHON);
  const camelPayload = toCamel(payload) as { backups: BackupEntry[] };
  return camelPayload.backups;
}

export async function restoreBackup(
  docxPath: string,
  backupId: string,
  options?: RestoreOptions,
): Promise<RestoreResult> {
  const args = ["restore", docxPath, "--json", "--backup-id", backupId];
  if (options?.workDir) args.push("--work-dir", options.workDir);
  if (options?.actor) args.push("--actor", options.actor);
  const payload = await runCli(args, options?.pythonPath ?? DEFAULT_PYTHON);
  return toCamel(payload) as unknown as RestoreResult;
}
