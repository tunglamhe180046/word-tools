"""
tools/word-engine/core/safety_gateway.py - Commit Broker (Control Plane) cho tools/word-engine/,
cong duy nhat cho moi thao tac ghi de len file production trong Word Engine.

Nguon goc: port nguyen trang thuat toan Commit Broker 18 buoc da duoc kiem toan boi Codex Terra
va nghiem thu tai commit `7189b37` cua dich-thuat/engine/file_safety/commit_broker.py, gop voi
cac primitive baseline-tracking cua dich-thuat/engine/docx_guard.py (sha256_bytes, sha256_file,
atomic_write_bytes, file_lock, DocumentStateManifest/inspect_document_state/
record_generated_baseline). Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 1.2:
"Phuong an trien khai Core Broker" - tuyet doi khong viet lai thuat toan tu dau de khong mo lai
cac lo hong da duoc kiem toan sua (xem .claude/rules/file-safety.md "Buoc 2"/"Buoc 3").

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" - day la ban
sao doc lap thu 2 cua thuat toan Commit Broker (ban dau: dich-thuat/engine/file_safety/
commit_broker.py), dung `core/canonical_path.py` cua chinh tools/word-engine/ (khong phai
dich-thuat/engine/file_safety/canonical_path.py).

3 diem KHAC BIET co chu dich so voi ban goc (khong phai loi port, ghi ro de nguoi doc sau
khong nham la thieu trung thuc voi ban da kiem toan):

1. **`.jarvis/` duoc neo theo tham so `work_dir` thay vi ngam dinh theo cwd tien trinh.**
   Ban goc dung `Path(".jarvis/...").resolve()` truc tiep (ngam dinh: thu muc lam viec hien tai
   cua tien trinh). O day, moi ham cong khai nhan them tham so `work_dir: Optional[...] = None`
   - goc `.jarvis/` = `Path(work_dir).resolve() / ".jarvis"` neu co, nguoc lai fallback ve cwd
   (giu nguyen hanh vi cu khi khong truyen). Ly do: tools/word-engine/ la ha tang dung chung,
   phai goi duoc tu bat ky thu muc nao (3 du an con qua subprocess CLI) va phai test duoc trong
   1 thu muc cach ly (tmp_path) ma khong dung vao `.jarvis/` that cua repo.
2. **Buoc 10 (backup production file) ghi vao `.jarvis/backups/<target_key>/` (tap trung, co
   `<id>.meta.json` kem theo) thay vi thu muc `_backup/` nam canh file goc.** Day la yeu cau
   tuong minh cua ke hoach (section 3.4 "Safe Restore as Mutation Transaction": "backups
   <file.docx>: Liet ke cac ban sao luu trong `.jarvis/backups/`") va cung nhat quan voi cach
   `jobs/`, `audit/`, `secrets/`, `conflicts/` cua chinh file nay da tap trung duoi `.jarvis/`.
   Neu giu nguyen quy uoc `_backup/` canh file cua ban goc, `list_backups()`/`restore_backup()`
   se khong tim thay dung backup ma Buoc 10 vua tao ra. Cau truc 18 buoc, thu tu, dieu kien va
   hanh vi khoa/hash/atomic-replace cua Buoc 10 giu nguyen 100% - chi vi tri luu backup doi.
3. **Buoc 15 (cap nhat baseline trang thai tai lieu) chay cho MOI loai file, khong chi `.docx`.**
   Ban goc gioi han `if ext.lower() == ".docx"` vi dich-thuat/ chi ghi .docx qua broker nay.
   `safety_gateway.py` la ha tang chung cho moi loai file Word Engine co the ghi (khong chi
   .docx), nen dieu kien tracking baseline duoc mo rong. Buoc 9 (DOCX Zip Package Integrity
   Check) van CHI ap dung cho `.docx` nhu ban goc - do la kiem tra dac thu dinh dang, khong phai
   baseline tracking.

Ngoai 3 diem tren, thu tu 18 buoc, cac exception, HMAC canonical JSON signing, Dedicated Commit
Lease Lock (`.commit.lock`), Ticket Lock (`.lock`), va postcondition verification giu nguyen
tuyet doi so voi ban da kiem toan.

Bo sung Phase 1 - Safe Restore as Mutation Transaction (`list_backups`, `restore_backup`):
xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 3.4. `restore_backup()` tu no KHONG
viet lai logic ghi de - no tai su dung dung `issue_job()` + `request_commit()` (Buoc 10 se tu
dong tao 1 backup cua trang thai HIEN TAI truoc khi ghi de bang noi dung backup duoc phuc hoi),
chi gan nhan lai outcome thanh "restored_from_backup" sau khi commit thanh cong.
"""

from __future__ import annotations

import json
import os
import secrets
import shutil
import time
import uuid
import zipfile
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from hashlib import sha256
from hmac import HMAC
from pathlib import Path
from typing import Any, Callable, Dict, List, Literal, Optional, Union

from core.canonical_path import canonicalize_path, is_path_inside


# ---------------------------------------------------------------------------
# Baseline-tracking primitives (ported from dich-thuat/engine/docx_guard.py)
# ---------------------------------------------------------------------------
def sha256_bytes(data: bytes) -> str:
    return sha256(data).hexdigest()


def sha256_file(path: Union[str, Path]) -> str:
    return sha256_bytes(Path(path).read_bytes())


def atomic_write_bytes(
    path: Union[str, Path],
    data: bytes,
    before_replace: Optional[Callable[[], None]] = None,
) -> None:
    """Write beside the target, then rename onto it directly - a single atomic replace on both
    POSIX (rename(2)) and Windows (`os.replace` uses `MoveFileExW` with
    `MOVEFILE_REPLACE_EXISTING`). Mirrors dich-thuat/engine/docx_guard.py::atomic_write_bytes
    (2026-09-23 Codex Terra audit fix: single-step replace, no separate "displace old file"
    rename that could lose the target if the process is killed mid-way)."""
    path = Path(path)
    nonce = f"{os.getpid() % 100000}-{os.urandom(4).hex()}"
    tmp_path = path.with_name(f".{path.name}.tmp-{nonce}")
    try:
        tmp_path.write_bytes(data)
        if before_replace is not None:
            before_replace()
        os.replace(tmp_path, path)
    except Exception:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except OSError:
            pass
        raise


class FileLockTimeout(RuntimeError):
    pass


@contextmanager
def file_lock(target_path: Union[str, Path], timeout_s: float = 15.0, work_dir: Optional[Union[str, Path]] = None):
    """Cooperative lock via an exclusive-create lockfile beside the target. Mirrors
    dich-thuat/engine/docx_guard.py::file_lock - canonicalizes `target_path` internally so every
    alias (different case, path through a junction...) of the same real file maps to the same
    `.write.lock`."""
    target_canonical = canonicalize_path(str(target_path), base_dir=str(work_dir) if work_dir else None)
    target_path = Path(target_canonical.absolute)
    lock_path = target_path.with_name(target_path.name + ".write.lock")
    deadline = time.monotonic() + timeout_s
    fd = None
    while fd is None:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            if time.monotonic() >= deadline:
                raise FileLockTimeout(f"Timed out waiting for write lock: {lock_path}")
            time.sleep(0.04 + secrets.randbelow(80) / 1000.0)
    try:
        try:
            os.write(fd, json.dumps({"pid": os.getpid(), "created_at": datetime.now(timezone.utc).isoformat()}).encode("utf-8"))
        finally:
            os.close(fd)
        yield
    finally:
        try:
            lock_path.unlink()
        except OSError:
            pass


DOCUMENT_STATE_SCHEMA = "word-engine-document-state.v1"

StatusKind = Literal["missing", "untracked", "clean", "modified", "invalid_state"]


@dataclass
class DocumentStateManifest:
    schema: str
    revision: int
    document_name: str
    baseline_docx_sha256: str
    baseline_file: str
    generated_at: str

    def to_json(self) -> str:
        return json.dumps(self.__dict__, indent=2, ensure_ascii=False) + "\n"

    @staticmethod
    def from_json(raw: str) -> "DocumentStateManifest":
        data = json.loads(raw)
        required = ("schema", "revision", "document_name", "baseline_docx_sha256", "baseline_file", "generated_at")
        missing = [key for key in required if key not in data]
        if missing:
            raise ValueError(f"manifest missing required field(s): {', '.join(missing)}")
        if data["schema"] != DOCUMENT_STATE_SCHEMA:
            raise ValueError(f"unrecognized manifest schema: {data['schema']!r}")
        return DocumentStateManifest(**{key: data[key] for key in required})


@dataclass
class DocumentStatus:
    kind: StatusKind
    docx_path: Path
    current_hash: Optional[str] = None
    manifest: Optional[DocumentStateManifest] = None
    reason: Optional[str] = None


def _jarvis_root(work_dir: Optional[Union[str, Path]] = None) -> Path:
    base = Path(work_dir).resolve() if work_dir is not None else Path.cwd()
    return base / ".jarvis"


def _target_key(target_path: Union[str, Path], work_dir: Optional[Union[str, Path]] = None) -> str:
    """Opaque, filesystem-safe identity key for a target path - used to namespace its backups
    (`.jarvis/backups/<key>/`) and its document-state manifest (`.jarvis/state/<key>/`) so two
    differently-spelled aliases of the same real file (case, junction...) share one history, per
    core/canonical_path.py."""
    canonical = canonicalize_path(str(target_path), base_dir=str(work_dir) if work_dir else None)
    return sha256(canonical.compare_key.encode("utf-8")).hexdigest()[:24]


def document_state_paths(
    docx_path: Union[str, Path], work_dir: Optional[Union[str, Path]] = None
) -> tuple[Path, Path, Path]:
    """(state_dir, manifest_path, baseline_path) for one tracked file - centralized under
    `.jarvis/state/<target_key>/`, unlike dich-thuat/engine/docx_guard.py's sibling
    `_document-state/<name>/` (see module docstring, deviation 2's centralization rationale
    applies here too)."""
    docx_path = Path(docx_path)
    target_key = _target_key(docx_path, work_dir)
    state_dir = _jarvis_root(work_dir) / "state" / target_key
    baseline_name = f"baseline{docx_path.suffix or '.bin'}"
    return state_dir, state_dir / "manifest.json", state_dir / baseline_name


def inspect_document_state(
    docx_path: Union[str, Path], work_dir: Optional[Union[str, Path]] = None
) -> DocumentStatus:
    docx_path = Path(docx_path)
    if not docx_path.exists():
        return DocumentStatus(kind="missing", docx_path=docx_path)

    current_hash = sha256_file(docx_path)
    _, manifest_path, baseline_path = document_state_paths(docx_path, work_dir)

    if not manifest_path.exists():
        return DocumentStatus(kind="untracked", docx_path=docx_path, current_hash=current_hash)

    try:
        manifest = DocumentStateManifest.from_json(manifest_path.read_text(encoding="utf-8"))
    except (ValueError, json.JSONDecodeError) as exc:
        return DocumentStatus(
            kind="invalid_state", docx_path=docx_path, current_hash=current_hash,
            reason=f"invalid manifest {manifest_path}: {exc}",
        )

    if not baseline_path.exists():
        return DocumentStatus(
            kind="invalid_state", docx_path=docx_path, current_hash=current_hash,
            reason=f"missing baseline file: {baseline_path}",
        )

    baseline_hash = sha256_file(baseline_path)
    if baseline_hash != manifest.baseline_docx_sha256:
        return DocumentStatus(
            kind="invalid_state", docx_path=docx_path, current_hash=current_hash,
            reason=f"baseline file hash does not match manifest: {baseline_path}",
        )

    kind: StatusKind = "clean" if current_hash == manifest.baseline_docx_sha256 else "modified"
    return DocumentStatus(kind=kind, docx_path=docx_path, current_hash=current_hash, manifest=manifest)


def record_generated_baseline(
    docx_path: Union[str, Path],
    generated_bytes: bytes,
    work_dir: Optional[Union[str, Path]] = None,
    previous: Optional[DocumentStateManifest] = None,
) -> DocumentStateManifest:
    docx_path = Path(docx_path)
    state_dir, manifest_path, baseline_path = document_state_paths(docx_path, work_dir)
    state_dir.mkdir(parents=True, exist_ok=True)
    atomic_write_bytes(baseline_path, generated_bytes)
    manifest = DocumentStateManifest(
        schema=DOCUMENT_STATE_SCHEMA,
        revision=(previous.revision if previous else 0) + 1,
        document_name=docx_path.name,
        baseline_docx_sha256=sha256_bytes(generated_bytes),
        baseline_file=baseline_path.name,
        generated_at=datetime.now(timezone.utc).isoformat(),
    )
    atomic_write_bytes(manifest_path, manifest.to_json().encode("utf-8"))
    return manifest


class DocumentModifiedError(RuntimeError):
    """File da bi thay doi tren dia so voi baseline duoc ticket ghi nhan luc issue_job (nhieu kha
    nang la mot chinh sua tay). Mirrors dich-thuat/engine/docx_guard.py::DocumentModifiedError."""

    def __init__(self, reason: str, docx_path: Path, candidate_path: Path):
        super().__init__(f"{reason}. Kept the existing file untouched: {docx_path} - candidate: {candidate_path}")
        self.reason = reason
        self.docx_path = docx_path
        self.candidate_path = candidate_path


class DocumentLockedError(RuntimeError):
    """Phat hien Word dang mo file dich (owner-file `~$...`). Mirrors
    dich-thuat/engine/docx_guard.py::DocumentLockedError."""

    def __init__(self, reason: str, docx_path: Path, candidate_path: Path):
        super().__init__(f"{reason}. Kept the existing file untouched: {docx_path} - candidate: {candidate_path}")
        self.reason = reason
        self.docx_path = docx_path
        self.candidate_path = candidate_path


# ---------------------------------------------------------------------------
# Exceptions (ported from dich-thuat/engine/file_safety/commit_broker.py)
# ---------------------------------------------------------------------------
class JobExpiredError(Exception):
    def __init__(self, job_id: str):
        super().__init__(f"Job Ticket {job_id} da het han.")
        self.job_id = job_id


class JobAlreadyCommittedError(Exception):
    def __init__(self, job_id: str):
        super().__init__(f"Job Ticket {job_id} da o trang thai committed (terminal).")
        self.job_id = job_id


class JobTicketTamperedError(Exception):
    def __init__(self, job_id: str, reason: str):
        super().__init__(f"Job Ticket {job_id} bi thay doi bat hop phap: {reason}")
        self.job_id = job_id
        self.reason = reason


class JobRecoveryRequiredError(Exception):
    def __init__(self, job_id: str):
        super().__init__(f"Job Ticket {job_id} dang o trang thai recovery_required.")
        self.job_id = job_id


class JobConcurrentCommitError(Exception):
    def __init__(self, job_id: str):
        super().__init__(f"Job Ticket {job_id} dang duoc commit boi tien trinh khac.")
        self.job_id = job_id


class CandidateMutatedDuringSnapshotError(Exception):
    def __init__(self, candidate_path: Path):
        super().__init__(f"candidate {candidate_path} bi thay doi trong luc snapshot (hash1 != hash2).")
        self.candidate_path = candidate_path


class PathEscapesWorkDirError(Exception):
    def __init__(self, candidate_path: Path, work_dir: Path):
        super().__init__(f"candidate_path {candidate_path} nam ngoai work_dir {work_dir}.")
        self.candidate_path = candidate_path
        self.work_dir = work_dir


class PostconditionVerificationError(Exception):
    def __init__(self, target_path: Path, expected_hash: str, actual_hash: str):
        super().__init__(f"Postcondition verification failed cho {target_path}: expected {expected_hash}, got {actual_hash}")
        self.target_path = target_path
        self.expected_hash = expected_hash
        self.actual_hash = actual_hash


class RestoreConflictError(RuntimeError):
    """Nem boi restore_backup() khi hash dia hien tai cua target khac baseline/last-known-hash da
    ghi nhan - nguoi dung da tu sua tay tren dia ke tu lan cuoi he thong ghi nhan trang thai.
    Tu choi ghi de, khong tu dong resolve."""

    def __init__(self, target_path: str, current_hash: Optional[str] = None, baseline_hash: Optional[str] = None):
        super().__init__(
            "Target document has uncommitted manual modifications on disk. "
            "Refusing to overwrite user edits without confirmation."
        )
        self.target_path = target_path
        self.current_hash = current_hash
        self.baseline_hash = baseline_hash


class BackupNotFoundError(RuntimeError):
    def __init__(self, target_path: str, backup_id: str):
        super().__init__(f"Backup {backup_id} khong ton tai cho {target_path}.")
        self.target_path = target_path
        self.backup_id = backup_id


# ---------------------------------------------------------------------------
# Data Models & Canonical JSON Serialization
# ---------------------------------------------------------------------------
@dataclass
class JobTicket:
    job_id: str
    created_at: str
    expires_at: str
    actor: str
    allowed_roots: List[str]
    target_path: str
    work_dir: str
    status: str
    baseline_sha256: Optional[str] = None
    updated_at: Optional[str] = None
    task_id: Optional[str] = None
    signature: Optional[str] = None
    outcome: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def canonical_json_dumps(obj: Any) -> str:
    """Sort keys recursive, khong dung join(',') de tranh collision."""
    if obj is None or isinstance(obj, (bool, int, float, str)):
        return json.dumps(obj, ensure_ascii=False)
    if isinstance(obj, list):
        return "[" + ",".join(canonical_json_dumps(x) for x in obj) + "]"
    if isinstance(obj, dict):
        sorted_keys = sorted(obj.keys())
        return "{" + ",".join(f"{json.dumps(k, ensure_ascii=False)}:{canonical_json_dumps(obj[k])}" for k in sorted_keys) + "}"
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")


def _get_or_create_hmac_key(work_dir: Optional[Union[str, Path]] = None) -> bytes:
    secret_dir = _jarvis_root(work_dir) / "secrets"
    key_path = secret_dir / "hmac.key"
    lock_path = secret_dir / "hmac.key.lock"

    if key_path.exists():
        return key_path.read_bytes()

    secret_dir.mkdir(parents=True, exist_ok=True)
    start = time.time()
    while time.time() - start < 5.0:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                if key_path.exists():
                    os.close(fd)
                    try:
                        lock_path.unlink()
                    except OSError:
                        pass
                    return key_path.read_bytes()

                new_key = secrets.token_bytes(32)
                key_path.write_bytes(new_key)
                os.close(fd)
                try:
                    lock_path.unlink()
                except OSError:
                    pass
                return new_key
            except Exception:
                os.close(fd)
                try:
                    lock_path.unlink()
                except OSError:
                    pass
                raise
        except FileExistsError:
            time.sleep(0.05)
            if key_path.exists():
                return key_path.read_bytes()

    return key_path.read_bytes()


def sign_ticket(ticket: JobTicket, work_dir: Optional[Union[str, Path]] = None) -> str:
    key = _get_or_create_hmac_key(work_dir)
    payload = {
        "actor": ticket.actor,
        "allowed_roots": sorted(ticket.allowed_roots),
        "baseline_sha256": ticket.baseline_sha256,
        "created_at": ticket.created_at,
        "expires_at": ticket.expires_at,
        "job_id": ticket.job_id,
        "status": ticket.status,
        "target_path": ticket.target_path,
        "task_id": ticket.task_id if ticket.task_id is not None else None,
        "work_dir": ticket.work_dir,
    }
    canonical = canonical_json_dumps(payload)
    return HMAC(key, canonical.encode("utf-8"), sha256).hexdigest()


def verify_ticket_signature(ticket: JobTicket, work_dir: Optional[Union[str, Path]] = None) -> bool:
    if not ticket.signature:
        return False
    expected = sign_ticket(ticket, work_dir)
    return secrets.compare_digest(expected, ticket.signature)


# ---------------------------------------------------------------------------
# Lock Helpers (commit.lock lease & micro-lock)
# ---------------------------------------------------------------------------
class MicroLock:
    def __init__(self, lock_path: Path, timeout: float = 5.0):
        self.lock_path = lock_path
        self.timeout = timeout
        self.fd: Optional[int] = None

    def acquire(self) -> "MicroLock":
        self.lock_path.parent.mkdir(parents=True, exist_ok=True)
        start = time.time()
        while time.time() - start < self.timeout:
            try:
                self.fd = os.open(str(self.lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                payload = json.dumps({"pid": os.getpid(), "time": time.time()}).encode("utf-8")
                os.write(self.fd, payload)
                return self
            except FileExistsError:
                try:
                    mtime = self.lock_path.stat().st_mtime
                    if time.time() - mtime > 30.0:  # 30s stale lock
                        self.lock_path.unlink()
                        continue
                except OSError:
                    pass
                time.sleep(0.04)
        raise TimeoutError(f"Timed out acquiring lock: {self.lock_path}")

    def release(self) -> None:
        if self.fd is not None:
            try:
                os.close(self.fd)
            except OSError:
                pass
            self.fd = None
        try:
            self.lock_path.unlink()
        except OSError:
            pass

    def __enter__(self) -> "MicroLock":
        return self.acquire()

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.release()


def append_audit_log(entry: Dict[str, Any], work_dir: Optional[Union[str, Path]] = None) -> None:
    audit_dir = _jarvis_root(work_dir) / "audit"
    audit_dir.mkdir(parents=True, exist_ok=True)
    ym = datetime.now(timezone.utc).strftime("%Y-%m")
    log_file = audit_dir / f"{ym}.jsonl"
    lock_path = _jarvis_root(work_dir) / "audit.lock"

    with MicroLock(lock_path, timeout=5.0):
        entry_copy = dict(entry)
        entry_copy["timestamp"] = datetime.now(timezone.utc).isoformat()
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry_copy, ensure_ascii=False) + "\n")


def _default_case_data_roots(work_dir: Optional[Union[str, Path]] = None) -> List[str]:
    """Cac thu muc du lieu ho so canonical o goc repo (khop root CLAUDE.md "Cau truc goc" va
    scripts/lib/boundary-scanner.mjs's CASE_DATA_DIRS) - dung lam allowed_roots mac dinh cho
    issue_job() khi caller khong tu truyen. KHAC danh sach dich-thuat-specific
    (_case-state/demo/.local-case-storage/cases/Google_Drive_Phong_Ho_So) cua ban goc
    commit_broker.py, vi tools/word-engine/ la ha tang dung chung cho ca 3 du an con, khong rieng
    dich-thuat/."""
    base = Path(work_dir).resolve() if work_dir is not None else Path.cwd()
    names = [".local-case-storage", "demo", "_json-goc", ".uploads", "scratch"]
    return [str(Path(canonicalize_path(name, base_dir=str(base)).absolute)) for name in names]


# ---------------------------------------------------------------------------
# Public API: issue_job()
# ---------------------------------------------------------------------------
def issue_job(
    target_path: Union[str, Path],
    actor: str,
    task_id: Optional[str] = None,
    ttl_seconds: int = 7200,
    allowed_roots: Optional[List[str]] = None,
    work_dir: Optional[Union[str, Path]] = None,
) -> JobTicket:
    base_dir = str(Path(work_dir).resolve()) if work_dir is not None else None
    canonical = canonicalize_path(str(target_path), base_dir=base_dir)
    target_abs = Path(canonical.absolute)

    now = datetime.now(timezone.utc)
    expires_at = datetime.fromtimestamp(now.timestamp() + ttl_seconds, tz=timezone.utc).isoformat()

    roots = allowed_roots or _default_case_data_roots(work_dir)
    target_key = canonical.compare_key
    norm_roots = [str(canonicalize_path(r, base_dir=base_dir).absolute) for r in roots]
    norm_roots_keys = [canonicalize_path(r, base_dir=base_dir).compare_key for r in roots]

    # Containment check luc issue dung is_path_inside chuan
    is_contained = any(is_path_inside(target_key, r_key) for r_key in norm_roots_keys)
    if not is_contained:
        raise ValueError(f"target_path {target_abs} khong nam trong allowed_roots.")

    baseline_hash = sha256_file(target_abs) if target_abs.exists() else None
    job_id = str(uuid.uuid4())
    job_work_dir = _jarvis_root(work_dir) / "work" / job_id
    job_work_dir.mkdir(parents=True, exist_ok=True)

    ticket = JobTicket(
        job_id=job_id,
        created_at=now.isoformat(),
        expires_at=expires_at,
        actor=actor,
        allowed_roots=norm_roots,
        target_path=str(target_abs),
        work_dir=str(job_work_dir),
        status="active",
        baseline_sha256=baseline_hash,
        task_id=task_id,
    )
    ticket.signature = sign_ticket(ticket, work_dir)

    jobs_dir = _jarvis_root(work_dir) / "jobs"
    jobs_dir.mkdir(parents=True, exist_ok=True)
    ticket_file = jobs_dir / f"{job_id}.json"
    ticket_file.write_text(json.dumps(ticket.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")

    return ticket


# ---------------------------------------------------------------------------
# Public API: request_commit() (18-step pipeline)
# ---------------------------------------------------------------------------
def request_commit(job_id: str, work_dir: Optional[Union[str, Path]] = None) -> Dict[str, Any]:
    commit_id = str(uuid.uuid4())
    jobs_dir = _jarvis_root(work_dir) / "jobs"
    ticket_file = jobs_dir / f"{job_id}.json"
    commit_lock_path = jobs_dir / f"{job_id}.commit.lock"
    ticket_lock_path = jobs_dir / f"{job_id}.lock"

    if not ticket_file.exists():
        raise FileNotFoundError(f"Khong tim thay Job Ticket {job_id}")

    ticket_data = json.loads(ticket_file.read_text(encoding="utf-8"))
    ticket = JobTicket(**ticket_data)

    # Buoc 0b: Gianh Dedicated Commit Lease Lock (giu suot tu Buoc 0 den Buoc 16)
    commit_lock = MicroLock(commit_lock_path, timeout=0.5)
    try:
        commit_lock.acquire()
    except TimeoutError:
        raise JobConcurrentCommitError(job_id)

    target_lock_ctx = None
    did_replace = False
    staged_path: Optional[Path] = None
    staged_hash: Optional[str] = None
    target_abs = Path(ticket.target_path)
    job_work_dir = Path(ticket.work_dir)
    ext = target_abs.suffix

    try:
        # Buoc 0c: Gianh ticket.lock & verify ticket
        with MicroLock(ticket_lock_path, timeout=5.0):
            fresh_ticket = JobTicket(**json.loads(ticket_file.read_text(encoding="utf-8")))
            if not verify_ticket_signature(fresh_ticket, work_dir):
                append_audit_log({"commit_id": commit_id, "job_id": job_id, "event": "DENIED", "outcome": "denied_ticket_invalid"}, work_dir)
                raise JobTicketTamperedError(job_id, "Chu ky HMAC khong hop le")

            if fresh_ticket.status == "committed":
                raise JobAlreadyCommittedError(job_id)
            if fresh_ticket.status == "recovery_required":
                raise JobRecoveryRequiredError(job_id)
            if fresh_ticket.status == "committing":
                raise JobConcurrentCommitError(job_id)
            if fresh_ticket.status != "active":
                raise ValueError(f"Trang thai ticket khong hop le: {fresh_ticket.status}")

            exp_time = datetime.fromisoformat(fresh_ticket.expires_at).timestamp()
            if time.time() > exp_time:
                fresh_ticket.status = "expired"
                ticket_file.write_text(json.dumps(fresh_ticket.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
                raise JobExpiredError(job_id)

            # Buoc 0d: Chuyen active -> committing
            fresh_ticket.status = "committing"
            fresh_ticket.updated_at = datetime.now(timezone.utc).isoformat()
            fresh_ticket.signature = sign_ticket(fresh_ticket, work_dir)
            ticket_file.write_text(json.dumps(fresh_ticket.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
            ticket = fresh_ticket

        # Buoc 1: Audit STARTED
        append_audit_log({
            "commit_id": commit_id,
            "job_id": job_id,
            "event": "STARTED",
            "actor": ticket.actor,
            "target_path": ticket.target_path,
            "baseline_sha256": ticket.baseline_sha256,
        }, work_dir)

        # Buoc 2: Path resolution & Containment
        candidate_path = job_work_dir / f"candidate{ext}"
        try:
            candidate_path.resolve().relative_to(job_work_dir.resolve())
        except ValueError:
            raise PathEscapesWorkDirError(candidate_path, job_work_dir)

        # Buoc 3: Snapshot with Triple-Hash
        if not candidate_path.exists():
            raise FileNotFoundError(f"Khong tim thay file candidate: {candidate_path}")
        h1 = sha256_file(candidate_path)
        staged_path = job_work_dir / f"staged{ext}"
        shutil.copy2(candidate_path, staged_path)
        h2 = sha256_file(candidate_path)
        if h1 != h2:
            staged_path.unlink(missing_ok=True)
            raise CandidateMutatedDuringSnapshotError(candidate_path)

        staged_hash = sha256_file(staged_path)
        if staged_hash != h1:
            staged_path.unlink(missing_ok=True)
            raise ValueError("Staged hash khong khop voi candidate hash ban dau.")

        # Buoc 5: Audit STAGED
        append_audit_log({
            "commit_id": commit_id,
            "job_id": job_id,
            "event": "STAGED",
            "staged_sha256": staged_hash,
        }, work_dir)

        # Buoc 6: Word owner file deny check
        if ext.lower() == ".docx":
            base_name = target_abs.name
            if len(base_name) > 2:
                owner_file = target_abs.parent / f"~${base_name[2:]}"
                if owner_file.exists():
                    raise DocumentLockedError(f"Phat hien Word owner-file: {owner_file}", target_abs, candidate_path)

        # Buoc 7: Gianh target.write.lock
        target_lock_ctx = file_lock(target_abs, work_dir=work_dir)
        target_lock_ctx.__enter__()

        # Buoc 8: Optimistic Concurrency Check
        current_target_hash = sha256_file(target_abs) if target_abs.exists() else None
        if ticket.baseline_sha256 is not None and current_target_hash != ticket.baseline_sha256:
            raise DocumentModifiedError("File da bi thay doi tren dia so voi luc issue job.", target_abs, candidate_path)

        # Buoc 9: DOCX Zip Package Integrity Check (neu .docx)
        if ext.lower() == ".docx":
            try:
                with zipfile.ZipFile(staged_path, "r") as zf:
                    if zf.testzip() is not None or "[Content_Types].xml" not in zf.namelist():
                        raise ValueError("File .docx khong hop le hoac bi loi zip archive.")
            except Exception as zip_err:
                raise ValueError(f"DOCX package integrity check fail: {zip_err}")

        # Buoc 10: Backup production file - tap trung duoi .jarvis/backups/<target_key>/ (xem
        # module docstring, diem khac biet 2)
        backup_path: Optional[str] = None
        if target_abs.exists():
            target_key = _target_key(target_abs, work_dir)
            backup_dir = _jarvis_root(work_dir) / "backups" / target_key
            backup_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%f")
            backup_id = f"{ts}-{uuid.uuid4().hex[:8]}"
            backup_file = backup_dir / f"{backup_id}{ext}"
            shutil.copy2(target_abs, backup_file)
            pre_backup_hash = sha256_file(backup_file)
            backup_meta = {
                "schema": "jarvis-backup.v1",
                "backup_id": backup_id,
                "target_path": str(target_abs),
                "created_at": datetime.now(timezone.utc).isoformat(),
                "sha256": pre_backup_hash,
                "size_bytes": backup_file.stat().st_size,
                "backup_file": str(backup_file),
                "job_id": job_id,
                "commit_id": commit_id,
            }
            (backup_dir / f"{backup_id}.meta.json").write_text(
                json.dumps(backup_meta, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
            )
            backup_path = str(backup_file)

        # Buoc 11: Tao file tam tren cung volume
        nonce = f"{os.getpid()}-{int(time.time()*1000)}-{secrets.token_hex(4)}"
        temp_file = target_abs.parent / f".{target_abs.name}.tmp-{nonce}"
        shutil.copy2(staged_path, temp_file)

        # Buoc 12 & 13: Atomic Replace
        did_replace = False
        os.replace(temp_file, target_abs)
        did_replace = True

        # Audit REPLACED
        append_audit_log({
            "commit_id": commit_id,
            "job_id": job_id,
            "event": "REPLACED",
            "staged_sha256": staged_hash,
        }, work_dir)

        # Buoc 14: Postcondition Verification
        post_hash = sha256_file(target_abs)
        if post_hash != staged_hash:
            raise PostconditionVerificationError(target_abs, staged_hash, post_hash)

        # Buoc 15: Cap nhat baseline trang thai tai lieu - cho MOI loai file (xem module
        # docstring, diem khac biet 3), khong chi .docx nhu ban goc
        try:
            record_generated_baseline(target_abs, target_abs.read_bytes(), work_dir=work_dir)
        except Exception:
            pass

        # Buoc 16: Terminal State COMMITTED
        with MicroLock(ticket_lock_path, timeout=5.0):
            ticket.status = "committed"
            ticket.outcome = "committed_clean"
            ticket.updated_at = datetime.now(timezone.utc).isoformat()
            ticket.signature = sign_ticket(ticket, work_dir)
            ticket_file.write_text(json.dumps(ticket.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")

        # Thu tu nha lock: target.write.lock -> commit.lock
        if target_lock_ctx is not None:
            target_lock_ctx.__exit__(None, None, None)
            target_lock_ctx = None

        commit_lock.release()

        # Buoc 17: Audit COMMITTED
        append_audit_log({
            "commit_id": commit_id,
            "job_id": job_id,
            "event": "COMMITTED",
            "outcome": "committed_clean",
            "sha256": post_hash,
        }, work_dir)

        if staged_path and staged_path.exists():
            staged_path.unlink()

        return {
            "success": True,
            "job_id": job_id,
            "target_path": str(target_abs),
            "sha256": post_hash,
            "backup_path": backup_path,
            "outcome": "committed_clean",
        }

    except Exception as exc:
        next_status = "active"
        outcome = "failed_pre_replace"

        if did_replace:
            if isinstance(exc, PostconditionVerificationError):
                next_status = "recovery_required"
                outcome = "recovery_required_postcondition_failure"
            else:
                next_status = "committed"
                outcome = "committed_manifest_update_failed"
        else:
            next_status = "active"
            outcome = f"failed_error_{exc.__class__.__name__}"

        try:
            with MicroLock(ticket_lock_path, timeout=5.0):
                ticket.status = next_status
                ticket.outcome = outcome
                ticket.updated_at = datetime.now(timezone.utc).isoformat()
                ticket.signature = sign_ticket(ticket, work_dir)
                ticket_file.write_text(json.dumps(ticket.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

        append_audit_log({
            "commit_id": commit_id,
            "job_id": job_id,
            "event": "FAILED",
            "outcome": outcome,
            "error": str(exc),
            "did_replace": did_replace,
        }, work_dir)

        if next_status != "recovery_required" and staged_path and staged_path.exists():
            staged_path.unlink(missing_ok=True)

        raise exc

    finally:
        if target_lock_ctx is not None:
            try:
                target_lock_ctx.__exit__(None, None, None)
            except Exception:
                pass
        commit_lock.release()


# ---------------------------------------------------------------------------
# Public API: Safe Restore as Mutation Transaction
# (docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 3.4)
# ---------------------------------------------------------------------------
def list_backups(target_path: Union[str, Path], work_dir: Optional[Union[str, Path]] = None) -> List[Dict[str, Any]]:
    """Liet ke cac ban sao luu da duoc Buoc 10 cua request_commit() tao ra cho target_path, moi
    lan file da ton tai bi ghi de - moi (docx-)ban ghi trong .jarvis/backups/<target_key>/, moi
    hoc theo *.meta.json. Sap xep moi den cu."""
    target_key = _target_key(target_path, work_dir)
    backup_dir = _jarvis_root(work_dir) / "backups" / target_key
    if not backup_dir.exists():
        return []
    entries: List[Dict[str, Any]] = []
    for meta_file in sorted(backup_dir.glob("*.meta.json")):
        try:
            entries.append(json.loads(meta_file.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            continue
    entries.sort(key=lambda e: e.get("created_at", ""), reverse=True)
    return entries


def restore_backup(
    target_path: Union[str, Path],
    backup_id: str,
    work_dir: Optional[Union[str, Path]] = None,
    allowed_roots: Optional[List[str]] = None,
    actor: str = "system",
) -> Dict[str, Any]:
    """Phuc hoi target_path ve dung mot ban backup_id da liet ke boi list_backups(). Day la mot
    mutation transaction co Optimistic Lock (section 3.4): neu hash dia hien tai cua target khac
    baseline/last-known-hash da ghi nhan (nguoi dung da sua tay tren dia ke tu lan cuoi he thong
    ghi nhan trang thai), tu choi va nem RestoreConflictError - khong bao gio am tham ghi de.

    Khi khong xung dot: dua noi dung backup lam candidate roi chay qua dung Commit Broker 18
    buoc (issue_job + request_commit, tu dong tao 1 backup cua trang thai HIEN TAI truoc khi ghi
    de o Buoc 10 - phuc hoi cung la mot lan ghi, khong phai duong tat rieng), roi gan nhan lai
    outcome thanh "restored_from_backup"."""
    base_dir = str(Path(work_dir).resolve()) if work_dir is not None else None
    canonical = canonicalize_path(str(target_path), base_dir=base_dir)
    target_abs = Path(canonical.absolute)

    status = inspect_document_state(target_abs, work_dir=work_dir)
    if status.kind in ("modified", "untracked", "invalid_state"):
        raise RestoreConflictError(
            str(target_abs),
            current_hash=status.current_hash,
            baseline_hash=status.manifest.baseline_docx_sha256 if status.manifest else None,
        )

    backups = list_backups(target_abs, work_dir=work_dir)
    backup_entry = next((b for b in backups if b.get("backup_id") == backup_id), None)
    if backup_entry is None:
        raise BackupNotFoundError(str(target_abs), backup_id)

    backup_file = Path(backup_entry["backup_file"])
    if not backup_file.exists():
        raise FileNotFoundError(f"Backup file missing on disk: {backup_file}")

    ticket = issue_job(target_abs, actor=actor, allowed_roots=allowed_roots, work_dir=work_dir)
    ext = target_abs.suffix
    candidate_path = Path(ticket.work_dir) / f"candidate{ext}"
    shutil.copy2(backup_file, candidate_path)

    result = dict(request_commit(ticket.job_id, work_dir=work_dir))
    result["outcome"] = "restored_from_backup"
    result["restored_from_backup_id"] = backup_id

    append_audit_log({
        "job_id": ticket.job_id,
        "event": "RESTORED_FROM_BACKUP",
        "target_path": str(target_abs),
        "backup_id": backup_id,
        "restored_sha256": backup_entry.get("sha256"),
    }, work_dir)

    return result
