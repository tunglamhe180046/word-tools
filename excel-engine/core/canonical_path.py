"""
tools/excel-engine/core/canonical_path.py - Canonicalize a filesystem path for identity/
containment checks (case-fold, drive-letter normalization, symlink/directory-junction
resolution).

Why this exists: `safety_gateway.py`'s commit broker and Safe Restore Mutation Transaction
need to answer "is this the same file/directory" reliably before acquiring a write lock,
checking a target sits inside an allowed root, or keying a backup/state directory by target
identity. A plain string compare of two path spellings is not enough on this project's
Windows/OneDrive environment: NTFS is case-insensitive but case-preserving, and OneDrive can
route a path through a directory junction (see root CLAUDE.md, "Loi vat moi truong
Windows/OneDrive"), so two differently-spelled or differently-routed paths can point at the
same file on disk.

Ported 2026-09-24 from dich-thuat/engine/file_safety/canonical_path.py, per
docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 1.2: `tools/word-engine/` is
lower-level shared infrastructure that must NOT import from dich-thuat/ (or any of the
3 subprojects) - see .claude/rules/file-safety.md and root CLAUDE.md "Cau truc goc" for the
no-cross-project-import invariant this repo enforces everywhere else. This is therefore a
FIFTH independent copy of the same algorithm, alongside scripts/lib/canonical-path.mjs,
phan-tich/lib/storage/canonicalPath.ts, and dich-thuat/engine/file_safety/canonical_path.py -
same algorithm, four separate implementations, all reading the same test vectors
(scripts/__tests__/fixtures/canonical-paths.vectors.json) so "canonical" means the same thing
everywhere.

Also independent from scripts/lib/boundary-scanner.mjs's own toCanonicalPath(): that one is a
lightweight string normalizer for statically scanning source-code import literals (targets
usually do not exist on disk at scan time); this module resolves real filesystem state
(realpath through symlinks/junctions, on-disk case), which the scanner intentionally does not
need.
"""

from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass(frozen=True)
class CanonicalPath:
    absolute: str
    compare_key: str


def _resolve_existing_prefix(absolute_path: str) -> tuple[str, list[str]]:
    """Walk up from absolute_path until an existing ancestor is found, realpath it (resolving
    symlinks/junctions and on-disk case), then return that plus the not-yet-existing tail
    segments (closest-to-root last) to rejoin literally. If even the root does not resolve (e.g.
    a bogus drive letter), return absolute_path unchanged - a merely nonexistent path must not
    raise."""
    tail: list[str] = []
    current = absolute_path
    while True:
        if os.path.exists(current):
            return os.path.realpath(current), list(reversed(tail))
        parent = os.path.dirname(current)
        if parent == current:
            return current, list(reversed(tail))
        tail.append(os.path.basename(current))
        current = parent


def canonicalize_path(
    raw_path: str,
    base_dir: str | None = None,
    case_insensitive: bool | None = None,
) -> CanonicalPath:
    base_dir = base_dir or os.getcwd()
    p = raw_path.strip()
    resolved = os.path.abspath(p) if os.path.isabs(p) else os.path.abspath(os.path.join(base_dir, p))
    real, tail = _resolve_existing_prefix(resolved)
    absolute = os.path.join(real, *tail) if tail else real
    if case_insensitive is None:
        case_insensitive = os.name == "nt"
    compare_key = absolute.lower() if case_insensitive else absolute
    return CanonicalPath(absolute=absolute, compare_key=compare_key)


def is_path_inside(child_compare_key: str, parent_compare_key: str) -> bool:
    """So sanh 2 compare_key da chuan hoa: childKey co nam trong (hoac chinh la) parentKey khong."""
    if child_compare_key == parent_compare_key:
        return True
    prefix = parent_compare_key if parent_compare_key.endswith(os.sep) else parent_compare_key + os.sep
    return child_compare_key.startswith(prefix)
