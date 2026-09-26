#!/usr/bin/env python3
"""Unified CLI Dispatcher for Document Tools (Word & Excel Surgical Engines).

Automatically dispatches commands to:
  - word-engine/cli.py (for .docx files)
  - excel-engine/cli.py (for .xlsx/.xlsm files)

Or explicit routing via subcommands:
  - python cli.py word ...
  - python cli.py excel ...
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
WORD_CLI = ROOT_DIR / "word-engine" / "cli.py"
EXCEL_CLI = ROOT_DIR / "excel-engine" / "cli.py"


def _detect_engine(args: list[str]) -> Path:
    """Detect appropriate engine based on file argument extension."""
    for arg in args:
        lower = arg.lower()
        if lower.endswith(".docx"):
            return WORD_CLI
        if lower.endswith(".xlsx") or lower.endswith(".xlsm"):
            return EXCEL_CLI
    # Default to word-engine for backward compatibility
    return WORD_CLI


def main() -> None:
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print("Document Surgical Engine - Unified CLI Dispatcher")
        print("\nUsage:")
        print("  python cli.py word <args...>      # Invoke Word Tools directly")
        print("  python cli.py excel <args...>     # Invoke Excel Tools directly")
        print("  python cli.py <cmd> <file> ...    # Auto-dispatch by file extension (.docx / .xlsx)")
        print("\nEngines available:")
        print("  - word-engine/  : Microsoft Word (.docx) surgical patcher")
        print("  - excel-engine/ : Microsoft Excel (.xlsx) surgical patcher")
        sys.exit(0)

    first = args[0].lower()
    if first == "word":
        target_cli = WORD_CLI
        forward_args = args[1:]
    elif first == "excel":
        target_cli = EXCEL_CLI
        forward_args = args[1:]
    else:
        target_cli = _detect_engine(args)
        forward_args = args

    if not target_cli.exists():
        print(f"Error: Target engine CLI not found at {target_cli}", file=sys.stderr)
        sys.exit(1)

    cmd = [sys.executable, str(target_cli)] + forward_args
    proc = subprocess.run(cmd)
    sys.exit(proc.returncode)


if __name__ == "__main__":
    main()
