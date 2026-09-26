"""Pytest configuration for tools/word-engine/tests/.

Adds tools/word-engine/ itself (the parent of `core/`) to sys.path so tests can
`import core.safety_gateway` regardless of the directory pytest is invoked from -
mirrors dich-thuat/tests/conftest.py's WORKSPACE_ROOT pattern, scoped to this
project's own root instead (tools/word-engine/ is self-contained - see
core/safety_gateway.py's module docstring).
"""
import sys
from pathlib import Path

WORKSPACE_ROOT = Path(__file__).resolve().parent.parent
if str(WORKSPACE_ROOT) not in sys.path:
    sys.path.insert(0, str(WORKSPACE_ROOT))
