"""tools/word-engine/core - self-contained core primitives for the dedicated Word Engine.

Self-contained means exactly that: nothing under tools/word-engine/ imports from
dich-thuat/, phan-tich/, or "nhan vien ho so/" (see root CLAUDE.md "Cau truc goc" and
docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 1.2 - "Xoa bo Coupling nguoc").
Where this project needs logic that already exists in one of those three subprojects
(canonical path resolution, the commit-broker safety pipeline), it keeps its own
independent copy here rather than importing it.
"""
