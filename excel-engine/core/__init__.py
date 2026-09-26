"""tools/excel-engine/core - self-contained core primitives for the dedicated Excel Engine.

Self-contained means exactly that: nothing under tools/excel-engine/ imports from
word-engine/, dich-thuat/, phan-tich/, or "nhan vien ho so/" (see root CLAUDE.md "Cau truc goc").
Where this project needs logic that already exists elsewhere (canonical path resolution, the
commit-broker safety pipeline), it keeps its own independent copy here rather than importing it.
"""
