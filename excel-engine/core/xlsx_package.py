"""
tools/excel-engine/core/xlsx_package.py - Mo va dong goi lai goi OPC (ZIP) cua file .xlsx theo
tung entry (entry-by-entry), giu nguyen thu tu entry va noi dung byte cua moi part khong bi sua
(SHA-256 cua part khong doi).

Fail-closed khi mo: chi nhan ZIP that (magic bytes `PK\\x03\\x04`). File ma hoa Office/.xls cu
(OLE2, magic `D0CF11E0`), `.xls`, `.xlsb` (`xl/workbook.bin`), entry ZIP bi ma hoa, ten entry
thoat thu muc (zip-slip) hoac tong dung luong giai nen qua lon (zip-bomb) deu bi tu choi.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from typing import Dict, List, Union

ZIP_MAGIC = b"PK\x03\x04"
OLE2_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
REJECTED_SUFFIXES = {".xls", ".xlsb"}
MAX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_ENTRIES = 20_000


class XlsxFormatError(ValueError):
    pass


class UnsupportedWorkbookError(XlsxFormatError):
    """File ma hoa, .xls, .xlsb hoac dinh dang khong phai xlsx."""


def reject_legacy_suffix(path: Union[str, Path]) -> None:
    if Path(path).suffix.lower() in REJECTED_SUFFIXES:
        raise UnsupportedWorkbookError(f"Dinh dang {Path(path).suffix} khong duoc ho tro (chi .xlsx).")


@dataclass
class _Entry:
    name: str
    date_time: tuple
    compress_type: int
    external_attr: int
    create_system: int


class XlsxPackage:
    def __init__(self, entries: List[_Entry], parts: Dict[str, bytes]):
        self._entries: List[_Entry] = entries
        self._parts: Dict[str, bytes] = parts
        self._original_hashes: Dict[str, str] = {n: sha256(b).hexdigest() for n, b in parts.items()}

    # ---- construction ------------------------------------------------------
    @classmethod
    def open(cls, path: Union[str, Path]) -> "XlsxPackage":
        reject_legacy_suffix(path)
        return cls.from_bytes(Path(path).read_bytes())

    @classmethod
    def from_bytes(cls, data: bytes) -> "XlsxPackage":
        if data[:8] == OLE2_MAGIC:
            raise UnsupportedWorkbookError("File OLE2: workbook bi ma hoa hoac la dinh dang .xls cu.")
        if data[:4] != ZIP_MAGIC:
            raise UnsupportedWorkbookError("Khong phai goi ZIP/xlsx (thieu magic bytes PK\\x03\\x04).")
        try:
            zf = zipfile.ZipFile(io.BytesIO(data), "r")
        except zipfile.BadZipFile as exc:
            raise XlsxFormatError(f"ZIP hong: {exc}") from exc
        with zf:
            infos = zf.infolist()
            if len(infos) > MAX_ENTRIES:
                raise XlsxFormatError("Qua nhieu entry trong goi.")
            if sum(i.file_size for i in infos) > MAX_UNCOMPRESSED_BYTES:
                raise XlsxFormatError("Tong dung luong giai nen vuot gioi han (nghi zip-bomb).")
            entries: List[_Entry] = []
            parts: Dict[str, bytes] = {}
            for info in infos:
                name = info.filename
                if name.startswith("/") or "\\" in name or ".." in name.split("/"):
                    raise XlsxFormatError(f"Ten entry khong an toan: {name!r}")
                if name in parts:
                    raise XlsxFormatError(f"Entry trung ten: {name!r}")
                if info.flag_bits & 0x1:
                    raise UnsupportedWorkbookError(f"Entry {name!r} bi ma hoa.")
                entries.append(_Entry(name, info.date_time, info.compress_type, info.external_attr, info.create_system))
                parts[name] = zf.read(info)
        if "xl/workbook.bin" in parts:
            raise UnsupportedWorkbookError("Dinh dang .xlsb (xl/workbook.bin) khong duoc ho tro.")
        for required in ("[Content_Types].xml", "xl/workbook.xml"):
            if required not in parts:
                raise XlsxFormatError(f"Thieu part bat buoc: {required}")
        return cls(entries, parts)

    # ---- access -----------------------------------------------------------
    @property
    def names(self) -> List[str]:
        return [e.name for e in self._entries]

    def has(self, name: str) -> bool:
        return name in self._parts

    def get(self, name: str) -> bytes:
        try:
            return self._parts[name]
        except KeyError:
            raise XlsxFormatError(f"Khong co part {name!r} trong goi.") from None

    def set(self, name: str, data: bytes) -> None:
        if name not in self._parts:
            self._entries.append(_Entry(name, (2020, 1, 1, 0, 0, 0), zipfile.ZIP_DEFLATED, 0, 0))
        self._parts[name] = data

    def remove(self, name: str) -> None:
        if name in self._parts:
            del self._parts[name]
            self._entries = [e for e in self._entries if e.name != name]

    # ---- integrity --------------------------------------------------------
    def modified_parts(self) -> List[str]:
        """Ten cac part moi/da sua so voi luc mo (part bi xoa khong nam trong danh sach nay)."""
        return [n for n, b in self._parts.items() if self._original_hashes.get(n) != sha256(b).hexdigest()]

    def removed_parts(self) -> List[str]:
        return [n for n in self._original_hashes if n not in self._parts]

    def unchanged_hashes(self) -> Dict[str, str]:
        return {n: h for n, h in self._original_hashes.items() if n in self._parts and h == sha256(self._parts[n]).hexdigest()}

    # ---- serialization ----------------------------------------------------
    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as out:
            for entry in self._entries:
                info = zipfile.ZipInfo(entry.name, entry.date_time)
                info.compress_type = entry.compress_type
                info.external_attr = entry.external_attr
                info.create_system = entry.create_system
                out.writestr(info, self._parts[entry.name])
        return buf.getvalue()

