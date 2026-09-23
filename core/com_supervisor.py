"""
tools/word-engine/core/com_supervisor.py - Word COM Supervisor & Dual-Layer QA (Phase 4) cho
tools/word-engine/.

Xem docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md muc 3.5 ("Quy trinh Visual QA Chuan"), muc 3.6
("Bat bien Van hanh Word COM Supervisor"), muc 4 ("Phase 4: Word COM Supervisor & Visual QA") va
muc 5.2 ("Profile 2: COM FALLBACK MODE").

KHONG import bat ky module nao tu dich-thuat/, phan-tich/ hay "nhan vien ho so/" (xem
core/__init__.py va docs/plans muc 1.2). Logic watchdog PID+create_time va cach resolve Word PID
qua unique Caption + FindWindowW ("OpusApp") LAY CAM HUNG tu
dich-thuat/engine/word_com_worker.py::resolve_word_pid_from_caption - day la mot ban cai dat DOC
LAP, khong import, dung tinh than nac 2 cua CLAUDE.md goc (khong tai su dung ma nguon xuyen du an,
chi tai su dung Y TUONG). Cung ly do, `WordComUnavailableError` o day KHONG phai cung class voi
`dich-thuat/engine/word_com.py::WordComUnavailableError` du trung ten va trung tinh than "khong bia
dat du kien" (audit R21 cua dich-thuat/: khong co duong fallback tao PDF gia/uoc luong khi Word
COM khong san sang).

Module nay chiu trach nhiem 2 phan doc lap:

1. **Word COM Supervisor** (`ComSupervisor`, plan muc 3.6) - dedicated `WINWORD.EXE` worker so
   huu boi chinh no:
   - Luon `win32com.client.DispatchEx("Word.Application")` (KHONG BAO GIO `GetActiveObject`/
     `Dispatch` - 2 ham do co the attach vao 1 instance Word dang mo san cua nguoi dung, vi pham
     truc tiep bat bien "tuyet doi khong bao gio attach vao WINWORD dang mo cua nguoi dung").
   - `Visible=False`, `DisplayAlerts=0`, `ScreenUpdating=False` ngay sau khoi tao.
   - Resolve dung PID cua tien trinh WINWORD vua spawn (khong doan qua danh sach tien trinh he
     thong) bang cach dat 1 `Caption` duy nhat (uuid4) roi `FindWindowW("OpusApp", caption)` +
     `GetWindowThreadProcessId` - ghi lai CA `pid` LAN `create_time` (qua psutil) ngay sau do.
   - Watchdog (`_Watchdog`): 1 thread nen giam sat song song job; het `watchdog_timeout_s` ma job
     chua xong thi buoc dung dung tien trinh da ghi lai o tren.
   - Bat bien an toan cung (`_terminate_owned_process`/`_process_is_same`): CHI kill 1 pid khi
     ca pid VA create_time con khop voi luc supervisor tu spawn - PID bi he dieu hanh tai su dung
     cho 1 tien trinh hoan toan khac (xay ra tu nhien sau khi tien trinh cu ket thuc) se KHONG BAO
     GIO bi dung nham, du watchdog het gio.
   - Serialized execution (`run_job()`): dung `threading.Lock()` dam bao CHI 1 job COM chay tai 1
     thoi diem cho 1 instance `ComSupervisor`.
   - Cleanup (`_quit()`): luon `Quit()` + doi tien trinh tu ket thuc trong `finally` cua
     `run_job()`, ke ca khi `fn(app)` nem loi.
   - Graceful fallback (`is_available()`): False (khong spawn gi ca, chi kiem tra platform +
     import pywin32) khi khong phai Windows / thieu pywin32 - moi ham phu thuoc Word COM
     (`export_pdf`) nem `WordComUnavailableError` ro rang thay vi crash hoac tao ket qua gia.
   - Cong thuc `mm` -> `points` (Shapes.AddPicture, plan muc 3.6): `mm_to_points()`.

2. **Dual-Layer QA (Profile 2)** (plan muc 3.5/5.2):
   - Layer 1 (Structural VML Validation) - `validate_vml_integrity()`: so sanh cay `w:pict`/
     `v:shape`/`v:imagedata` va quan he hinh anh (`r:id`) giua 2 ban `word/document.xml` da parse
     (before/after), phat hien VML bi rung. `require_vml_integrity()` nem `VmlIntegrityError` neu
     khong dat. `inspect_vml_signature()` la wrapper 1-file (doc dung tu docx_path) dung boi CLI
     `validate-vml`.
   - Layer 2 (Visual Masked Diff) - pipeline `export_pdf()` (qua `ComSupervisor`) -> PDF ->
     `render_pdf_page_to_pixmap()` (PyMuPDF rasterize) -> `compute_masked_visual_diff()`: loai tru
     `authorized_rois` (vung toa do pixel duoc phep sua) roi so sanh PIXEL cua phan con lai, tra
     ve `pixel_diff_ratio` (phai = 0.0 ngoai ROI). `require_masked_visual_diff_clean()` nem
     `VisualDiffViolationError` neu khong dat.
"""

from __future__ import annotations

import platform
import threading
import time
import uuid
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple, TypeVar, Union

from lxml import etree

from core.inspector import WORD_NS

_V_NS = "urn:schemas-microsoft-com:vml"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

MM_TO_POINTS = 72.0 / 25.4
WD_EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0

T = TypeVar("T")
Roi = Tuple[int, int, int, int]  # (x0, y0, x1, y1) toa do pixel, nua-mo [x0,x1) x [y0,y1)


def mm_to_points(value_mm: float) -> float:
    """Cong thuc chuyen doi chuan cho Shapes.AddPicture (plan muc 3.6): 1 mm = 72 / 25.4 points."""
    return value_mm * MM_TO_POINTS


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------
class WordComUnavailableError(RuntimeError):
    """Word COM khong san sang tren may nay (khong phai Windows, thieu pywin32, hoac Word chua
    duoc cai dat/khoi tao duoc) - khong co duong fallback tao PDF gia/uoc luong (xem module
    docstring)."""


class WatchdogTimeoutError(RuntimeError):
    """Job COM vuot qua watchdog timeout - tien trinh WINWORD so huu boi supervisor da bi buoc
    dung (xem _Watchdog/_terminate_owned_process)."""


class VmlIntegrityError(RuntimeError):
    """Nem boi require_vml_integrity() khi validate_vml_integrity() phat hien VML bi rung
    (Layer 1, plan muc 5.2 diem 3)."""


class VisualDiffViolationError(RuntimeError):
    """Nem boi require_masked_visual_diff_clean() khi pixel_diff_ratio ngoai authorized_rois > 0
    (Layer 2, plan muc 5.2 diem 4)."""


# ---------------------------------------------------------------------------
# Platform / availability
# ---------------------------------------------------------------------------
def is_windows() -> bool:
    return platform.system().lower() == "windows"


def is_available() -> bool:
    """True neu Word COM CO THE kha dung tren may nay - KHONG spawn Word, chi kiem tra dieu kien
    can (Windows + pywin32 import duoc). Khong dam bao Word THAT SU da cai dat (DispatchEx co the
    van that bai sau do); do la ly do _start() van tu bat WordComUnavailableError rieng khi
    DispatchEx that bai du is_available() == True."""
    if not is_windows():
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


# ---------------------------------------------------------------------------
# Watchdog identity checks - bat bien an toan cung (plan muc 3.6)
# ---------------------------------------------------------------------------
def _process_is_same(pid: int, expected_create_time: float) -> bool:
    """True CHI khi PID con dang chay VA create_time khop expected_create_time (dung sai < 1s) -
    chan tuyet doi viec he dieu hanh tai su dung 1 PID cho tien trinh khac danh lua watchdog kill
    nham (xem module docstring, bat bien an toan)."""
    import psutil

    try:
        proc = psutil.Process(pid)
    except psutil.NoSuchProcess:
        return False
    try:
        return abs(proc.create_time() - expected_create_time) < 1.0
    except psutil.Error:
        return False


def _terminate_owned_process(pid: int, expected_create_time: float) -> bool:
    """Buoc dung 1 tien trinh CHI khi danh tinh (pid + create_time) khop dung tien trinh
    supervisor da tu spawn - tra False (khong lam gi ca) neu khong khop hoac tien trinh khong con
    ton tai. Day la HAM DUY NHAT trong module nay duoc phep goi terminate()/kill() tren mot PID
    Word - moi loi goi terminate khac deu phai di qua ham nay."""
    import psutil

    if not _process_is_same(pid, expected_create_time):
        return False
    try:
        proc = psutil.Process(pid)
        proc.terminate()
        try:
            proc.wait(timeout=5.0)
        except psutil.TimeoutExpired:
            proc.kill()
        return True
    except psutil.NoSuchProcess:
        return False


class _Watchdog:
    """Giam sat 1 job COM song song tren 1 thread nen: het `timeout_s` ma job chua goi cancel()
    thi buoc dung dung tien trinh (pid, create_time) da ghi nhan - tuyet doi khong dung tien trinh
    nao khac (xem _terminate_owned_process). Thiet ke tach rieng khoi ComSupervisor de co the unit
    test bang 1 tien trinh Python thuong (khong can Word that) - xem tests/test_com_roundtrip.py."""

    def __init__(self, pid: int, create_time: float, timeout_s: float):
        self._pid = pid
        self._create_time = create_time
        self._timeout_s = timeout_s
        self._cancel_event = threading.Event()
        self._triggered = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def start(self) -> "_Watchdog":
        self._thread.start()
        return self

    def _run(self) -> None:
        if not self._cancel_event.wait(self._timeout_s):
            if _terminate_owned_process(self._pid, self._create_time):
                self._triggered.set()

    def cancel(self) -> None:
        self._cancel_event.set()
        self._thread.join(timeout=5.0)

    @property
    def triggered(self) -> bool:
        return self._triggered.is_set()


def _resolve_word_pid_via_caption(word_app: Any, timeout_s: float = 5.0) -> int:
    """Resolve DUNG PID cua tien trinh WINWORD vua duoc DispatchEx() spawn: dat 1 Caption duy
    nhat (uuid4) tren chinh instance do, roi FindWindowW(class "OpusApp", caption do) +
    GetWindowThreadProcessId - khong doan qua danh sach tien trinh WINWORD dang chay tren he
    thong (co the co WINWORD cua nguoi dung dang mo song song), dung tinh than bat bien "tuyet
    doi khong bao gio attach vao WINWORD dang mo cua nguoi dung" (xem module docstring)."""
    if not is_windows():
        raise WordComUnavailableError("Khong the resolve Word PID ngoai Windows.")

    import ctypes
    import ctypes.wintypes

    unique_caption = f"__WORD_ENGINE_SUPERVISOR_{uuid.uuid4().hex}__"
    word_app.Caption = unique_caption

    user32 = ctypes.windll.user32
    deadline = time.monotonic() + timeout_s
    hwnd = 0
    while time.monotonic() < deadline:
        hwnd = user32.FindWindowW("OpusApp", unique_caption)
        if hwnd:
            break
        time.sleep(0.05)
    if not hwnd:
        raise WordComUnavailableError(
            "Khong tim thay window handle cua Word (class OpusApp) sau khi dat unique Caption."
        )

    pid = ctypes.wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value)


# ---------------------------------------------------------------------------
# ComSupervisor - dedicated Word COM worker (plan muc 3.6)
# ---------------------------------------------------------------------------
class ComSupervisor:
    """Dedicated Word COM Worker doc lap - 1 instance quan ly toi da 1 tien trinh WINWORD.EXE do
    CHINH no spawn, thuc thi job tuan tu (serialized) qua run_job(). Xem module docstring cho day
    du bat bien an toan."""

    def __init__(self, watchdog_timeout_s: float = 120.0):
        self._watchdog_timeout_s = watchdog_timeout_s
        self._lock = threading.Lock()
        self._app: Any = None
        self._owned_pid: Optional[int] = None
        self._owned_create_time: Optional[float] = None

    @staticmethod
    def is_available() -> bool:
        return is_available()

    def _start(self) -> None:
        if not self.is_available():
            raise WordComUnavailableError(
                "Word COM khong kha dung tren may nay (khong phai Windows hoac thieu pywin32)."
            )

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            app = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            raise WordComUnavailableError(f"Khong the khoi tao Word.Application COM: {exc}") from exc

        try:
            app.Visible = False
            app.DisplayAlerts = 0
            app.ScreenUpdating = False
            pid = _resolve_word_pid_via_caption(app)
            import psutil

            create_time = psutil.Process(pid).create_time()
        except Exception:
            try:
                app.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            raise

        self._app = app
        self._owned_pid = pid
        self._owned_create_time = create_time

    def _quit(self) -> None:
        if self._app is not None:
            try:
                for doc in list(self._app.Documents):
                    try:
                        doc.Close(WD_DO_NOT_SAVE_CHANGES)
                    except Exception:
                        pass
                self._app.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass

        if self._owned_pid is not None and self._owned_create_time is not None:
            # Don dep binh thuong (khong phai lop cuong che) - cho toi da 5s de tien trinh tu
            # ket thuc sau Quit(). Neu qua khong may van con song, run_job() da co Watchdog rieng
            # chiu trach nhiem buoc dung trong luc job dang chay - o day chi la don dep hau-job.
            deadline = time.monotonic() + 5.0
            while time.monotonic() < deadline and _process_is_same(self._owned_pid, self._owned_create_time):
                time.sleep(0.1)

        self._app = None
        self._owned_pid = None
        self._owned_create_time = None
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:
            pass

    def run_job(self, fn: Callable[[Any], T]) -> T:
        """Serialized execution (plan muc 3.6): self._lock dam bao CHI 1 job chay tai 1 thoi diem
        cho instance nay. fn(app) nhan Word.Application COM object da Visible=False/
        DisplayAlerts=0/ScreenUpdating=False san. Watchdog PID+create_time giam sat song song
        trong suot job; qua watchdog_timeout_s ma fn(app) chua tra ve thi bi buoc dung dung tien
        trinh WINWORD do CHINH supervisor nay spawn (khong bao gio dung tien trinh khac) va
        run_job() nem WatchdogTimeoutError. finally luon Quit() + kill-if-still-alive, ke ca khi
        fn(app) tu no nem loi."""
        with self._lock:
            self._start()
            assert self._owned_pid is not None and self._owned_create_time is not None
            watchdog = _Watchdog(self._owned_pid, self._owned_create_time, self._watchdog_timeout_s).start()
            try:
                result = fn(self._app)
            finally:
                watchdog.cancel()
                self._quit()
            if watchdog.triggered:
                raise WatchdogTimeoutError(
                    f"Word COM job vuot qua watchdog timeout {self._watchdog_timeout_s}s - da "
                    "buoc dung tien trinh WINWORD so huu boi supervisor."
                )
            return result

    def export_pdf(self, docx_path: Union[str, Path], pdf_path: Union[str, Path]) -> Path:
        """Word Native PDF Export (plan muc 3.5): doc.ExportAsFixedFormat(ExportFormat=17)."""
        docx_path = Path(docx_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        if not docx_path.exists():
            raise FileNotFoundError(f"Khong tim thay file docx: {docx_path}")
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        def _job(app: Any) -> Path:
            doc = app.Documents.Open(
                FileName=str(docx_path), ReadOnly=True, AddToRecentFiles=False, Visible=False,
            )
            try:
                doc.ExportAsFixedFormat(OutputFileName=str(pdf_path), ExportFormat=WD_EXPORT_FORMAT_PDF)
            finally:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            return pdf_path

        self.run_job(_job)
        if not pdf_path.exists() or pdf_path.stat().st_size == 0:
            raise RuntimeError(f"Word COM ExportAsFixedFormat khong tao ra file PDF hop le: {pdf_path}")
        return pdf_path


def export_pdf(
    docx_path: Union[str, Path], pdf_path: Union[str, Path], watchdog_timeout_s: float = 120.0
) -> Path:
    """Helper cap module: tao 1 ComSupervisor dung 1 lan roi export_pdf() - dung boi CLI
    `export-pdf` (moi loi goi CLI la 1 tien trinh rieng, khong can tai su dung instance)."""
    return ComSupervisor(watchdog_timeout_s=watchdog_timeout_s).export_pdf(docx_path, pdf_path)


# ---------------------------------------------------------------------------
# Layer 1: Structural VML Validation (plan muc 3.5/5.2 diem 3)
# ---------------------------------------------------------------------------
def load_document_xml_root_from_docx(docx_path: Union[str, Path]) -> etree._Element:
    with zipfile.ZipFile(Path(docx_path), "r") as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def _collect_vml_signature(root: etree._Element) -> Dict[str, Any]:
    """Quet toan bo w:pict trong root, ghi lai danh tinh cau truc (khong phai vi tri/kich thuoc -
    do la trach nhiem cua Layer 2) cua tung v:shape/v:imagedata ben trong."""
    pict_signatures: List[Dict[str, Any]] = []
    shape_count = 0
    imagedata_count = 0
    for pict in root.iter(f"{{{WORD_NS}}}pict"):
        shapes = list(pict.iter(f"{{{_V_NS}}}shape"))
        imagedatas = list(pict.iter(f"{{{_V_NS}}}imagedata"))
        shape_count += len(shapes)
        imagedata_count += len(imagedatas)
        pict_signatures.append(
            {
                "shape_ids": sorted(s.get("id") or "" for s in shapes),
                "shape_types": sorted(s.get("type") or "" for s in shapes),
                "imagedata_rids": sorted(img.get(f"{{{_R_NS}}}id") or "" for img in imagedatas),
            }
        )
    return {
        "pict_count": len(pict_signatures),
        "shape_count": shape_count,
        "imagedata_count": imagedata_count,
        "picts": pict_signatures,
    }


def validate_vml_integrity(before_root: etree._Element, after_root: etree._Element) -> Dict[str, Any]:
    """Layer 1 (plan muc 3.5/5.2 diem 3): so sanh cay the VML (w:pict, v:shape, v:imagedata) va
    quan he hinh anh (r:id) giua before_root/after_root (2 ban word/document.xml da parse san) -
    phat hien VML bi rung (mat pict/shape/imagedata, hoac mat quan he r:id). Tra ve dict co khoa
    "valid" (bool) va "dropped" (list mo ta cac vi pham, rong neu valid)."""
    before_sig = _collect_vml_signature(before_root)
    after_sig = _collect_vml_signature(after_root)

    dropped: List[str] = []
    if after_sig["pict_count"] < before_sig["pict_count"]:
        dropped.append(f"w:pict count giam tu {before_sig['pict_count']} xuong {after_sig['pict_count']}")
    if after_sig["shape_count"] < before_sig["shape_count"]:
        dropped.append(f"v:shape count giam tu {before_sig['shape_count']} xuong {after_sig['shape_count']}")
    if after_sig["imagedata_count"] < before_sig["imagedata_count"]:
        dropped.append(
            f"v:imagedata count giam tu {before_sig['imagedata_count']} xuong {after_sig['imagedata_count']}"
        )

    before_rids = {rid for p in before_sig["picts"] for rid in p["imagedata_rids"] if rid}
    after_rids = {rid for p in after_sig["picts"] for rid in p["imagedata_rids"] if rid}
    missing_rids = before_rids - after_rids
    if missing_rids:
        dropped.append(f"v:imagedata r:id bi mat quan he hinh anh: {sorted(missing_rids)}")

    return {"valid": not dropped, "before": before_sig, "after": after_sig, "dropped": dropped}


def require_vml_integrity(before_root: etree._Element, after_root: etree._Element) -> Dict[str, Any]:
    report = validate_vml_integrity(before_root, after_root)
    if not report["valid"]:
        raise VmlIntegrityError("VML_INTEGRITY_VIOLATION: " + "; ".join(report["dropped"]))
    return report


def inspect_vml_signature(docx_path: Union[str, Path]) -> Dict[str, Any]:
    """Wrapper 1-file dung boi CLI `validate-vml <docx_path>`: doc dung word/document.xml cua
    docx_path va tra ve chu ky VML cua no (khong so sanh voi ban nao khac - so sanh before/after
    thuc su la validate_vml_integrity(), dung noi bo khi mot thao tac COM/patch khac can doi
    chieu round-trip)."""
    root = load_document_xml_root_from_docx(docx_path)
    return _collect_vml_signature(root)


# ---------------------------------------------------------------------------
# Layer 2: Masked Visual Diff (plan muc 3.5/5.2 diem 4)
# ---------------------------------------------------------------------------
def render_pdf_page_to_pixmap(pdf_path: Union[str, Path], page_index: int = 0, dpi: int = 150) -> Any:
    """DOCX (da xuat qua ComSupervisor.export_pdf truoc do) -> PDF -> PyMuPDF rasterize ->
    Pixmap, dung pipeline chuan hoa cua plan muc 3.5."""
    import pymupdf

    doc = pymupdf.open(str(Path(pdf_path)))
    try:
        if page_index < 0 or page_index >= doc.page_count:
            raise IndexError(f"page_index {page_index} vuot qua so trang {doc.page_count}")
        zoom = dpi / 72.0
        matrix = pymupdf.Matrix(zoom, zoom)
        return doc[page_index].get_pixmap(matrix=matrix, alpha=False)
    finally:
        doc.close()


def _point_in_rois(x: int, y: int, rois: Sequence[Roi]) -> bool:
    return any(x0 <= x < x1 and y0 <= y < y1 for x0, y0, x1, y1 in rois)


def compute_masked_visual_diff(
    before_pix: Any, after_pix: Any, authorized_rois: Optional[Sequence[Roi]] = None
) -> Dict[str, Any]:
    """Layer 2 (plan muc 3.5/5.2 diem 4): so sanh pixel giua before_pix/after_pix - 2 doi tuong
    "Pixmap-like" (co .width, .height, .n la so kenh mau, .samples la bytes tho hang-major, dung
    dinh dang cua fitz/pymupdf.Pixmap) - LOAI TRU cac vung chu nhat trong authorized_rois (toa do
    pixel, nua-mo [x0,x1) x [y0,y1) - vung duoc phep sua) roi tinh pixel_diff_ratio tren PHAN CON
    LAI (phai = 0.0 de dat Profile 2 diem 4). before_pix/after_pix phai CUNG kich thuoc - lech
    kich thuoc la loi ro rang cua caller (trang bi dich chuyen giua before/after do sua noi dung
    khac lam doi so trang), khong tu resize/doan de so sanh tiep."""
    if before_pix.width != after_pix.width or before_pix.height != after_pix.height:
        raise ValueError(
            f"before_pix ({before_pix.width}x{before_pix.height}) va after_pix "
            f"({after_pix.width}x{after_pix.height}) khac kich thuoc - khong the so sanh pixel."
        )

    width, height = before_pix.width, before_pix.height
    n_before, n_after = before_pix.n, after_pix.n
    n = min(n_before, n_after)
    rois = list(authorized_rois or [])
    before_samples, after_samples = before_pix.samples, after_pix.samples
    stride_before, stride_after = width * n_before, width * n_after

    compared = 0
    diff_count = 0
    sample_diff_pixels: List[Tuple[int, int]] = []
    for y in range(height):
        row_before_start = y * stride_before
        row_after_start = y * stride_after
        for x in range(width):
            if _point_in_rois(x, y, rois):
                continue
            compared += 1
            b_off = row_before_start + x * n_before
            a_off = row_after_start + x * n_after
            if before_samples[b_off : b_off + n] != after_samples[a_off : a_off + n]:
                diff_count += 1
                if len(sample_diff_pixels) < 50:
                    sample_diff_pixels.append((x, y))

    pixel_diff_ratio = (diff_count / compared) if compared else 0.0
    return {
        "width": width,
        "height": height,
        "compared_pixels": compared,
        "diff_pixels_count": diff_count,
        "pixel_diff_ratio": pixel_diff_ratio,
        "sample_diff_pixels": sample_diff_pixels,
        "authorized_rois": rois,
    }


def require_masked_visual_diff_clean(
    before_pix: Any, after_pix: Any, authorized_rois: Optional[Sequence[Roi]] = None
) -> Dict[str, Any]:
    report = compute_masked_visual_diff(before_pix, after_pix, authorized_rois)
    if report["pixel_diff_ratio"] > 0.0:
        raise VisualDiffViolationError(
            f"VISUAL_DIFF_VIOLATION: pixel_diff_ratio={report['pixel_diff_ratio']:.6f} > 0 ngoai "
            f"{len(report['authorized_rois'])} authorized_rois "
            f"({report['diff_pixels_count']}/{report['compared_pixels']} pixel khac nhau)."
        )
    return report
