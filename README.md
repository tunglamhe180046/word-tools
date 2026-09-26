# Document Tools (Office Surgical Engine Suite) 🚀

> **Hạ tầng phẫu thuật tài liệu Word (.docx) & bảng tính Excel (.xlsx) tại chỗ với độ chính xác cao dành cho AI Agent & Hệ thống tự động hóa.**  
> *Chỉnh sửa trực tiếp trên file Microsoft Office hiện có mà không cần tạo lại từ đầu, bảo tồn nguyên vẹn 100% nội dung sửa tay của con người, định dạng phức tạp, công thức, hình vẽ, và cấu trúc gói tệp OPC.*

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Phiên bản Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![Word Engine](https://img.shields.io/badge/Word%20Engine-Active-blue.svg)](word-engine/README.md)
[![Excel Engine](https://img.shields.io/badge/Excel%20Engine-Active-green.svg)](excel-engine/README.md)
[![Kiến trúc: Phẫu thuật tại chỗ](https://img.shields.io/badge/Kiến%20trúc-Phẫu%20thuật%20tại%20chỗ-success.svg)](#-cấu-trúc-hệ-thống-thư-mục)

---

## 🌟 1. Tổng quan Kiến trúc Bộ công cụ

Bộ công cụ được cấu trúc thành **2 Engine phẫu thuật độc lập** cùng một **Unified CLI Dispatcher** thông minh tại thư mục gốc:

```text
word-tools/
├── cli.py                     # Unified CLI Dispatcher (tự động nhận diện định dạng file)
├── LICENSE                    # MIT License
├── README.md                  # Tài liệu tổng quan hệ thống
│
├── word-engine/               # 📝 Docx Surgical Engine (chuyên xử lý file Word .docx)
│   ├── cli.py                 # CLI chuyên trách Word
│   ├── requirements.txt       # lxml, python-docx, pillow, pymupdf, psutil
│   ├── README.md              # Hướng dẫn chi tiết & đặc tả kỹ thuật Word Engine
│   ├── core/                  # Các module phẫu thuật Word (patch-cell, patch-text, stamp-ops, geometry)
│   ├── adapters/              # TypeScript client adapter cho Node/Next.js
│   └── tests/                 # Bộ kiểm thử 123 tests độc lập cho Word
│
└── excel-engine/              # 📊 Xlsx Surgical Engine (chuyên xử lý file Excel .xlsx)
    ├── cli.py                 # CLI chuyên trách Excel
    ├── pytest.ini             # Cách ly hoàn toàn bộ đệm kiểm thử
    ├── requirements.txt       # lxml, openpyxl (dev/test)
    ├── README.md              # Hướng dẫn chi tiết & đặc tả kỹ thuật Excel Engine
    ├── core/                  # Các module phẫu thuật Excel (inlineStr, cell_addressing, calcChain)
    └── tests/                 # Bộ kiểm thử 151 tests độc lập cho Excel (100% tmp_path)
```

---

## 💻 2. Giao diện Dòng lệnh Hợp nhất (`cli.py`)

Tại thư mục gốc, bạn có thể gọi trực tiếp `cli.py`. Hệ thống sẽ **tự động phát hiện định dạng file** (`.docx` hoặc `.xlsx`) để điều hướng lệnh đến đúng engine:

```bash
# Tự động điều phối theo phần mở rộng:
python cli.py inspect "tai_lieu.docx" --json      # Tự gọi word-engine
python cli.py inspect "bang_tinh.xlsx" --json     # Tự gọi excel-engine

# Hoặc chỉ định tường minh engine muốn gọi:
python cli.py word inspect "tai_lieu.docx" --json
python cli.py excel inspect "bang_tinh.xlsx" --json
```

---

## 📦 3. So sánh Hai Engine Chuyên biệt

| Tính năng cốt lõi | Word Engine (`word-engine/`) | Excel Engine (`excel-engine/`) |
| :--- | :--- | :--- |
| **Định dạng đích** | Microsoft Word (`.docx`) | Microsoft Excel (`.xlsx`) |
| **Cấu trúc OpenXML** | `WordprocessingML` (`w:p`, `w:r`, `w:tbl`, `w:tc`) | `SpreadsheetML` (`c`, `v`, `f`, `inlineStr`, `row`) |
| **Bảo tồn Byte** | Giữ nguyên 100% SHA-256 các part ngoài body | Giữ nguyên 100% SHA-256 của `sharedStrings.xml`, styles, vba |
| **Lưới an toàn** | `document_revision` SHA-256, `DocumentDriftError` | `document_revision` SHA-256, `DocumentDriftError` |
| **Quản lý Khóa** | Lock lạc quan + Lease Lock độc quyền | Lock lạc quan + Atomic replace giao dịch |
| **Phẫu thuật Ô** | Vá ô bảng, giữ nguyên viền, padding, màu nền | Vá ô bảng tính, ghi `inlineStr`, giữ nguyên cell style `s` |
| **Sao lưu & Phục hồi**| Tự động tạo snapshot trong `_backup/` | Tự động tạo snapshot, kiểm tra mã băm khi phục hồi |
| **Tính năng đặc thù** | Căn lề A4, chống tràn bảng `cantSplit`, chèn dấu | Bật `fullCalcOnLoad`, xử lý `calcChain` khi xóa công thức |
| **Độ phủ Kiểm thử** | **123 tests passed** | **151 tests passed** (100% cô lập trong `tmp_path`) |

---

## 🛠️ 4. Hướng dẫn Chi tiết từng Engine

- Xem hướng dẫn chi tiết của **Word Engine**: [`word-engine/README.md`](word-engine/README.md)
- Xem hướng dẫn chi tiết của **Excel Engine**: [`excel-engine/README.md`](excel-engine/README.md)

---

## 📄 5. Bản quyền & Giấy phép

Dự án được phát hành theo giấy phép nguồn mở [MIT License](LICENSE) — Bản quyền (c) 2026 Tùng Lâm.
