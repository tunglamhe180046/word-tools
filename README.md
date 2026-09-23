# Bộ Công cụ Chuyên Dụng Xử Lý Word (`tools/word-engine/`)

Hạ tầng kỹ thuật tầng thấp (Lower-level Infrastructure) dùng chung tại thư mục gốc của repository, chuyên trách thực hiện các thao tác kiểm tra, phẫu thuật, định dạng và bảo vệ toàn vẹn file Word (`.docx`).

Được gọi bởi AI (Jarvis/Claude) hoặc các phân hệ (`phan-tich/`, `nhan vien ho so/`, `dich-thuat/`) qua **Subprocess CLI** với hợp đồng JSON chuẩn trên `stdout`.

---

## 1. Cài đặt & Yêu cầu Môi trường

```bash
# Python >= 3.10
pip install -r tools/word-engine/requirements.txt
```

Các thư viện chính:
- `lxml`: Parse và sửa phẫu thuật OpenXML, kiểm định Canonical XML (C14N).
- `python-docx`: Đọc/ghi cấu trúc DOCX.
- `pillow`: Xử lý ảnh dấu và chữ ký.
- `psutil` & `pywin32`: Giám sát Watchdog và Word COM headless (Windows).
- `pymupdf`: Rasterize PDF và Masked Visual Diff.

---

## 2. Giao diện Dòng lệnh (CLI Usage)

Entry point duy nhất: `python tools/word-engine/cli.py <subcommand> [options]`

### 2.1. `inspect` — Quét Cấu trúc & Sinh Locators
Quét toàn bộ đoạn văn và bảng biểu trong tài liệu, tính mã băm `document_revision` và sinh các **Revision-bound Object Locators** kèm `context_sha256`:

```bash
python tools/word-engine/cli.py inspect path/to/document.docx --json
```

**Output JSON mẫu:**
```json
{
  "success": true,
  "outcome": "inspected",
  "document_revision": "sha256:e3b0c442...",
  "locators": [
    {
      "object_id": "cell_t0_r1_c2",
      "revision": "sha256:e3b0c442...",
      "kind": "table_cell",
      "table_index": 0,
      "row_index": 1,
      "col_index": 2,
      "structural_path": "/w:document/w:body/w:tbl[1]/w:tr[2]/w:tc[3]",
      "expected_text": "8.5",
      "context_sha256": "sha256:a1b2c3d4..."
    }
  ]
}
```

### 2.2. `patch-cell` — Sửa Phẫu thuật Ô Bảng
Sửa nội dung của một ô bảng duy nhất có kiểm tra khoá lạc quan (Optimistic Lock):

```bash
python tools/word-engine/cli.py patch-cell path/to/document.docx \
  --target-id "cell_t0_r1_c2" \
  --new-text "9.0" \
  --expected-revision "sha256:e3b0c442..." \
  --json
```

- Nếu hash đĩa khác `expected-revision` $\rightarrow$ Báo lỗi `DocumentDriftError` và chặn ghi đè để bảo vệ sửa tay của người dùng.
- Sau khi commit thành công, trả về `new_document_revision` và `locators` mới để chuỗi sửa tiếp theo không bị lệch toạ độ.

### 2.3. `patch-text` — Thay thế Văn bản Đa-Run (Multi-Run Resolver)
Tìm và thay thế văn bản phân tách qua nhiều run định dạng khác nhau:

```bash
python tools/word-engine/cli.py patch-text path/to/document.docx \
  --search "Hà Lội" \
  --replace "Hà Nội" \
  --json
```

- Tự động chuẩn hoá Unicode NFC.
- Fail-closed tuyệt đối (`ComplexSpanError`): Dừng ngay nếu gặp hyperlink, content controls, tracked changes, drawings hoặc breaks/tabs.
- Giữ nguyên định dạng in đậm/nghiêng (`Edit Fidelity`) hoặc báo `StyleBoundaryConflictError`.

### 2.4. `set-geometry` — Chuẩn hoá Dàn trang & Viền
Tách bạch hoàn toàn giữa Dàn trang (Pagination) và Viền (Borders):

```bash
# Chuẩn hoá khổ A4, lề công chứng (notary: 30-15-20-20mm) và chống vỡ trang (cantSplit, tblHeader)
python tools/word-engine/cli.py set-geometry path/to/document.docx \
  --page-size A4 \
  --margins notary \
  --pagination \
  --borders notary-standard \
  --json
```

### 2.5. `stamp-ops` — Quản trị Con dấu & Chữ ký
Chèn ảnh con dấu/chữ ký có kiểm tra xuất xứ pháp lý:

```bash
python tools/word-engine/cli.py stamp-ops path/to/document.docx \
  --asset-path path/to/stamp.png \
  --target-id "para_3" \
  --expected-revision "sha256:..." \
  --width-mm 35 \
  --height-mm 35 \
  --json
```

- Chỉ cho phép ảnh có mã hash SHA-256 nằm trong whitelist `.jarvis/security/stamp_allowlist.json`.
- Tự động ghi nhật ký kiểm toán vào `.jarvis/audit/stamp_provenance.jsonl`.

### 2.6. `backups` & `restore` — Lịch sử & Khôi phục An toàn
```bash
# Liệt kê các bản sao lưu
python tools/word-engine/cli.py backups path/to/document.docx --json

# Khôi phục an toàn (Mutation Transaction)
python tools/word-engine/cli.py restore path/to/document.docx \
  --backup-id "<backup_id>" \
  --json
```
- Nếu tài liệu trên đĩa có chỉnh sửa mới hơn mốc backup, hệ thống báo lỗi `RestoreConflictError` và từ chối ghi đè.

---

## 3. Sử dụng từ Node/TypeScript (`word_engine_client.ts`)

Các phân hệ Node/TypeScript (`phan-tich/`, `nhan vien ho so/`) sử dụng adapter client tại `tools/word-engine/adapters/word_engine_client.ts`:

```typescript
import { inspectDocument, patchCell } from "../tools/word-engine/adapters/word_engine_client.ts";

// 1. Quét tài liệu
const report = await inspectDocument("path/to/file.docx");
const cellLocator = report.locators.find(l => l.expected_text === "8.5");

// 2. Sửa ô
const result = await patchCell("path/to/file.docx", {
  targetId: cellLocator.object_id,
  newText: "9.0",
  expectedRevision: report.document_revision,
});
console.log("New revision:", result.document_revision);
```

---

## 4. Các Bất biến An toàn Cốt lõi (Invariants)

1. **Zero Cross-Imports:** Hoàn toàn độc lập, không import module từ bất kỳ subproject nào; ranh giới được kiểm tra tự động 2 chiều qua `scripts/check-project-boundaries.mjs`.
2. **Profile 1 (OOXML Surgical Mode):** Các OPC part không liên quan giữ nguyên raw SHA-256 100%; các subtree khác trong `document.xml` giữ nguyên C14N 100%.
3. **Commit Broker 18 bước:** Mọi thao tác ghi đều đi qua Commit Lease Lock (`.commit.lock`), tạo backup tự động, atomic replace (`os.replace`) và postcondition verification.
4. **Data Loss Guardrail:** Từ chối thao tác nếu phát hiện sửa tay trên đĩa (`DocumentDriftError` / `RestoreConflictError`).
