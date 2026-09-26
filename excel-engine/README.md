# Excel Tools (XLSX Surgical Engine)

Runtime/CLI chuyên dụng để **đọc và sửa phẫu thuật từng ô** trong workbook `.xlsx`, thuộc hạ tầng kỹ thuật tầng thấp dùng chung ở gốc repo (cùng nhóm với `tools/word-engine/`). Bộ công cụ được AI hoặc các phân hệ gọi qua subprocess CLI với hợp đồng JSON chuẩn trên stdout; **không import trực tiếp** chéo ngôn ngữ hay chéo dự án con.

Nguyên tắc xuyên suốt: sửa đúng ô được yêu cầu, giữ nguyên mọi phần còn lại của workbook (kiểu dáng, gộp ô, Table, công thức, ảnh, part phụ) và không bao giờ âm thầm ghi đè chỉnh sửa tay của người dùng.

---

## 1. Tại sao cần Excel Tools?

Ghi lại cả workbook bằng thư viện bảng tính thông thường sẽ làm mất hoặc biến dạng những phần thư viện không hiểu (định dạng có điều kiện, part tuỳ biến, mã định danh Table, công thức dùng chung...). Engine này chỉ chỉnh sửa XML của đúng part worksheet cần thiết và đóng gói lại ZIP theo từng entry:

- Mọi part không bị sửa giữ nguyên **100% byte** (SHA-256 không đổi), đúng thứ tự entry ban đầu.
- Mọi lần ghi đi qua Commit Broker: khoá, sao lưu, thay thế nguyên tử (atomic), kiểm tra hậu điều kiện.
- Dữ liệu đọc ra là JSON sạch, có phân biệt giá trị đã lưu (`cached_value`) và giá trị cần tính lại (`stale`).

## 2. Bảng lệnh

| Lệnh | Chức năng | Ghi đĩa |
| --- | --- | --- |
| `inspect` | Cấu trúc workbook: sheet, dimension, ô gộp, Table, vùng array formula, locator từng ô, `document_revision` | Không |
| `read-sheet` | Đọc các ô có trong XML của một sheet (tuỳ chọn giới hạn vùng) | Không |
| `read-cell` | Đọc một ô | Không |
| `patch-cell` | Sửa một ô (chuỗi, số, bool, công thức, xoá giá trị) | Có (qua Broker) |
| `patch-range` | Ghi ma trận 2D vào một vùng, một lô nguyên tử | Có (qua Broker) |
| `batch-patch` | Nhiều thao tác trong một lô nguyên tử, commit đúng một lần | Có (qua Broker) |
| `backups` | Liệt kê các bản sao lưu trong `.jarvis/backups/` | Không |
| `restore` | Khôi phục từ một bản sao lưu (có Optimistic Lock) | Có (qua Broker) |

## 3. Cài đặt và yêu cầu môi trường

- Python 3.10 trở lên (đã kiểm thử trên 3.13).
- Thư viện: `lxml>=5.0.0` (xem `requirements.txt`). Không cần Excel/COM: engine làm việc thẳng trên XML nên chạy được trên mọi hệ điều hành.

```bash
cd tools/excel-engine
pip install -r requirements.txt
```

## 4. Hướng dẫn sử dụng CLI

Mọi lệnh nhận `--json`. Khi có `--json`, stdout chỉ chứa **đúng một dòng JSON**:

- Thành công (exit 0): `{"success": true, "outcome": "...", ...}`
- Lỗi (exit 1): `{"success": false, "error": "<TênException>", "reason": "<thông điệp>"}`. Lỗi cú pháp tham số cũng tuân theo hợp đồng này (`"error": "ArgumentError"`).

Mặc định `work_dir` là thư mục chứa file (thư mục `.jarvis/` được neo ngay cạnh hồ sơ đang sửa), `allowed_roots` là thư mục chứa file, `actor` là `excel-engine-cli`.

### 4.1. Locator và quy ước địa chỉ

- Ô địa chỉ theo A1 (`B3`); `row`, `col` đều **1-indexed**; sheet chọn theo tên hoặc chỉ số 1-indexed theo thứ tự trong `xl/workbook.xml`.
- Locator ô: `cell_s{sheet}_r{row}_c{col}`, ví dụ `cell_s1_r3_c2` là ô B3 của sheet thứ nhất.
- `document_revision` là `sha256:<hex>` của toàn bộ file.

### 4.2. `inspect`

```bash
python cli.py inspect book.xlsx --json [--limit 5000]
```

Trả về `document_revision`, `sheets[]` (`index`, `name`, `state`, `dimension`, `merged_cells`, `tables`, `array_formulas`, `cell_count`, `formula_count`), `calc` (`full_calc_on_load`, `has_calc_chain`) và `locators[]` (bị cắt ở `--limit`, khi đó `locators_truncated` là `true`).

### 4.3. `read-sheet` và `read-cell`

```bash
python cli.py read-sheet book.xlsx --sheet Data --range A1:C10 --json
python cli.py read-cell  book.xlsx --sheet Data --cell B4 --json
```

Mỗi ô có `ref`, `locator`, `type` (`string|number|bool|error|date|blank`), `value`, `has_formula`, `formula`, `formula_kind` (`normal|shared|array|dataTable`), `cached_value`, `stale`, `style`, `merge`.

Quy ước ô công thức: `cached_value` là giá trị Excel đã lưu trong `<v>`. Ô công thức mới hoặc chưa có cache trả về `cached_value: null` và `stale: true`. Khi workbook đã đặt `fullCalcOnLoad="1"` (sau mọi lần patch), mọi ô công thức đều được đánh dấu `stale: true` vì giá trị lưu có thể đã cũ. **Engine không tự tính công thức**; Excel sẽ tính lại khi mở file. Ngày tháng không được diễn giải theo kiểu dáng: giá trị số được trả nguyên. Chuỗi đọc ra đã giải mã `_xHHHH_` nên khớp đúng chuỗi đã ghi.

### 4.4. `patch-cell`

```bash
python cli.py patch-cell book.xlsx --sheet Data --cell A2 --value "Nguyễn Văn A" \
    --expected-revision sha256:<hex> --json
python cli.py patch-cell book.xlsx --sheet Data --cell B3 --value 12.5 --type number --expected-revision ... --json
python cli.py patch-cell book.xlsx --sheet Data --cell C3 --value "=B3*2" --type formula --expected-revision ... --json
python cli.py patch-cell book.xlsx --target-id cell_s1_r2_c1 --type blank --expected-revision ... --json
```

`--type` gồm `text` (mặc định), `number`, `bool` (`true|false`), `formula`, `blank`. Chuỗi bắt đầu bằng `=` với `--type text` vẫn là **văn bản**, không bị hiểu thành công thức. `--expected-revision` là **bắt buộc**. `--truncate` cắt chuỗi quá dài thay vì từ chối.

### 4.5. `patch-range` và `batch-patch`

```bash
python cli.py patch-range book.xlsx --sheet Data --range A10:C11 \
    --values-json '[["a",1,true],[null,2.5,{"formula":"SUM(B10:B11)"}]]' --expected-revision ... --json

python cli.py batch-patch book.xlsx --expected-revision ... --json \
    --ops-json '[{"sheet":"Data","cell":"A2","value":"x"},{"sheet":"Second","cell":"C1","formula":"=A1*2"}]'
python cli.py batch-patch book.xlsx --ops-file ops.json --expected-revision ... --json
```

Quy tắc suy ra kiểu từ JSON: chuỗi là text, số là number, `true/false` là bool, `null` là ô trống, `{"formula": "..."}` là công thức. Mỗi op trong lô có thể dùng `target_id` thay cho `sheet` + `cell`. Toàn bộ lô được kiểm tra trong bộ nhớ: **một op lỗi thì huỷ cả lô, không ghi gì**; nếu hợp lệ thì commit đúng một lần. Ghi hai lần vào cùng một ô trong một lô bị từ chối.

### 4.6. `backups` và `restore`

```bash
python cli.py backups book.xlsx --json
python cli.py restore book.xlsx --backup-id <id> --json
```

`restore` từ chối (`RestoreConflictError`) nếu file trên đĩa đã bị sửa tay kể từ lần cuối engine ghi nhận, và file chưa từng đi qua engine cũng bị coi là chưa có baseline. Khi hợp lệ, bản hiện tại được sao lưu trước rồi mới ghi bản phục hồi.

## 5. Quy tắc ghi và các lỗi fail-closed

| Quy tắc | Hành vi |
| --- | --- |
| Chuỗi | Ghi dạng `<c t="inlineStr"><is><t xml:space="preserve">…</t></is></c>`; không đụng `sharedStrings.xml` |
| Escape | XML do lxml xử lý; chuỗi literal dạng `_xHHHH_` được escape thành `_x005F_xHHHH_` |
| Ký tự điều khiển | Ký tự ASCII không hợp lệ trong XML 1.0 (`\x00-\x08`, `\x0B`, `\x0C`, `\x0E-\x1F`, `U+FFFE/FFFF`, surrogate lẻ) bị từ chối |
| Độ dài | Chuỗi trên 32767 đơn vị UTF-16 bị từ chối (`TextTooLongError`) hoặc cắt an toàn khi có `--truncate` |
| Thứ tự phần tử | Trong `<c>`: `<f>` trước, rồi `<v>`/`<is>`; `<c>` tăng dần theo cột, `<row>` tăng dần theo dòng; dòng/ô chưa có được chèn đúng vị trí; `<dimension>` được mở rộng |
| Kiểu dáng | Thuộc tính `s="..."` của ô được giữ nguyên; ô mới thừa hưởng `s` của dòng nếu dòng có `customFormat` |
| Ô gộp | Chỉ sửa được ô anchor góc trên trái; ô còn lại trong vùng gộp bị từ chối (`MergedCellError`) |
| Table | Header của Table bị từ chối (`TableHeaderError`) |
| Array formula | Ô trong vùng array formula hoặc data table bị từ chối (`ArrayFormulaError`) |
| Shared formula | Không ghi đè ô gốc của shared formula (`SharedFormulaError`); ô con sửa được |
| Tính lại | Mọi lần patch đặt `fullCalcOnLoad="1"` trong `<calcPr>`; khi sửa/xoá ô công thức hoặc ghi công thức mới, `xl/calcChain.xml` bị gỡ cùng entry trong `[Content_Types].xml` và Relationship trong `xl/_rels/workbook.xml.rels` để Excel không báo repair |
| Định dạng file | Đọc: `.xlsx`/`.xlsm`; **ghi: chỉ `.xlsx`** (đuôi khác bị `UnsupportedWorkbookError`). Chỉ nhận ZIP thật (magic `PK\x03\x04`); từ chối file mã hoá/OLE2, `.xls`, `.xlsb`, entry ZIP mã hoá, tên entry thoát thư mục, ZIP-bomb |
| Sheet không phải worksheet | Chartsheet/dialogsheet hiện trong `inspect` với `kind` tương ứng và `cell_count: 0`; đọc hoặc ghi ô của chúng bị từ chối |
| Sao lưu | `restore` so SHA-256 của file backup với meta trước khi ghi; backup bị sửa thì `BackupCorruptedError` |
| XML | Parse với XXE tắt hoàn toàn; mọi tài liệu có DOCTYPE/ENTITY bị từ chối (`UnsafeXmlError`) |

Giới hạn đã biết: engine không tự tính công thức, không diễn giải ngày theo kiểu dáng, không cập nhật `sharedStrings.xml` (chuỗi cũ không còn tham chiếu vẫn nằm trong file), và ô mới trong cột có style cột (`<cols>`) không tự thừa hưởng style cột.

## 6. Kiến trúc lõi và lưới an toàn

```
cli.py  ->  core/inspector.py           (đọc, locator, Reader)
        ->  core/surgical_patcher.py    (kiểm tra + sửa XML trong bộ nhớ)
                 |-> core/xlsx_package.py    (mở/đóng gói ZIP theo entry, hash part)
                 |-> core/cell_addressing.py (A1 <-> row/col, locator)
                 `-> core/safety_gateway.py  (Commit Broker 18 bước, drift, XXE)
                         `-> core/canonical_path.py
```

- **Drift detection**: `--expected-revision` được so với SHA-256 của file ngay trước khi sửa, và lần nữa với baseline của Job Ticket ngay sau khi phát hành ticket. Khác nhau thì `DocumentDriftError` và không ghi gì.
- **Optimistic Lock**: Bước 8 của Broker so hash file với baseline lúc phát hành job ngay dưới khoá ghi; người dùng lưu file giữa chừng thì `DocumentModifiedError`, file của người dùng được giữ nguyên.
- **Commit Broker** (`safety_gateway.py`): bản sao độc lập từ `tools/word-engine/core/safety_gateway.py` tại commit `4a704a8` (Job Ticket ký HMAC, Dedicated Commit Lease Lock `.commit.lock`, khoá `.write.lock`, phát hiện owner-file `~$...` của Excel, sao lưu tập trung `.jarvis/backups/<khoá>/`, thay thế nguyên tử một bước, kiểm tra hậu điều kiện, audit log). Các điểm khác biệt có chủ đích được ghi ở docstring đầu file.
- **Không để lại file rác**: thư mục job tạm `.jarvis/work/<job_id>/` được dọn sau mỗi lần commit, kể cả khi thất bại (trừ ca `PostconditionVerificationError` cần khôi phục thủ công).

## 7. Tích hợp từ TypeScript / Node.js

```typescript
import { execFile } from "node:child_process";
import { promisify } from "node:util";
const run = promisify(execFile);

async function excel(args: string[]) {
  const { stdout } = await run("python", ["tools/excel-engine/cli.py", ...args, "--json"]);
  return JSON.parse(stdout); // đúng một dòng JSON
}
```

Khi exit code là 1, stdout vẫn chứa một dòng JSON `{"success": false, ...}`; `execFile` sẽ ném lỗi kèm `stdout` trong đối tượng lỗi.

## 8. Kiểm thử

Toàn bộ test chạy trong `tmp_path`, không tạo file rác trên đĩa (`.jarvis/`, sao lưu, khoá đều nằm trong thư mục tạm; `test_isolation.py` xác minh điều này).

```bash
# từ gốc repo
PYTHONDONTWRITEBYTECODE=1 python -m pytest tools/excel-engine/tests -v
```

Bộ test gồm: `test_cell_addressing`, `test_inspector`, `test_patch_cell`, `test_patch_batch`, `test_drift_and_lock`, `test_backup_restore`, `test_preservation` (SHA-256 các part không bị sửa giữ nguyên 100%), `test_isolation`, `test_cli`. Test `test_patched_workbook_opens_in_openpyxl` là kiểm chứng phụ, tự bỏ qua khi máy không có `openpyxl` (thư viện này không phải phụ thuộc của engine).
