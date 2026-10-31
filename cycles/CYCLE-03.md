# CYCLE-03 — Tách `review()` khỏi `gui()` (tiếp P1, M-3)

## Pha 1 — Kế hoạch

### Vấn đề người dùng thật
`review()` là hộp thoại quan trọng nhất trước khi chạy: hiển thị số liệu dự kiến,
bảng mẫu, nút "Chấp nhận và thực hiện", "Chưa chấp nhận", "Áp dụng hàng loạt".
Nó vẫn nằm trong `gui()` (793 dòng) nên khó đọc, khó test riêng và dễ sửa nhầm
luồng token. Đây là M-3 trong ROADMAP.

### Mục tiêu
1. Chuyển thân `review()` (~100 dòng) thành hàm cấp module
   `review_plan_dialog(app, source, reference, plan, preview, export_pdf, tk, ttk, messagebox, filedialog)`.
2. Method `review()` còn một dòng gọi lại, giữ nguyên chữ ký.
3. Không đổi bố cục, nhãn, thứ tự nút, luồng `current()` → `preview_plan()` →
   so token → `perform()`, và nhánh `apply_batch()`.

### Không làm trong chu kỳ này
- Không đổi logic token/xem trước, không đổi `perform()`/`start_batch_run()`.
- Không tách `App` khỏi `gui()` (M-1) — để chu kỳ sau.
- Không đổi giao diện hay phím tắt.

### Cách làm (một commit)
- Cắt nguyên thân hàm, dedent 8 dấu cách, `self.` → `app.`, truyền
  `tk/ttk/messagebox/filedialog` làm tham số.
- Đặt hàm mới ngay trước `def gui(...)`, cạnh `open_ai_settings_dialog()`.
- Giữ nguyên mọi hằng chuỗi; đối chiếu bằng AST như Chu kỳ 2.

### Tiêu chí nghiệm thu (đo được)
- AST: danh sách hằng chuỗi của thân cũ và mới trùng khớp theo thứ tự (chỉ thêm
  docstring).
- `test_gui.py` 14/14 (có bấm `approve`/`cancel` thật trong hộp thoại review).
- `test_ui_acceptance.py` 48/48.
- `test_app.py --integration` 97/97.
- Mutation: thay `review_plan_dialog` bằng hàm rỗng → `test_gui` phải fail.

### Rủi ro
- Bỏ sót closure `filedialog` (chỉ dùng ở nhánh "Áp dụng hàng loạt") → lỗi chỉ lộ
  khi bấm nút đó; giảm thiểu bằng kiểm tra AST/tên tự do và test thủ công nhánh batch.
- Nhánh `accept()` gọi `self.current`/`self.perform`; đổi thành `app.` phải đúng
  instance, không phải class.

### Lệnh kiểm tra
```powershell
python -X utf8 -m py_compile app.py
python -X utf8 test_gui.py
python -X utf8 test_ui_acceptance.py
python -X utf8 test_app.py --integration
```

## Pha 2 — Thực hiện (2026-10-31)

- Thân `review()` (100 dòng) chuyển nguyên vẹn thành `review_plan_dialog(...)` cấp
  module, đặt cạnh `open_ai_settings_dialog()`.
- Method `review()` còn một lời gọi có tên tham số; chữ ký công khai không đổi.
- Chỉ đổi cơ học: dedent 8 dấu cách, `self.` → `app.` (11 chỗ), thêm docstring.
- `gui()` giảm từ 793 xuống **697 dòng**; hàm mới 105 dòng.

## Pha 3 — QA / đo lại

- AST: 172 hằng chuỗi của thân cũ và mới trùng khớp theo thứ tự; khác biệt duy nhất
  là docstring. Không còn tên `self` trong hàm mới.
- `test_gui.py` 14/14 — có bấm `approve` và `cancel` thật trong hộp thoại review.
- `test_ui_acceptance.py` 48/48.

## Pha 4 — Mutation

Thay `review_plan_dialog` bằng hàm rỗng rồi chạy `test_gui`: bài test **bắt được**
(`StopIteration` vì không có nút `approve`), sau đó phục hồi.

## Pha 5 — Kiểm tra hồi quy

| Lệnh | Kết quả |
|------|---------|
| `python -m py_compile app.py` | OK |
| `test_gui.py` | 14/14 |
| `test_ui_acceptance.py` | 48/48 |
| `test_app.py --integration` | 97/97 |
| `test_usecases.py` | 37/37 |
| `test_all_edge_cases.py` | 61/61 |

## Pha 6 — Phát hành

Không tăng phiên bản (refactor thuần túy). Ghi `RELEASES.md`.

## Pha 7 — Chấm lại & chu kỳ kế

- M-3 **ĐÓNG**. `gui()` còn 697 dòng; `App` vẫn lồng trong `gui()` (M-1 mở).
- Chu kỳ 4: xử lý M-4 (`infer_smart_plan` ~478 dòng) theo từng nhánh thuật toán,
  dùng `test_infer_golden` (28 prompt vàng) làm lưới an toàn; chưa chọn tách `App`
  (M-1) vì rủi ro cao hơn và chưa có test cấp class.

## Giới hạn còn lại

- Chưa tách `App` khỏi `gui()`; `start_batch_run()` và các hàm vẽ giao diện vẫn
  nằm trong `gui()`.
- Chưa đo độ phức tạp bằng công cụ ngoài (chỉ đếm dòng/AST).

