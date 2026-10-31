# CYCLE-02 — Tách `open_ai_settings()` khỏi `gui()` (bắt đầu P1)

## Pha 1 — Kế hoạch

### Vấn đề người dùng thật
Không phải lỗi người dùng thấy ngay, nhưng `gui()` đang dài ~930 dòng và chứa cả
một class `App` lồng bên trong. Mỗi lần sửa hộp Cài đặt AI (thêm nhà cung cấp,
đổi nhãn, sửa kiểm tra kết nối) đều phải cuộn trong hàm khổng lồ, dễ sửa nhầm
phần khác và khó viết test trực tiếp. Đây là nợ bảo trì đứng đầu ROADMAP (M-2).

### Mục tiêu
1. Chuyển thân `open_ai_settings()` (~160 dòng) ra hàm cấp module
   `open_ai_settings_dialog(app, tk, ttk, messagebox)`; method chỉ còn 1 dòng gọi lại.
2. Giữ nguyên 100% giao diện, nhãn, thứ tự tab, hành vi lưu/kiểm tra kết nối,
   kích thước hộp thoại theo nội dung (UI-02).
3. Không đổi `config.json`, schema, hay luồng xem trước + token.

### Không làm trong chu kỳ này
- Không tách `review()`, `start_batch_run()` hay `App` ra khỏi `gui()`.
- Không đổi bố cục/thiết kế hộp Cài đặt AI.
- Không thêm nhà cung cấp AI mới.

### Cách làm (một commit)
- Thêm hàm module-level ngay trước `def gui(...)`; import `tkinter` cục bộ trong hàm
  (giữ nguyên đặc tính import lười của app).
- Thân hàm dùng `app` thay cho `self`; các tham chiếu `tk/ttk/messagebox` lấy từ tham số.
- Trong `gui()`, method `open_ai_settings` gọi
  `open_ai_settings_dialog(self, tk, ttk, messagebox)`.
- Không đổi bất kỳ chuỗi hiển thị nào (đối chiếu bằng `git diff` chỉ có di chuyển).

### Tiêu chí nghiệm thu (đo được)
- `python -X utf8 test_gui.py` 14/14 và `test_ui_acceptance.py` 48/48.
- `test_app.py --integration` 97/97 (không đổi hành vi).
- `git diff --stat` cho thấy `app.py` chỉ thêm/xóa dòng do di chuyển; không đổi
  chuỗi tiếng Việt nào ngoài thụt lề.
- Mở hộp Cài đặt AI bằng script Tk: hộp thoại vẫn lớn hơn hoặc bằng 650x550 và
  không cắt nút "Lưu cấu hình"/"Hủy" ở 1080x800.

### Rủi ro
- Closure `tk`, `ttk`, `messagebox` bị bỏ sót → lỗi NameError chỉ lộ khi mở hộp
  thoại; giảm thiểu bằng test mở thật hộp thoại, không chỉ import.
- Di chuyển làm lệch thụt lề khối lớn; giảm thiểu bằng so sánh số dòng và chạy
  `python -m py_compile app.py`.

### Lệnh kiểm tra
```powershell
python -X utf8 -m py_compile app.py
python -X utf8 test_gui.py
python -X utf8 test_ui_acceptance.py
python -X utf8 test_app.py --integration
```

## Pha 2 — Thực hiện (2026-10-31)

- Thân `open_ai_settings()` (163 dòng) đã chuyển nguyên vẹn thành
  `open_ai_settings_dialog(app, tk, ttk, messagebox)` đặt ngay trước `def gui(...)`.
- Method trong `App` còn 2 dòng:
  `open_ai_settings_dialog(self, tk, ttk, messagebox)`.
- Chỉ đổi ba thứ cơ học: bỏ 8 dấu cách thụt lề, `self.` → `app.` (5 chỗ), thêm
  docstring mô tả. Không đổi chuỗi hiển thị, nhãn, bố cục, logic lưu/kiểm tra.
- `gui()` giảm từ ~956 xuống **793 dòng**; hàm mới 167 dòng.

## Pha 3 — QA / đo lại

- So sánh bằng AST: 141 hằng chuỗi trong thân hàm cũ và mới **trùng khớp hoàn toàn
  theo thứ tự**; khác biệt duy nhất là docstring mới.
- `test_ui_acceptance.py` mở hộp Cài đặt AI thật qua wrapper và kiểm tra kích thước
  theo nội dung + nút không bị bóp: **48/48**.
- `test_gui.py` **14/14**.

## Pha 4 — Mutation

Thay `open_ai_settings_dialog` bằng hàm rỗng (hộp thoại không mở) rồi chạy lại
`test_ui_acceptance`: bài test **bắt được** (`StopIteration` khi không tìm thấy
Toplevel), rồi phục hồi. Nghĩa là lưới an toàn thực sự bao phủ phần vừa tách,
không chỉ import thành công.

## Pha 5 — Kiểm tra hồi quy

| Lệnh | Kết quả |
|------|---------|
| `python -m py_compile app.py` | OK |
| `test_gui.py` | 14/14 |
| `test_ui_acceptance.py` | 48/48 (mở hộp thoại thật) |
| `test_app.py --integration` | 97/97 |
| `test_usecases.py` | 37/37 |
| `test_all_edge_cases.py` | 61/61 |
| `test_privacy_guard_cases.py` | 80 PII / 37 non-PII, 2 known (không đổi) |

## Pha 6 — Phát hành

Không tăng phiên bản: đây là refactor thuần túy, không đổi hành vi người dùng.
Ghi vào `RELEASES.md` mục "Chu kỳ không tăng phiên bản".

## Pha 7 — Chấm lại & chu kỳ kế

- M-2 **ĐÓNG**. `gui()` vẫn còn 793 dòng và class `App` vẫn lồng trong hàm; M-1
  chưa đóng.
- Chu kỳ 3: tách `review()` (M-3, ~100 dòng) theo cùng cách, dùng
  `test_gui`/`test_ui_acceptance` làm lưới an toàn; sau đó mới cân nhắc M-1.

## Giới hạn còn lại

- Chưa tách `App` ra khỏi `gui()`; chưa tách `review()`/`start_batch_run()`.
- Chưa đo lại độ dài/độ phức tạp bằng công cụ ngoài (chỉ đếm dòng bằng AST).
