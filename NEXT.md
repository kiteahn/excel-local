# NEXT — bàn giao sau Chu kỳ 3, mở Chu kỳ 4

## Trạng thái (2026-10-31)

- Nhánh `muse-upgrade`, **không push**. Chu kỳ 1 **ĐÓNG** (PrivacyGuard 0.3.2, tag
  cục bộ `v0.3.2`), Chu kỳ 2 **ĐÓNG** (`open_ai_settings`), Chu kỳ 3 **ĐÓNG**
  (`review`).
- `open_ai_settings_dialog()` 167 dòng và `review_plan_dialog()` 105 dòng đã ở cấp
  module; `gui()` giảm 956 → 793 → **697 dòng**. Hằng chuỗi giữ nguyên thứ tự
  (141 và 172), chỉ thêm docstring.
- Mutation "hộp thoại không mở" bị `test_ui_acceptance`/`test_gui` bắt ở cả hai lát.
- Hồi quy sau Chu kỳ 3: `test_gui` 14/14, `test_ui_acceptance` 48/48,
  `test_app --integration` 97/97, `test_usecases` 37/37,
  `test_all_edge_cases` 61/61, PrivacyGuard 80 PII/37 non-PII (2 known).
- Đã đo lại hotspot bằng AST và **sửa ROADMAP**: `infer_smart_plan` chỉ còn 25
  dòng; hàm lớn thật là `prepare_sheets` 262, `App.__init__` 213,
  `_lookup_plan` 180, `_filter_plan` 178.

## Việc tiếp theo (thứ tự)

1. **Pha 2 — Chu kỳ 4:** tách nhánh `lookup` khỏi `prepare_sheets` theo
   `cycles/CYCLE-04.md`; một commit, không đổi công thức/nhãn.
2. **Pha 3/4:** so AST hằng chuỗi; chạy `test_app --integration` (lookup thật),
   `test_usecases` (UC-11..16), `test_gui`, `test_ui_acceptance`; mutation đổi nhãn
   `'Gia_tri'` phải bị bắt.
3. **Pha 5/6:** không tăng phiên bản; ghi RELEASES.
4. **Pha 7:** chấm lại; chu kỳ 5 tách nhánh `summary` (hoặc `filter`), tiếp theo
   `clean`; sau đó mới tới `App.__init__`/M-1.

## Quyết định đã chốt

- 0.3.2 là bản hành vi gần nhất; các refactor giữ nguyên số.
- Mỗi lát cắt phải có (a) so AST hằng chuỗi, (b) test thật, (c) mutation chứng minh
  lưới an toàn bắt được lỗi.

## Giới hạn/rủi ro đã biết

- PrivacyGuard vẫn là regex: tên 2 từ, tên dân tộc, định dạng lạ có thể lọt; tên
  đường trùng họ người ngoài danh sách loại trừ có thể bị che nhầm. Chưa đo trên
  câu người dùng thật, chưa gọi cloud thật.
- Chưa đo lại benchmark sau thay đổi PrivacyGuard (prompt ngắn, chưa có số).
- `gui()` còn 697 dòng; `App` vẫn lồng trong `gui()`, `start_batch_run()` 90 dòng
  vẫn ở trong class.
- `test_headers.py` từng timeout COM một lần rồi đạt khi chạy lại; nếu lặp lại 3
  lần liền thì ghi BLOCKER.
- Guard UC-33 dựa trên owner file `~$`; file khóa sót sau crash có thể chặn nhầm.
- Ảnh UI chụp bằng GDI sau `lift()`; nếu cửa sổ bị app khác che, ảnh sẽ sai.
- Không push; kiểm tra `STOP.txt` trước mỗi chu kỳ.
