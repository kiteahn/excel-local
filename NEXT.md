# NEXT — bàn giao sau Chu kỳ 2, mở Chu kỳ 3

## Trạng thái (2026-10-31)

- Nhánh `muse-upgrade`, **không push**. Chu kỳ 1 **ĐÓNG** (PrivacyGuard 0.3.2, tag
  cục bộ `v0.3.2`), Chu kỳ 2 **ĐÓNG** (refactor `open_ai_settings`, không tăng phiên bản).
- `open_ai_settings()` đã tách thành `open_ai_settings_dialog()` cấp module;
  `gui()` giảm từ ~956 xuống **793 dòng**. 141 hằng chuỗi giữ nguyên thứ tự; chỉ
  thêm docstring. Mutation "hộp thoại không mở" bị `test_ui_acceptance` bắt.
- Hồi quy sau refactor: `test_gui` 14/14, `test_ui_acceptance` 48/48,
  `test_app --integration` 97/97, `test_usecases` 37/37,
  `test_all_edge_cases` 61/61, PrivacyGuard 80 PII/37 non-PII (2 known).
- Tài liệu: `cycles/CYCLE-02.md` (kết quả thật), `cycles/CYCLE-03.md` (kế hoạch),
  `ROADMAP.md` (M-2 đóng), `RELEASES.md` (chu kỳ không tăng phiên bản).

## Việc tiếp theo (thứ tự)

1. **Pha 2 — Chu kỳ 3:** tách `review()` thành `review_plan_dialog()` theo
   `cycles/CYCLE-03.md`. Một commit di chuyển thuần túy, không đổi chuỗi.
2. **Pha 3/4:** so AST hằng chuỗi; chạy `test_gui` (có bấm approve/cancel thật),
   `test_ui_acceptance`, `test_app --integration`; mutation thay hàm bằng rỗng phải
   bị `test_gui` bắt.
3. **Pha 5/6:** không tăng phiên bản (refactor); ghi RELEASES.
4. **Pha 7:** chấm lại; chu kỳ 4 cân nhắc M-4 (`infer_smart_plan` ~478 dòng) hoặc
   M-1 (tách `App` khỏi `gui()`) — chọn theo mức rủi ro đo được.

## Quyết định đã chốt

- Chu kỳ 1 tăng **0.3.2** (hành vi người dùng thấy); tag `v0.3.2` là tag cục bộ,
  chưa push. Nếu muốn gộp vào 0.4.0, xóa tag và đổi số ở lượt sau.
- Refactor không tăng phiên bản; mỗi lát cắt phải có test bảo vệ + mutation chứng
  minh test bắt được lỗi.

## Giới hạn/rủi ro đã biết

- PrivacyGuard vẫn là regex: tên 2 từ, tên dân tộc, định dạng lạ có thể lọt; tên
  đường trùng họ người ngoài danh sách loại trừ có thể bị che nhầm. Chưa đo trên
  câu người dùng thật, chưa gọi cloud thật.
- Chưa đo lại benchmark sau thay đổi PrivacyGuard (prompt ngắn, chưa có số).
- `gui()` vẫn còn 793 dòng, `App` vẫn lồng trong `gui()`.
- `test_headers.py` từng timeout COM một lần rồi đạt khi chạy lại; nếu lặp lại 3
  lần liền thì ghi BLOCKER.
- Guard UC-33 dựa trên owner file `~$`; file khóa sót sau crash có thể chặn nhầm.
- Ảnh UI chụp bằng GDI sau `lift()`; nếu cửa sổ bị app khác che, ảnh sẽ sai.
- Không push; kiểm tra `STOP.txt` trước mỗi chu kỳ.
