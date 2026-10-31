# NEXT — bàn giao sau Chu kỳ 1, mở Chu kỳ 2

## Trạng thái (2026-10-31)

- Nhánh `muse-upgrade`, **không push**. Chu kỳ 1 **ĐÓNG**: PrivacyGuard chính xác
  hơn, phát hành cục bộ **0.3.2**.
- Đã bịt 2 ca lọt tên đứng một mình; bỏ che nhầm số tiền 9 chữ số, MSSV 9 chữ số,
  mã 13 chữ số; sửa thêm lỗi ăn dấu cách sau SĐT/STK có phân cách.
- Bộ đo `test_privacy_guard_cases.py`: **80 PII + 37 non-PII**, 0 rò lọt / 0 che
  nhầm ngoài danh sách, **2 known hit** có ghi chú (tên 2 từ; tên dân tộc).
- Mutation 5 chỗ cho luật mới: cả 5 bị bắt (4/4/3 che nhầm, 14 lọt tên).
- Hồi quy: `test_security_privacy` 43/43, `test_router` 7/7,
  `test_app --integration` 97/97, `test_usecases` 37/37,
  `test_all_edge_cases` 61/61, `test_gui` 14/14.
- Tài liệu: `cycles/CYCLE-01.md` (kết quả thật), `release-notes-0.3.2.md`,
  `RELEASES.md`, `ROADMAP.md` (S-1..S-4 đóng), README cập nhật số mẫu/giới hạn.

## Việc tiếp theo (thứ tự)

1. **Pha 2 — Chu kỳ 2:** tách `open_ai_settings()` khỏi `gui()` theo
   `cycles/CYCLE-02.md`. Một commit di chuyển thuần túy, không đổi chuỗi giao diện.
2. **Pha 3/4 — Chu kỳ 2:** mở hộp thoại thật bằng Tk để xác nhận không NameError,
   không cắt nút; chạy `test_gui`, `test_ui_acceptance`, `test_app --integration`.
   Mutation nhẹ: bỏ một tham số `tk/ttk` truyền vào để chắc test bắt được.
3. **Pha 5/6 — Chu kỳ 2:** không tăng phiên bản (chỉ refactor, không đổi hành vi);
   ghi vào RELEASES mục "chu kỳ không tăng phiên bản".
4. **Pha 7:** chấm lại, cập nhật ROADMAP; chu kỳ 3 dự kiến tách `review()` (M-3).

## Quyết định đã chốt trong Chu kỳ 1

1. **Che tên đứng một mình:** đã làm theo đề xuất (từ điển họ + 3–4 từ Title Case
   + loại trừ địa danh). Hai giới hạn chấp nhận: tên 2 từ và tên dân tộc không
   thuộc họ Kinh.
2. **Phiên bản:** đã tăng **0.3.2** vì thay đổi hành vi người dùng thấy (che PII
   khác đi). Nếu chủ sản phẩm muốn gộp vào 0.4.0 cùng tính năng mới, ghi lại để
   đổi số ở lượt sau; tag `v0.3.2` là tag cục bộ, có thể xóa nếu đổi quyết định.

## Giới hạn/rủi ro đã biết

- PrivacyGuard vẫn là regex: tên 2 từ, tên dân tộc, định dạng lạ có thể lọt; tên
  đường trùng họ người ngoài danh sách loại trừ có thể bị che nhầm. Chưa đo trên
  câu người dùng thật, chưa gọi cloud thật.
- Chưa đo lại benchmark sau thay đổi PrivacyGuard (prompt ngắn, chưa có số).
- `test_headers.py` từng timeout COM một lần rồi đạt khi chạy lại; nếu lặp lại 3
  lần liền thì ghi BLOCKER.
- Guard UC-33 dựa trên owner file `~$`; file khóa sót sau crash có thể chặn nhầm.
- Ảnh UI chụp bằng GDI sau `lift()`; nếu cửa sổ bị app khác che, ảnh sẽ sai.
- Không push; kiểm tra `STOP.txt` trước mỗi chu kỳ.
