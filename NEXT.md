# NEXT — bàn giao sau Chu kỳ 0, mở Chu kỳ 1

## Trạng thái (2026-10-31)

- Nhánh `muse-upgrade`, không push. Chu kỳ 0 **ĐÓNG**: nghiệm thu Windows đủ 3 nhóm
  test và nghiệm thu giao diện bằng ảnh chụp thật (`results/ui-shots/`).
- Đã sửa 4 lỗi bố cục phát hiện bằng ảnh (UI-01..UI-04, chi tiết trong
  `cycles/CYCLE-00.md`): hàng động cơ bị cắt, hộp Cài đặt AI cắt nút, thanh đáy cửa
  sổ chính bị cắt, chip gợi ý/checkbox bị bóp. Thêm `test_ui_acceptance.py` 48 check.
- Test sau sửa: nhóm 1 21/21 lệnh, nhóm 2 7/7 + 43/43, nhóm 3 integration 97/97,
  GUI 14/14, read-boundary 4/4. Benchmark 1.000 dòng: inspect 4,05s, lookup 6,28s,
  summary 6,13s, clean 5,31s.
- Đã lập `ROADMAP.md`, `RELEASES.md`, `cycles/CYCLE-01.md` (Pha 1 xong).

## Việc tiếp theo (thứ tự)

1. **Pha 2 — Chu kỳ 1:** sửa PrivacyGuard theo `cycles/CYCLE-01.md`:
   bịt 2 ca lọt tên đứng một mình; bỏ che nhầm số tiền 9 chữ số, mã 13 chữ số,
   MSSV 9 chữ số; mở bộ đo lên >= 80 PII + >= 30 non-PII. Kiểm tra bằng
   `python -X utf8 test_privacy_guard_cases.py` (0 rò rỉ/0 che nhầm ngoài known)
   rồi chạy `test_security_privacy.py`, `test_router.py`, `test_app.py`,
   `test_usecases.py`.
2. **Pha 3/4:** QA + thử như người dùng với bộ mẫu mở rộng; ghi kết quả vào
   `cycles/CYCLE-01.md`. Mutation 5 chỗ cho luật mới.
3. **Pha 5/6:** nếu đạt, tăng phiên bản 0.3.2 (sửa hành vi người dùng thấy: che PII
   chính xác hơn), viết `release-notes-0.3.2.md`, cập nhật README/USECASES/RELEASES,
   tag cục bộ `v0.3.2`.
4. **Pha 7:** chấm lại 8 hạng mục, cập nhật ROADMAP, mở Chu kỳ 2 (lát cắt `gui()`
   đầu tiên: `open_ai_settings`, có `test_gui` + `test_ui_acceptance` bảo vệ).

## Câu hỏi cho chủ sản phẩm

1. **Che tên riêng đứng một mình**: đề xuất dùng từ điển họ Việt + >= 3 từ Title
   Case, loại trừ địa danh. Ưu: bịt lỗ hổng PII thật. Nhược: có thể che nhầm tên
   đường/phố hiếm gặp. Đề xuất: làm, đo trên mẫu không phải PII.
2. **Phiên bản 0.3.2** cho thay đổi PrivacyGuard: đề xuất tăng (hành vi người dùng
   thấy), không đổi định dạng file xuất. Nếu bạn muốn gộp vào 0.4.0 cùng tính năng
   mới thì ghi vào NEXT khi chạy lượt sau.

## Giới hạn/rủi ro đã biết
- Ảnh UI chụp bằng GDI sau khi `lift()`; nếu cửa sổ bị app khác che, ảnh sẽ sai —
  script không còn trong repo, chỉ giữ ảnh bằng chứng.
- `test_headers.py` từng timeout COM một lần rồi đạt khi chạy lại; nếu lặp lại 3 lần
  liền thì ghi BLOCKER.
- Guard UC-33 dựa trên owner file `~$`; file khóa sót sau crash có thể chặn nhầm.
- Không push; `STOP.txt` không tồn tại tại thời điểm bàn giao (đã kiểm tra).
