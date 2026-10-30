# NEXT — bàn giao Chu kỳ 0

## Trạng thái (2026-10-30)

- Làm việc trên checkout cô lập nhánh `muse-upgrade`, commit nền `a9a1c5a`.
- Đã sửa UC-33: `inspect_file` chặn workbook `.xlsx/.xlsm` có owner marker `~$` với thông báo tiếng Việt rõ ràng. Regression được thêm vào `test_read_boundary.py`.
- Sau sửa: nhóm 1 20/20 lệnh, 437 check có bộ đếm và 6 kịch bản; nhóm 2 50/50; nhóm 3 đã rerun integration 97/97, GUI 14/14, read-boundary 4/4. Benchmark COM sau sửa chưa rerun theo yêu cầu dừng tại handoff; baseline trước sửa ở `results/benchmark-current-1000.json`.
- Kết quả/đầu ra nguyên văn: `cycles/CYCLE-00.md`. UC-33 probe mở workbook synthetic bằng Excel thật và xác nhận app chặn đọc.
- Chu kỳ 0 chưa đóng vì còn xác nhận UI trực tiếp. Không qua Chu kỳ 1 trước khi xác nhận các mục bên dưới.

## Bước bàn giao

1. Khi Computer Use chụp đúng cửa sổ app, hoàn tất UC-32 trong app, xem biên bản/sheets; W-1 đã kiểm riêng, không có `Sheet1` mặc định.
2. Xem thông báo UI cho UC-33 và UC-35; lõi kiểm tra UC-33 và từ chối công thức lỗi đã qua test.
3. Xác nhận tooltip Ollama/Preview, `Ctrl+Enter` đề xuất, và hộp lỗi Batch với file hỏng.
4. PDF integration đã render/xem: không cắt/chồng, mẫu ngắn nên có khoảng trắng lớn. Có thể xác nhận thêm luồng bật PDF trong app.
5. Sau UI check, chạy benchmark COM sau sửa nếu còn thời gian, ghi kết quả; rồi mới đóng Chu kỳ 0 và bắt đầu Chu kỳ 1.

## Giới hạn/rủi ro đã biết

- Computer Use trả screenshot không khớp target app (hiển thị cửa sổ Codex/Chrome), nên chưa gửi click/typing. Không dùng ảnh đó làm bằng chứng UI.
- Guard UC-33 dựa vào owner file Excel `~$`; file khóa cũ còn sót sau crash có thể chặn đọc nhầm. Nếu gặp, xác nhận Excel đã đóng trước khi xóa file khóa.
- Tại thời điểm ghi handoff không còn Excel process do test. Không đóng hoặc tương tác với Excel của người dùng.
- Không push nhánh; việc push thuộc chủ sản phẩm. Nếu `STOP.txt` xuất hiện, dừng theo prompt.
