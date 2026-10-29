# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 18 — Rà soát tĩnh toàn repo + dọn dẹp nhỏ an toàn

Sau 17 vòng thay đổi, làm một vòng rà soát không đổi hành vi:
1. Chạy pyflakes (nếu có; không thì py_compile + grep thủ công) trên
   app.py/ai_client.py và các test: tìm biến không dùng, import thừa,
   hàm chết (không nơi nào gọi — đối chiếu bằng grep toàn repo, kể cả
   excel_bridge.ps1/start.cmd/README nhắc tới).
2. Chỉ gỡ những gì CHẮC CHẮN chết (0 tham chiếu). Mỗi lần gỡ chạy lại toàn
   bộ suite. Thứ không chắc: ghi vào log, không gỡ.
3. Kiểm tra tính nhất quán tài liệu lần cuối: README ↔ USECASES.md ↔ hành
   vi (đếm số test file/case ghi trong README so với thực tế).

## Hàng đợi sau đó
- Vòng 19: chuẩn bị tổng kết — đối chiếu mọi mục còn mở trong log
  (VĐ-07 phần Excel COM, VĐ-16 phần 'cập nhật', quyết định người dùng).
- Vòng 20: TỔNG KẾT 20 VÒNG theo PROMPT_MUSE (bắt buộc): chấm lại 8 hạng
  mục, việc còn dở, hàng đợi Windows, câu hỏi chờ người dùng, đề xuất
  hướng sau vòng 20 (nếu người dùng muốn tiếp tục).
- Hàng đợi Windows: UC-32/33/35/36; benchmark lại; test_gui + trực quan
  Vòng 10; nghiệm thu DPAPI thật.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định: Smart Router (có cloud) hay Ollama-only?
- '500.000' một dấu chấm = 500 nghìn hay 500 thập phân? (USECASES UC-20b)
- Nhánh `muse-upgrade`: 18 commit (Bước 0 + Vòng 1–17), CHƯA push — cần
  người dùng xác nhận mới push lên GitHub (repo công khai).
