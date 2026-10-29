# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 19 — Đối chiếu mọi mục còn mở, chuẩn bị tổng kết

1. Rà UPGRADE_LOG từ Bước 0: liệt kê trạng thái từng VĐ (01–17) — đã
   sửa/chấp nhận/hàng đợi Windows — kiểm tra không mục nào bị quên.
2. Đếm lại toàn bộ số liệu sẽ dùng trong tổng kết: số commit, số file test
   và tổng số check chạy được trên Linux (chạy lại cả suite một lượt và
   cộng từ output thật), số dòng app.py/ai_client.py trước (từ git bản gốc
   28f5679) và nay.
3. Soát README + USECASES.md lần cuối cho khớp số liệu vòng 19.

## Vòng 20 — TỔNG KẾT 20 VÒNG (bắt buộc theo PROMPT_MUSE)
- Viết mục "TỔNG KẾT 20 VÒNG" vào UPGRADE_LOG.md: chấm lại 8 hạng mục,
  bảng VĐ-01…17 và trạng thái, những gì CHƯA làm được (hàng đợi Windows
  nguyên vẹn), câu hỏi chờ người dùng, và đề xuất hướng tiếp theo sau
  vòng 20 nếu người dùng muốn chạy tiếp (ứng viên: nghiệm thu Windows,
  tách gui có kiểm chứng, xử lý nốt từ kích hoạt 'cập nhật').
- Cập nhật NEXT.md thành kế hoạch hậu-20 vòng.
- Commit tổng kết. Sau vòng 20, nếu không có chỉ đạo mới từ người dùng,
  vòng lặp vẫn tiếp tục theo NEXT (chế độ không cạn việc) — trừ khi có
  STOP.txt.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định: Smart Router (có cloud) hay Ollama-only?
- '500.000' một dấu chấm = 500 nghìn hay 500 thập phân? (USECASES UC-20b)
- Nhánh `muse-upgrade`: 19 commit (Bước 0 + Vòng 1–18), CHƯA push — cần
  người dùng xác nhận mới push lên GitHub (repo công khai).
