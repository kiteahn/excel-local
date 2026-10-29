# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 1 — Sửa README theo hành vi thật (VĐ-01, VĐ-02, VĐ-08)

Vì sao: tài liệu đang nói sai về việc dữ liệu có đi cloud hay không — đây là vấn
đề nghiêm trọng nhất vì liên quan tới kỳ vọng quyền riêng tư của người dùng.

Làm:
1. Đọc `ai_client.request_plan` và `get_active_candidates` để xác nhận chuỗi
   provider thật cho từng lựa chọn (smart / từng cloud / ollama).
2. Sửa README: phân biệt rõ 2 chế độ — Ollama local (chỉ loopback) và Smart
   Router/cloud (tên cột + yêu cầu đã khử PII được gửi tới nhà cung cấp đã chọn).
   Bỏ "Zero-Leakage"/"tuyệt đối", thay bằng mô tả có giới hạn.
3. Sửa: v0.3.0 → v0.3.1 ở đoạn cuối; "ba thao tác" → bốn; test_cli 4 → 6;
   bước 1 bỏ XLS hoặc ghi rõ bị từ chối theo mục Giới hạn.
4. Kiểm tra: `grep -n "tuyệt đối\|Zero-Leakage\|v0.3.0\|ba thao tác" README.md`
   phải sạch; chạy lại test_app.py (79) để chắc chắn không đụng code.
5. Commit, cập nhật UPGRADE_LOG.md, chuyển NEXT sang Vòng 2 (đường dẫn cứng).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn; tài liệu sẽ mô tả đúng như code, không tự đổi
  chính sách. (Xem mục "DỪNG VÀ HỎI TÔI" trong PROMPT_MUSE.md.)
