# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 2 — Bỏ đường dẫn cứng (VĐ-03, mục B của PROMPT)

Làm:
1. `test_router.py:30`: bỏ `Path(r'd:\Excel\excel-local\config.json')`. Test phải
   hermetic: tự tạo config.json tạm trong thư mục tạm (hoặc monkeypatch
   `ai_client.CONFIG_FILE`) với khóa giả, kiểm tra vòng đời mã hóa DPAPI/đọc-ghi
   trên đó; trên máy không có DPAPI thì phần DPAPI skip rõ ràng (liên quan VĐ-04).
2. `generate_project_pdf.py:7`: `OUTPUT_DIR` đổi sang đường dẫn theo `__file__`
   (`Path(__file__).parent / 'output'`). Đường dẫn Edge/Excel trong nội dung
   HTML báo cáo (dòng ~613, ~1055) là dữ kiện môi trường máy đích — giữ nhưng
   cho phép ghi đè bằng biến môi trường nếu đơn giản.
3. Kiểm tra: chạy test_router.py trên Linux — phần không cần DPAPI phải chạy,
   phần DPAPI skip có thông báo; chạy lại test_app.py (79) và
   test_all_edge_cases.py (61).
4. Commit, cập nhật UPGRADE_LOG.md, chuyển NEXT sang Vòng 3 (skip theo nền tảng).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); Vòng 1 chỉ sửa
  tài liệu cho đúng code, chưa đổi chính sách.
