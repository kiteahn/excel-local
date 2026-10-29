# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 10 — Giao diện nhẹ (VĐ-10, mục I) + Tổng kết 10 vòng

Phần giao diện chỉ chỉnh NHẸ, giữ nguyên bố cục và luồng (chọn file → tìm
dòng → chọn AI → xem trước → chấp nhận). Mọi thay đổi phải qua test_gui.py
trên Windows hoặc ghi rõ CHƯA kiểm tra trực quan (môi trường này headless,
test_gui SKIP).

Làm:
1. Soi code gui() trong app.py và lập danh sách kiểm chứng được bằng code:
   - Nút nào chưa bị khóa khi `busy` (work() đã khóa run_button; kiểm tra
     các nút Chấp nhận / Batch / Dừng và ô nhập còn lại).
   - Tooltip: đã có cơ chế tooltip chưa; các tùy chọn khó hiểu (Tự chữa
     lành, Lenient, auto_fallback) đã có giải thích chưa.
   - Phím tắt: Enter/Ctrl+Enter chạy đề xuất? Esc đóng hộp thoại?
   - Thứ tự Tab và focus ban đầu.
2. Chỉ sửa những điểm thiếu rõ ràng, nhỏ, không đổi tên tính năng/luồng.
   Mỗi sửa đổi: py_compile + suite logic đạt như trước; ghi rõ phần nào
   CHƯA kiểm tra trực quan.
3. Tổng kết 10 vòng (theo PROMPT): chấm lại 8 hạng mục của Bước 0, nêu điểm
   nào tăng/giảm kèm bằng chứng, sắp lại backlog cho chặng tiếp theo
   (Vòng 11+: tầng kiểm tra 2 — liệt kê ≥30 use case và tự chạy phần chạy
   được; tách Lookup/Filter; hàng đợi Windows).

## Hàng đợi Windows (chỉ làm được trên máy Windows có Excel)
- Đo lại `benchmark.py` cho code hiện tại; tối ưu khởi động Excel COM nếu
  số đo xác nhận nút thắt ~5 giây.
- Chạy test_gui.py thật + kiểm tra trực quan các thay đổi Vòng 3/6/10.
- Nghiệm thu DPAPI thật (test_security_privacy mục 16–17, test_router).

## Dài hạn
- Tách tiếp khối Lookup/Filter trong infer_smart_plan, `gui` (~890 dòng),
  `prepare_sheets` (~261 dòng) theo cách cơ học từng khối như Vòng 7.
- Tên riêng đứng một mình: cần từ điển/NER ngoài regex (Vòng 5 đã đo giới hạn).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
