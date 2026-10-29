# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 12 — Fuzz nhẹ cho parser CSV/header + mutation test bộ test

Theo "CHẾ ĐỘ KHÔNG CẠN VIỆC" trong PROMPT_MUSE.md.

Làm:
1. Fuzz CSV: sinh có hệ thống các biến thể quanh load_csv/inspect CSV —
   BOM lạ, dòng trống xen kẽ, nháy đơn/kép lồng nhau, phân cách hỗn hợp,
   ô đúng 32767 ký tự (đạt) và 32768 (từ chối), dòng cuối thiếu newline,
   CRLF vs LF, ký tự điều khiển ở các vị trí khác nhau. Kỳ vọng: hoặc đọc
   đúng, hoặc ValueError rõ ràng — không bao giờ traceback lạ/sập.
2. Fuzz header: snapshot dict với header rỗng từng phần, toàn khoảng trắng,
   ký tự đặc biệt, tên cột cực dài, trùng sau NFC/NFD — qua select_header
   và detect_header_row.
3. Mutation test: chọn 4–6 đột biến nhỏ trong app.py (đổi toán tử so sánh
   trong validate số, bỏ chuẩn NFC ở key_value, đổi dung sai...) — áp dụng
   tạm thời, chạy suite liên quan, ghi đột biến nào bị test bắt / lọt.
   Đột biến lọt → bổ sung test. Hoàn tác mọi đột biến sau khi đo (kiểm tra
   `git diff` sạch trước khi commit phần test mới).
4. Cập nhật UPGRADE_LOG (số mẫu fuzz, số đột biến bị bắt/lọt), USECASES.md
   nếu có use case mới. Commit.

## Hàng đợi sau đó
- Vòng 13+: tách tiếp khối Lookup (~150 dòng) rồi Filter (~170 dòng) trong
  infer_smart_plan theo cách cơ học như Vòng 7; tách prepare_sheets, gui.
- Quyết định chờ người dùng: '500.000' một dấu chấm nên hiểu là 500 nghìn
  hay 500 thập phân trong ngữ cảnh tiền VN? (chi tiết ở USECASES.md UC-20b)
- Hàng đợi Windows (chỉ làm được trên máy Windows có Excel): UC-32/33/35/36;
  đo lại benchmark.py cho code hiện tại; chạy test_gui.py thật + kiểm tra
  trực quan tooltip/phím tắt (Vòng 10) và messagebox DPAPI (Vòng 3);
  nghiệm thu DPAPI thật (security mục 16–17, router mục 1/7).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
- Nhánh `muse-upgrade` hiện có 12 commit (Bước 0 + Vòng 1–11) CHƯA push lên
  GitHub (repo công khai) — cần người dùng xác nhận mới push.
