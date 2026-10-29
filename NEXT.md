# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 9 — Dọn repo (VĐ-09, mục H của PROMPT)

Hiện trạng root repo có 15+ file *.json kết quả/benchmark nằm lẫn mã nguồn
(acceptance-v0.2.json, benchmark-*.json, *-test-results.json,
logic-regression-0.2.1.json) và script phụ (create_diverse_examples.py,
generate_project_pdf.py, benchmark_ai.py).

Làm (cẩn thận — nhiều test GHI file kết quả theo đường dẫn cứng ở root):
1. Liệt kê mọi nơi ghi/đọc các file này: `grep -rn "test-results\|benchmark-" --include="*.py" .`
   và trong README/PROMPT.
2. Tạo `results/` và chuyển các file kết quả/benchmark lịch sử vào đó bằng
   `git mv`; cập nhật đường dẫn ghi trong các test tương ứng (test_app,
   test_cli, test_gui, test_headers, test_read_boundary...) sang `results/`.
   File nào là đầu vào test (không phải đầu ra) thì giữ nguyên vị trí.
3. Script phụ: cân nhắc `tools/` cho create_diverse_examples.py và
   generate_project_pdf.py — chỉ chuyển khi không có tham chiếu đường dẫn
   nào vỡ; nếu rủi ro cao hơn lợi ích thì giữ nguyên và ghi lý do vào log.
4. Kiểm tra: chạy lại toàn bộ suite trên Linux — tất cả exit 0 như Vòng 4,
   và `git status` cho thấy file kết quả mới sinh ra nằm trong results/.
5. Cập nhật README mục "Kiểm tra và mã nguồn" trỏ tới results/.
6. Commit, cập nhật UPGRADE_LOG, NEXT sang Vòng 10.

## Hàng đợi sau đó
- Vòng 10: VĐ-10 giao diện nhẹ (khóa nút khi busy — đã có một phần: run_button
  bị disable trong work(); kiểm tra các nút còn lại, tooltip, phím tắt,
  thứ tự Tab) + vòng tổng kết 10 vòng theo PROMPT (chấm lại điểm Bước 0,
  sắp lại backlog).
- Hàng đợi Windows (chỉ làm được trên máy Windows có Excel): đo lại
  benchmark.py cho v0.3.1+, tối ưu khởi động Excel COM; chạy test_gui thật;
  nghiệm thu DPAPI thật.
- Dài hạn: tách tiếp Lookup/Filter trong infer_smart_plan, `gui`,
  `prepare_sheets`; NER/từ điển cho tên riêng đứng một mình.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
