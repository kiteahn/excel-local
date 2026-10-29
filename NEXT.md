# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 8 — Đo hiệu năng phân giai đoạn (VĐ-07, mục F của PROMPT)

Nguyên tắc: chỉ giữ thay đổi khi số đo cải thiện thật và đối soát vẫn đúng;
không bịa số. Môi trường này KHÔNG có Excel COM nên chỉ đo được phần thuần
Python; phần bridge phải ghi "CHƯA ĐO ĐƯỢC ở đây" kèm cách đo trên Windows.

Làm:
1. Đọc benchmark.py và benchmark_ai.py hiện có: đo cái gì, ghi ra đâu,
   hash đã lệch code hiện tại ra sao (VĐ-07).
2. Viết/chạy đo phân giai đoạn trên Linux cho phần đo được: đọc file bằng
   `load_xlsx_fast` (nếu openpyxl có sẵn; nếu không, ghi rõ), parse CSV,
   preview_plan thuần Python trên dữ liệu mẫu sinh trong bộ nhớ ở quy mô
   100 / 1.000 / 5.000 / 20.000 dòng (lookup + summary). Ghi số thật vào
   UPGRADE_LOG kèm cấu hình máy đo.
3. Nếu thấy nút thắt thuần Python rõ ràng (vòng lặp bậc hai, chuẩn hóa lặp
   lại...) thì sửa tối thiểu, đo trước/sau, chạy toàn bộ suite đối soát.
   Không có nút thắt thì ghi kết luận "không có" — đó cũng là kết quả.
4. Cập nhật benchmark-*.json hoặc tạo `results/` ghi chú số đo cũ đã hết hạn
   (liên quan Vòng 9 dọn repo — phối hợp, đừng làm trùng).
5. Commit, cập nhật log, NEXT sang Vòng 9.

## Hàng đợi sau đó
- Vòng 9: VĐ-09 — dọn repo: gom *.json kết quả/benchmark vào `results/`,
  script phụ vào `tools/` (cập nhật đường dẫn ghi file trong test tương ứng).
- Vòng 10: VĐ-10 giao diện nhẹ + vòng tổng kết 10 vòng (chấm lại điểm Bước 0,
  sắp lại backlog theo PROMPT).
- Dài hạn: tách tiếp Lookup/Filter trong infer_smart_plan, `gui`,
  `prepare_sheets`; đo phần Excel COM trên Windows; NER/từ điển cho tên
  riêng đứng một mình (giới hạn PrivacyGuard đã đo ở Vòng 5).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
