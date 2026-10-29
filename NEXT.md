# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 7 — Tách hàm dài đợt 1 (VĐ-06, mục D của PROMPT)

Ứng viên đầu: `ai_client.infer_smart_plan` (~478 dòng). Nguyên tắc: KHÔNG đổi
hành vi — tách thuần cơ học theo nhóm case, chạy test sau mỗi bước tách nhỏ;
nếu một bước làm test đỏ thì hoàn tác bước đó.

Làm:
1. Đọc toàn bộ infer_smart_plan, lập bản đồ các khối độc lập (Q&A giáo dục,
   nhận diện lookup/summary/clean/filter, trích cột, đồng nghĩa...).
2. Tách 1–2 khối rõ ràng nhất thành hàm module-level có tên mô tả, giữ chữ ký
   và thứ tự đánh giá điều kiện y hệt.
3. Kiểm tra sau mỗi bước: router 7, security 38, app 79, edge 61,
   privacy corpus — phải đạt như trước (đây là các test chạm infer_smart_plan).
4. Commit từng bước tách nếu đủ lớn; cập nhật UPGRADE_LOG (hàm mới, số dòng
   trước/sau), NEXT chuyển sang Vòng 8 (hiệu năng).

## Hàng đợi sau đó
- Vòng 8: VĐ-07 — đo hiệu năng phân giai đoạn (benchmark.py; tách phần đo
  được trên Linux khỏi phần cần Excel COM; không nhận cải tiến nào không có
  số đo trước/sau).
- Vòng 9: VĐ-09 — dọn repo: gom *.json kết quả/benchmark vào `results/`,
  script phụ vào `tools/` (cập nhật đường dẫn ghi file trong test tương ứng).
- Vòng 10: VĐ-10 giao diện nhẹ + vòng tổng kết 10 vòng (chấm lại điểm Bước 0,
  sắp lại backlog theo PROMPT).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
