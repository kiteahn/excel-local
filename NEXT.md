# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 14 — Tách nốt khối Clean + đánh giá khối Lookup trong infer_smart_plan

infer_smart_plan còn 204 dòng: khối Lookup (mục 2, ~150 dòng, nhiều biến
trạng thái target_dest/target_val/shared_key) và khối Clean (mục 4, ~25 dòng).

Làm:
1. Tách khối Clean trước (nhỏ, ít rủi ro) thành `_clean_plan(p_lower,
   source_columns)` theo đúng mẫu các vòng 7/13; kiểm chứng bằng suite +
   so sánh song song bản cũ/mới trên bộ prompt (giữ lại script so sánh ở
   /tmp hoặc viết thành file test_equivalence_infer.py trong repo với
   ~20 prompt — cách này bền hơn cho các vòng sau).
2. Khối Lookup: đọc kỹ luồng biến; nếu tách được theo các tiểu khối 2a/2b/2c
   (cột đích / cột giá trị / khóa chung) mà không phải truyền quá 4-5 tham
   số thì tách từng tiểu khối; nếu không, ghi rõ lý do giữ nguyên vào log
   thay vì tách gượng.
3. Kiểm tra: toàn bộ 15 file test exit 0 như các vòng trước.
4. Commit, cập nhật UPGRADE_LOG, NEXT sang Vòng 15 (prepare_sheets hoặc
   mutation đợt 2 cho ai_client).

## Hàng đợi sau đó
- Mutation đợt 2 cho ai_client.py (PrivacyGuard, sanitize_plan, cascade
  router) — khôi phục bằng bản sao file, đối chiếu byte (bài học Vòng 12).
- Tách prepare_sheets (~261 dòng), gui (~890 dòng) khi test đủ phủ.
- Tổng kết vòng 20 theo PROMPT (mỗi 10 vòng).
- Quyết định chờ người dùng: '500.000' (USECASES.md UC-20b); chính sách
  cloud/local mặc định.
- Hàng đợi Windows: UC-32/33/35/36; benchmark lại; test_gui + kiểm tra trực
  quan Vòng 10; nghiệm thu DPAPI thật.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Nhánh `muse-upgrade`: 14 commit (Bước 0 + Vòng 1–13), CHƯA push — cần
  người dùng xác nhận mới push lên GitHub (repo công khai).
