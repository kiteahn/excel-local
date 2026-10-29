# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 13 — Tách tiếp infer_smart_plan: khối Filter (~170 dòng)

Tiếp VĐ-06 theo cách cơ học của Vòng 7 (tách từng khối tự chứa, không đổi
hành vi, test sau mỗi bước; bước nào đỏ thì hoàn tác bước đó).

Làm:
1. Đọc khối Filter (mục 5, 5a–5e) trong infer_smart_plan hiện tại (375 dòng):
   xác định biến vào/ra — khối này dùng p_lower, source_columns và các hàm
   extract_requested_columns/find...; kiểm tra nó có phụ thuộc biến cục bộ
   nào từ khối Lookup/Summary không.
2. Nếu tự chứa: tách thành `_filter_plan(p_lower, source_columns)`. Nếu dính
   biến chung thì chỉ tách các tiểu khối 5a (đăng ký đơn/team) hoặc 5c (điều
   kiện số) nếu chúng tự chứa; phần còn lại để nguyên và ghi lý do.
3. Kiểm tra sau mỗi bước: router 7, security 38, app 79, edge 61, usecases 34,
   fuzz 20, privacy corpus — đạt như trước.
4. Commit, cập nhật UPGRADE_LOG, NEXT sang Vòng 14 (khối Lookup hoặc
   prepare_sheets).

## Hàng đợi sau đó
- Vòng 14+: tách khối Lookup trong infer_smart_plan; tách prepare_sheets
  (~261 dòng); tách gui (~890 dòng) — chỉ khi test đủ phủ.
- Mutation đợt 2 cho ai_client.py (PrivacyGuard regex, sanitize_plan,
  cascade router) theo quy trình khôi phục bằng bản sao file.
- Quyết định chờ người dùng: '500.000' một dấu chấm = 500 nghìn hay 500
  thập phân? (USECASES.md UC-20b)
- Hàng đợi Windows: UC-32/33/35/36; benchmark lại; test_gui + kiểm tra trực
  quan Vòng 10; nghiệm thu DPAPI thật.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định: Smart Router (có cloud) hay Ollama-only?
- Nhánh `muse-upgrade`: 13 commit (Bước 0 + Vòng 1–12), CHƯA push — cần
  người dùng xác nhận mới push lên GitHub (repo công khai).
