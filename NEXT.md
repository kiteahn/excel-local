# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 17 — Đánh giá tách prepare_sheets (app.py, ~261 dòng)

Làm:
1. Đọc prepare_sheets: liệt kê các giai đoạn bên trong (tiền xử lý lookup /
   summary / clean / filter?) và biến trạng thái dùng chung. Nếu các giai
   đoạn theo action độc lập nhau, tách theo action như mẫu infer (mỗi khối
   một helper), kiểm chứng bằng: test_app 79, edge 61, usecases 34, và bộ
   test vàng mới cho prepare (chạy prepare_sheets trên fixture của test_app
   cho 4 action, ghim output plan/preview trước khi tách).
2. Nếu trạng thái đan xen không tách sạch được: ghi lý do vào log, chuyển
   sang mục dưới thay vì tách gượng.
3. Việc thay thế: xử lý VĐ-16 — (a) guard passthrough ở nhánh m_key của
   _lookup_plan giống Vòng 16; (b) mở rộng cụm kích hoạt lookup cho mẫu
   "áp <cột> lên" bị chen tên cột — chỉ làm nếu ghim đủ prompt vàng trước.

## Hàng đợi sau đó
- Vòng 18–19: tách gui (~890 dòng) nếu đủ phủ; rà soát lỗi còn lại trong
  UPGRADE_LOG (điểm cấu trúc/GUI).
- Vòng 20: TỔNG KẾT 20 VÒNG theo PROMPT_MUSE (bắt buộc): chấm lại toàn bộ,
  việc còn dở, hàng đợi Windows, câu hỏi chờ người dùng.
- Quyết định chờ người dùng: '500.000' (UC-20b); cloud/local mặc định.
- Hàng đợi Windows: UC-32/33/35/36; benchmark lại; test_gui + trực quan
  Vòng 10; nghiệm thu DPAPI thật.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Nhánh `muse-upgrade`: 17 commit (Bước 0 + Vòng 1–16), CHƯA push — cần
  người dùng xác nhận mới push lên GitHub (repo công khai).
