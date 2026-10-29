# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 15 — Mutation test đợt 2 cho ai_client.py

Đợt 1 (Vòng 12) phủ app.py. Đợt này phủ ai_client.py, dùng quy trình đã
rút kinh nghiệm: khôi phục bằng BẢN SAO FILE + đối chiếu byte, xác minh đột
biến áp dụng thật bằng diff trước khi chạy suite.

Đột biến ứng viên (chọn 5–6):
1. PrivacyGuard EMAIL_REGEX: bỏ phần domain → bộ corpus 53 mẫu phải đỏ.
2. sanitize_plan: cho phép action lạ đi qua → test_security phải đỏ.
3. build_local_fallback_plan / cascade: đảo thứ tự ưu tiên provider.
4. parse model JSON: bỏ bước gác sanitize (chỉ nếu có điểm chèn rõ ràng).
5. _lookup_plan: đổi điều kiện `reference_columns and ...` thành `or`.
6. redact_for_cloud: bỏ một lớp thay thế (ví dụ không che số tài khoản).

Quy trình mỗi đột biến: sao lưu file → áp dụng → xác minh diff → chạy suite
liên quan (router, security, privacy corpus, infer golden, usecases, app)
→ ghi BẮT/LỌT → khôi phục + đối chiếu byte. Đột biến lọt → bổ sung test
trong cùng vòng rồi chạy lại xác nhận bị bắt.

## Hàng đợi sau đó
- Vòng 16: xử lý VĐ-15 (lookup value='bảng') nếu phân tích cho thấy sửa an
  toàn: thêm prompt vàng cho các mẫu "điền/sang/áp" và sửa regex cột đích.
- Vòng 17+: tách prepare_sheets (~261 dòng, app.py) theo mẫu có test vàng;
  gui (~890 dòng) sau cùng.
- Tổng kết vòng 20 theo PROMPT (mỗi 10 vòng).
- Quyết định chờ người dùng: '500.000' (UC-20b); cloud/local mặc định.
- Hàng đợi Windows: UC-32/33/35/36; benchmark lại; test_gui + trực quan
  Vòng 10; nghiệm thu DPAPI thật.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Nhánh `muse-upgrade`: 15 commit (Bước 0 + Vòng 1–14), CHƯA push — cần
  người dùng xác nhận mới push lên GitHub (repo công khai).
