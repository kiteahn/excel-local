# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 5 — Đo PrivacyGuard bằng bộ mẫu ≥50 (mục E của PROMPT)

Mục tiêu: thay khẳng định "che PII" bằng số đo thật: tỷ lệ lọt (PII còn nguyên
sau anonymize_prompt) và tỷ lệ che nhầm (chuỗi hợp lệ không phải PII bị che).

Làm:
1. Đọc `PrivacyGuard` trong ai_client.py (anonymize_prompt, các regex, thứ tự
   áp dụng, anonymize_columns) để viết mẫu đúng cơ chế.
2. Tạo `test_privacy_guard_cases.py` với ≥50 mẫu có nhãn:
   - PII phải che: SĐT các đầu số 03/05/07/08/09, dạng +84, có dấu chấm/gạch/
     khoảng trắng xen giữa; CCCD 12 số, CMND 9 số; MST 10 và 13 số; STK ngân
     hàng độ dài hay gặp; email thường + email có dấu chấm/plus; mật khẩu sau
     chữ "mật khẩu/password"; tên riêng đứng một mình trong câu yêu cầu.
   - KHÔNG phải PII (không được che): mã hồ sơ, MSSV (vd 52400208), mã lớp,
     số tiền, điểm số, năm, mã `00123`, số điện thoại nội bộ ngắn, dãy số là
     một phần của yêu cầu nghiệp vụ.
3. Chạy, ghi số đo thật (lọt bao nhiêu mẫu nào, che nhầm bao nhiêu mẫu nào)
   vào UPGRADE_LOG. Mẫu nào lọt/che nhầm thì sửa regex trong ai_client ở mức
   tối thiểu, chạy lại: bộ mẫu mới + test_security_privacy (38) + test_router
   (7) + test_app (79) phải cùng đạt.
4. Ghi giới hạn còn lại (loại PII regex không thể bắt chắc) vào README mục
   Bảo mật nếu số đo cho thấy cần.
5. Commit, cập nhật log, chuyển NEXT sang Vòng 6 (VĐ-05: thu hẹp except).

## Hàng đợi sau đó
- Vòng 6: VĐ-05 — thu hẹp `except Exception` (30 chỗ), từng cụm có test.
- Vòng 7: VĐ-06 — tách hàm dài (bắt đầu từ infer_smart_plan hoặc prepare_sheets).
- Vòng 8: VĐ-07 — đo hiệu năng phân giai đoạn bằng benchmark.py.
- Vòng 9: VĐ-09 — dọn repo (gom *.json kết quả/benchmark vào thư mục riêng).
- Vòng 10: VĐ-10 — giao diện nhẹ (khóa nút khi busy, tooltip, phím tắt) +
  vòng tổng kết 10 vòng theo PROMPT.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
