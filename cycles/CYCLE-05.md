# Chu kỳ 5 — PrivacyGuard trên câu người dùng thật + xem trước gửi cloud

## Mục tiêu (Q2)
- Kho ≥200 câu yêu cầu tiếng Việt thực tế về bảng tính, tách PII và câu bình thường.
- Lọt PII thật không tệ hơn mốc cũ; che nhầm câu bình thường ≤ 2%.
- Người dùng xem đúng đoạn chữ/tên cột sẽ gửi cloud và có nút hủy; chỉ khi cloud bật.
- README ghi giới hạn, không dùng "tuyệt đối"/"không bao giờ lọt".

## Cổng nghiệm thu
- `test_privacy_guard_cases.py` in số đo thật và assert tỷ lệ che nhầm ≤ 2%.
- `test_cloud_preview.py` chứng minh preview không gọi mạng, khớp `anonymize_prompt`, hủy thì không gửi.
- `test_gui.py`, `test_ui_acceptance.py`, `test_app.py`, `run_all_tests.ps1` đạt.
- `secret_scan.py` không còn finding ngoài allowlist đã đọc tay.

## Thay đổi
- `test_privacy_guard_cases.py`: 125 PII + 87 non-PII = **212 câu**; thêm assert tỷ lệ che nhầm.
- `ai_client.py`:
  - `PASSWORD_REGEX` cho phép tối đa 2 từ đệm trước `:`/`=` (vá `"Mật khẩu mới: ..."`).
  - `cloud_preview(prompt, source_columns, reference_columns=..., provider=..., api_key=...)` trả prompt/cột đã che, provider đích, `will_send`, `reason`; không gọi mạng.
- `app.py`: `confirm_cloud_preview(...)` (hàm cấp module, test được) và gọi trong `App.plan()` trước worker; Cancel dừng, ghi chú không dữ liệu rời máy.
- `test_cloud_preview.py`: 20 check.
- `README.md`: số mẫu, tỷ lệ, giới hạn, bước xem trước/hủy.

## Kết quả đo
- PII: 125 câu, **0 rò lọt ngoài 2 giới hạn đã ghi** (tên 2 từ; tên dân tộc không có họ Kinh).
- Non-PII: 87 câu, **0 che nhầm = 0,00%** (ngưỡng ≤2%).
- Test mục tiêu: `test_privacy_guard_cases.py` đạt; `test_cloud_preview.py` 20/20; `test_security_privacy.py` 43/43; `test_gui.py` 14/14; `test_ui_acceptance.py` 48/48; `test_app.py` 87/87.
- `secret_scan.py`: 0 finding ngoài allowlist (15 mẫu tổng hợp working tree, 12 lịch sử).
- `run_all_tests.ps1`: xem log `.work/qa-logs/` của lượt chạy cuối.

## Giới hạn còn lại
- Vẫn là regex: tên 2 từ, tên dân tộc không bắt đầu bằng họ Kinh, định dạng lạ có thể lọt.
- Tên đường/phố trùng họ người ngoài danh sách loại trừ có thể bị che nhầm.
- Chưa thử với câu người dùng thật ngoài 212 mẫu; chưa gọi cloud thật (đúng quy định).
- Hộp xem trước dự đoán `will_send` theo cùng logic router; nếu cache/router thay đổi trong tương lai phải cập nhật cùng lúc.
