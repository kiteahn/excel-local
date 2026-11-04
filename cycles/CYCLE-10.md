# Chu kỳ 10 — Q7: nền tảng phát hành

Ngày: 2026-11-04. Nhánh `muse-upgrade`; **không push**; không tự thêm `LICENSE`.

## Phiên bản và đồng bộ hiển thị

- Tăng phiên bản từ 0.3.2 lên **0.4.0** vì đợt này có thay đổi người dùng thấy:
  PrivacyGuard 212 câu, xem trước cloud, dữ liệu giống thật, ảnh GUI và CI.
- Đồng bộ tiêu đề cửa sổ và nhãn đầu trang thành `v0.4.0` (trước đó lệch
  `v0.3.2` / `v0.3`). Ảnh `docs/screens/` được chụp lại sau khi đổi tiêu đề.
- Cập nhật `README.md`, `RELEASES.md`, `release-notes-0.4.0.md`; tag cục bộ
  `v0.4.0`. Không tạo tag remote, không push.

## CI không cần Excel

`.github/workflows/tests.yml` chạy trên `windows-latest` và `ubuntu-latest`,
Python 3.12, chỉ các test không cần Excel:

- 17 file: `test_app.py`, `test_all_edge_cases.py`, `test_upgrades.py`,
  `test_cli.py`, `test_usecases.py`, `test_usecases2.py`, `test_fuzz_light.py`,
  `test_fuzz_headers.py`, `test_fuzz_xlsx.py`, `test_infer_golden.py`,
  `test_sanitize_plan.py`, `test_policy_defaults.py`, `test_safety_limits.py`,
  `test_prepare_sheets_golden.py`, `test_batch_summary.py`,
  `test_privacy_guard_cases.py`, `test_cloud_preview.py`.
- Trên Windows chạy thêm `test_router.py` và `test_security_privacy.py` (có phần
  DPAPI thật).
- **Loại khỏi CI:** `test_headers.py`, `test_single_file_filter.py`,
  `test_diverse_scenarios_e2e.py` — trên Windows các bộ này gọi Excel Desktop,
  GitHub runner không có Excel. Trên Linux chúng tự SKIP nhưng không đủ để gọi
  là "không cần Excel" trên Windows.
- Cú pháp YAML đã kiểm tra bằng PyYAML 6.0.3: đọc được, có `jobs.test`, matrix
  đúng 2 hệ điều hành, danh sách đúng 17 file.
- Bộ chọn 17 file đã chạy trong `run_all_tests.ps1` trên Windows (31 lệnh, 0 lỗi).
  **CHƯA KIỂM CHỨNG** một lượt GitHub Actions thật trên Ubuntu.

## Kiểm thử sau khi đổi tiêu đề

- `python -m py_compile app.py`: đạt.
- `test_gui.py`: **14/14**.
- `test_ui_acceptance.py`: **48/48**.
- Chụp lại 3 ảnh GUI: 1160×900, 1080×680, 668×560; ảnh chính hiển thị
  `Excel Local Expert v0.4.0`.
- Toàn bộ lưới Excel gần nhất (trước khi chỉ đổi chuỗi tiêu đề) vẫn là
  **31 lệnh, 0 lỗi, 1018,7s, exit 0**; thay đổi Q7 chỉ là chuỗi hiển thị, tài
  liệu, workflow và ảnh, không đổi logic xử lý.

## Chính sách và giới hạn

- Cloud vẫn chỉ chạy khi người dùng tự bật; trước khi gửi có hộp xem trước đúng
  nội dung đã che + nút hủy. Không dùng ngôn ngữ "tuyệt đối"/"không bao giờ lọt".
- **Không thêm LICENSE.** Đây là quyết định của chủ sản phẩm; câu hỏi và lựa
  chọn nằm trong `NEXT.md`.
- **Không push**, không tạo tag remote. Chỉ tag cục bộ `v0.4.0`.
- Giao diện 100%/150% DPI và hộp xem trước cloud vẫn **CẦN NGƯỜI XÁC NHẬN
  BẰNG MẮT**; ảnh hiện có ở 125% DPI.
- CI Ubuntu mới kiểm tra cú pháp và chạy bộ chọn trên Windows; chưa có bằng
  chứng runner Ubuntu thật.
