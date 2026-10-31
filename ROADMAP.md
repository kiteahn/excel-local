# ROADMAP — Excel Local Expert

Cập nhật: 2026-10-31, sau khi đóng Chu kỳ 3 (tách hộp kiểm tra/chấp nhận khỏi `gui()`).
Thứ tự ưu tiên của sản phẩm: đúng > an toàn/riêng tư > dễ bảo trì > hiệu năng >
trải nghiệm > tính năng mới. Mỗi mục chỉ vào chu kỳ khi có tiêu chí đo được.

## P0 — Đúng & an toàn

Không còn mục mở. S-1..S-4 **ĐÓNG** ở bản 0.3.2; bằng chứng:
`test_privacy_guard_cases.py` 80 PII + 37 non-PII, 0 rò lọt / 0 che nhầm ngoài
danh sách, 2 known hit có ghi chú; mutation 5 chỗ đều bị bắt.

| # | Vấn đề | Trạng thái | Bằng chứng |
|---|--------|------------|------------|
| S-1 | PrivacyGuard bỏ lọt tên riêng đứng một mình | ĐÓNG 0.3.2 | 9 mẫu tên đứng một mình mới; loại trừ địa danh |
| S-2 | Che nhầm số tiền 9 chữ số thành CMND | ĐÓNG 0.3.2 | 4 mẫu tiền (đồng/vnđ/đ/nghìn) giữ nguyên |
| S-3 | Che nhầm mã nghiệp vụ 13 chữ số thành STK | ĐÓNG 0.3.2 | 2 mẫu "mã hồ sơ"/"mã đơn hàng" giữ nguyên |
| S-4 | MSSV 9 chữ số bị che thành CMND | ĐÓNG 0.3.2 | 3 mẫu MSSV/mã SV/mã sinh viên giữ nguyên |

Known còn lại (chấp nhận, đã đo): tên 2 từ; tên dân tộc không bắt đầu bằng họ Kinh.

## P1 — Dễ bảo trì (trọng tâm Chu kỳ 2–4)

| # | Vấn đề | Số đo | Hướng |
|---|--------|-------|-------|
| M-1 | `gui()` còn 697 dòng (App lồng trong hàm) | `app.py` (AST) | Tách từng khối ra hàm cấp module, mỗi lát một commit, `test_gui` + `test_ui_acceptance` phải xanh |
| M-2 | `open_ai_settings()` ~155 dòng | ĐÓNG Chu kỳ 2 | Đã thành `open_ai_settings_dialog()`; 141 chuỗi giữ nguyên, UI test 48/48 |
| M-3 | `review()` ~100 dòng | ĐÓNG Chu kỳ 3 | Đã thành `review_plan_dialog()`; 172 chuỗi giữ nguyên, `test_gui` 14/14 |
| M-4 | `prepare_sheets` 262 dòng (không phải `infer_smart_plan`: hàm này đã chỉ còn 25 dòng sau các vòng trước) | `app.py` (AST) | Lát cắt tiếp theo (Chu kỳ 4): tách theo nhánh `lookup`/`summary`/`clean`/`filter`, `test_app` + `test_usecases` bảo vệ |
| M-6 | `ai_client._lookup_plan` 180 dòng, `_filter_plan` 178 dòng | `ai_client.py` (AST) | Tách sau M-4, dùng `test_infer_golden` (28 prompt vàng) |
| M-5 | ~30 `except Exception` rộng | `app.py`, `ai_client.py` | Thu hẹp dần, mỗi chỗ một test |

## P2 — Hiệu năng (số đo 1.000 dòng, `results/benchmark-current-1000.json`)

| # | Số đo hiện tại | Nút thắt | Hướng |
|---|----------------|----------|-------|
| P-1 | inspect 4,05s; lookup 6,28s; summary 6,13s; clean 5,31s | Mở Excel COM ~1s/lần + ghi ô theo khối 1,4–2,0s | Đo lại ở 5.000/20.000 dòng; thử đọc XLSX thuần Python cho bước kiểm tra (không thay đường ghi) |
| P-2 | Mỗi thao tác mở/đóng Excel riêng | bridge.ps1 | Chỉ giữ thay đổi nếu đối soát vẫn đúng và không để tiến trình mồ côi |

## P3 — Trải nghiệm

| # | Việc | Ghi chú |
|---|------|---------|
| U-1 | PDF bảng rộng có thể bị cắt cột | Cần đo bố cục với bảng 8–10 cột, điều chỉnh vùng in |
| U-2 | Nhớ thư mục file gần nhất | Giảm thao tác chọn file |
| U-3 | Bảng phím tắt trong app | Hiện có Ctrl+Enter; thêm Ctrl+O (chọn file), F5 (đọc lại) nếu test được |

## P4 — Tính năng (chỉ sau khi P0–P2 ổn)

| # | Tính năng | Giá trị | Ghi chú |
|---|-----------|---------|---------|
| F-1 | Tổng hợp có ĐẾM / TRUNG BÌNH / MIN / MAX | Kế toán, điểm | Mở rộng schema plan + template công thức + đối soát |
| F-2 | Lọc nhiều điều kiện (AND) | Danh sách lớp, điểm | Schema `conditions[]`, cần test biên |
| F-3 | Đổi tên cột trước khi xuất | Báo cáo | Giữ nguyên dữ liệu nguồn |

## Nguyên tắc cố định
- Không thêm thư viện ngoài (Python Standard Library).
- Mặc định Ollama local; cloud chỉ khi người dùng bật.
- AI chỉ đề xuất; thực thi sau xem trước + token khớp.
- Không sửa file người dùng; chỉ `examples/` và dữ liệu tự sinh.
- Không push; chỉ commit trên `muse-upgrade`.
