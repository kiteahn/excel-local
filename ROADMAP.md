# ROADMAP — Excel Local Lazy

Cập nhật: 2026-11-03, sau Chu kỳ 7 (Q4 — dữ liệu giống thật).
Thứ tự ưu tiên của sản phẩm: đúng > an toàn/riêng tư > dễ bảo trì > hiệu năng >
trải nghiệm > tính năng mới. Mỗi mục chỉ vào chu kỳ khi có tiêu chí đo được.

## P0 — Đúng & an toàn

Không còn mục mở. S-1..S-4 **ĐÓNG** ở bản 0.3.2; bằng chứng:
`test_privacy_guard_cases.py` **125 PII + 87 non-PII = 212 câu** (mở rộng ở Chu kỳ 5), 0 rò lọt ngoài danh sách, 0% che nhầm, 2 known hit có ghi chú; mutation 5 chỗ đều bị bắt.

| # | Vấn đề | Trạng thái | Bằng chứng |
|---|--------|------------|------------|
| S-1 | PrivacyGuard bỏ lọt tên riêng đứng một mình | ĐÓNG 0.3.2 | 9 mẫu tên đứng một mình mới; loại trừ địa danh |
| S-2 | Che nhầm số tiền 9 chữ số thành CMND | ĐÓNG 0.3.2 | 4 mẫu tiền (đồng/vnđ/đ/nghìn) giữ nguyên |
| S-3 | Che nhầm mã nghiệp vụ 13 chữ số thành STK | ĐÓNG 0.3.2 | 2 mẫu "mã hồ sơ"/"mã đơn hàng" giữ nguyên |
| S-4 | MSSV 9 chữ số bị che thành CMND | ĐÓNG 0.3.2 | 3 mẫu MSSV/mã SV/mã sinh viên giữ nguyên |

Known còn lại (chấp nhận, đã đo): tên 2 từ; tên dân tộc không bắt đầu bằng họ Kinh.

## P0.5 — Đúng trên dữ liệu giống thật (Chu kỳ 7 / Q4)

| # | Việc | Kết quả đo | Ghi chú |
|---|------|------------|---------|
| R-1 | 9 workbook tổng hợp, 6 loại bảng nguồn | 9/9 đọc được | `examples/real_like/`, tái sinh bằng `create_real_like_examples.py` |
| R-2 | 55 câu tiếng Việt, kế hoạch đúng ngay lần đầu | **53/55 = 96,4%** | 50 curated + 5 câu tự nhiên; `results/plan-accuracy-q4.json` |
| R-3 | Chạy đầy đủ preview → Excel → verify → preserve | **49 thành công, 1 từ chối đúng `#N/A`, 0 lỗi** | Mọi bảng có ít nhất một thao tác chạy thật |
| R-4 | Ô rỗng `inlineStr` / dòng `TỔNG CỘNG` | ĐÃ VÁ | Crash `NoneType.findall`; dòng tổng bị bỏ qua và ghi `total_rows_skipped` |
| R-5 | Trung bình/AVG | **CHƯA HỖ TRỢ** | 2 câu còn sai; planner trả `None`, cần clarify và mở F-1 |

## P1 — Dễ bảo trì (trọng tâm Chu kỳ 2–4)

| # | Vấn đề | Số đo | Hướng |
|---|--------|-------|-------|
| M-1 | `gui()` 713 → **23 dòng**; `App` chuyển vào `_build_app_class` 699 dòng | `app.py` (AST) | ĐÓNG phần vỏ `gui()` ở Chu kỳ 9; không tách tiếp lớp `App` vì rủi ro giao diện lớn hơn lợi ích |
| M-2 | `open_ai_settings()` ~155 dòng | ĐÓNG Chu kỳ 2 | Đã thành `open_ai_settings_dialog()`; 141 chuỗi giữ nguyên, UI test 48/48 |
| M-3 | `review()` ~100 dòng | ĐÓNG Chu kỳ 3 | Đã thành `review_plan_dialog()`; 172 chuỗi giữ nguyên, `test_gui` 14/14 |
| M-4 | `prepare_sheets` 177 dòng sau Q4 (đã tách nhánh `lookup` ở Chu kỳ 4) | `app.py` (AST) | Lát cắt tiếp theo: tách `summary`/`clean`/`filter` theo cùng quy trình vàng + mutation |
| M-6 | `_filter_plan` 226 → **121 dòng**; `_lookup_plan` còn 202 dòng | `ai_client.py` (AST) | ĐÓNG phần `_filter_plan` ở Chu kỳ 9 (3 helper, 6/6 mutation); `_lookup_plan` để vòng sau |
| M-5 | ~30 `except Exception` rộng | `app.py`, `ai_client.py` | Thu hẹp dần, mỗi chỗ một test |

## P2 — Hiệu năng (số đo 5 lần, trung vị; `results/benchmark-q3-summary.json`)

| # | Số đo hiện tại (20.000 dòng) | Nút thắt | Hướng |
|---|----------------|----------|-------|
| P-1 | lookup 19,9s; summary 16,8s; clean 12,0s; đọc nguồn 5,1s; đọc danh mục 4,8s | Ghi công thức/COM + JSON bridge; `Write-Grid` đã ghi khối (`paste` 0,007s) | Không giữ thay đổi <10%; đo lại nếu đổi giao thức bridge |
| P-2 | Mỗi thao tác mở/đóng Excel riêng (~0,8–1,2s/lần) | bridge.ps1 | Tái dùng phiên tiết kiệm ~8% một lượt benchmark nhưng tăng rủi ro COM; chưa giữ |
| P-3 | Excel treo tạm thời 253–262s ở 2 lần đo | Môi trường Excel/máy | Theo dõi; dùng trung vị và ghi ngoại lệ, không giấu |

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
