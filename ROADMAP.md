# ROADMAP — Excel Local Expert

Cập nhật: 2026-10-31, sau khi đóng Chu kỳ 0 (nghiệm thu Windows + giao diện).
Thứ tự ưu tiên của sản phẩm: đúng > an toàn/riêng tư > dễ bảo trì > hiệu năng >
trải nghiệm > tính năng mới. Mỗi mục chỉ vào chu kỳ khi có tiêu chí đo được.

## P0 — Đúng & an toàn (làm trước)

| # | Vấn đề | Bằng chứng hiện tại | Hướng xử lý |
|---|--------|---------------------|-------------|
| S-1 | PrivacyGuard bỏ lọt tên riêng đứng một mình | `test_privacy_guard_cases.py`: 2 known hit (`Nguyễn Văn An`, `nguyen van an`) | Thêm luật họ phổ biến + >=2 từ Title Case, trừ địa danh; đo lại trên bộ mở rộng |
| S-2 | PrivacyGuard che nhầm số tiền 9 chữ số thành CMND | 1 known hit: `123456789 đồng` → `[CMND_AN_DANH]` | Luật ngữ cảnh: bỏ qua khi sau số là đơn vị tiền (đồng/vnđ/vnd/đ) |
| S-3 | Che nhầm mã nghiệp vụ 13 chữ số thành STK | 1 known hit: `mã 0012345678901` → `[STK_AN_DANH]` | Bỏ qua khi trước số là "mã"/"mã hồ sơ"/"mã đơn" |
| S-4 | MSSV 9 chữ số bị che thành CMND | 1 known hit: `MSSV 524002081` | Bỏ qua khi trước số là MSSV/mã sinh viên/mã SV |

Tiêu chí chung cho S-1..S-4: 0 rò rỉ ngoài danh sách known, 0 che nhầm ngoài danh
sách known trên bộ >= 80 mẫu PII + >= 30 mẫu không phải PII; không nới lỏng luật
SĐT/CCCD/STK/MST/email/khóa.

## P1 — Dễ bảo trì

| # | Vấn đề | Số đo | Hướng |
|---|--------|-------|-------|
| M-1 | `gui()` gần 930 dòng (App lồng trong hàm) | `app.py` | Tách từng khối ra hàm/module cấp module, mỗi lát một commit, `test_gui` + `test_ui_acceptance` phải xanh |
| M-2 | `open_ai_settings()` ~155 dòng | `app.py` | Ứng viên lát cắt đầu tiên |
| M-3 | `review()` ~100 dòng | `app.py` | Lát cắt sau |
| M-4 | `infer_smart_plan` ~478 dòng, `prepare_sheets` ~261 dòng | `app.py` | Tách theo nhánh thuật toán, golden test `test_infer_golden` bảo vệ |
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
