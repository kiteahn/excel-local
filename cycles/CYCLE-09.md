# Chu kỳ 9 — Q6: tách `_filter_plan`, `App.__init__` và `gui()`

Ngày: 2026-11-04. Nhánh `muse-upgrade`; không push; không dùng file thật.

## Mục tiêu và cách làm

Q1 đã tách nhánh lookup thành công nên Q6 được phép tách tiếp. Mỗi lát là một
phép rút hàm thuần túy, không đổi công thức, nhãn, thứ tự cột, thông điệp hay
luồng xem trước/chấp nhận. Đo AST trước/sau, chạy bộ vàng, mutation và test Excel
thật ở cuối.

## Kết quả đo AST

| Hàm | Trước | Sau | Ghi chú |
|---|---:|---:|---|
| `ai_client._filter_plan` | 226 dòng | **121 dòng** | tách 3 helper: đăng ký 46, cán sự 40, điều kiện số 38 |
| `app.gui` | 713 dòng | **23 dòng** | phần thân App chuyển sang `_build_app_class` (699 dòng) |
| `App.__init__` | 213 dòng | **16 dòng** | phần dựng widget chuyển sang `_build_ui` (199 dòng) |

`gui()` giữ nguyên chữ ký `gui(smoke=False, on_ready=None)`, vẫn import Tk bên
trong hàm, tạo `App` rồi mở cửa sổ như cũ. `_build_app_class(tk, ttk, filedialog,
messagebox)` nhận đúng các module Tk mà `gui()` đã import, nên không thêm phụ
thuộc và không đổi cách chạy headless/smoke.

## Vàng và mutation

- Thêm 1 ca vàng cho nhánh **đăng ký đội** (`Đăng ký đội cho sinh viên`), trước
  đây không được bộ vàng chạm tới. `test_infer_golden.py` nay có **34 ca**.
- Mutation 6 chỗ trên mã đã tách, tất cả **6/6 bị bắt**:
  1. đổi giá trị nhánh đội `Đội` → `ĐộiX`;
  2. bỏ dấu cách trong `filter_val = f"{op} {num_clean}"`;
  3. thêm một vai trò lạ vào danh sách cán sự;
  4. đổi giá trị nhánh đơn `Đơn` → `ĐơnX`;
  5. đổi khóa nhánh đơn từ cột `Hình thức` sang cột đầu tiên;
  6. thêm `X` vào giá trị phòng ban.
- Script `mutation_q6.py` ghi file bằng bytes, giữ nguyên CRLF, và phục hồi
  `ai_client.py` trong `finally`; đã kiểm tra SHA-256 trước/sau khớp nhau.

## Kiểm thử

- `test_infer_golden.py`: **34/34**.
- `test_single_file_filter.py`: **30/30**.
- `test_usecases.py`: **37/37**.
- `test_real_like_plans.py --no-execute`: **50/50** kế hoạch curated đúng,
  **53/55** tổng thể (2 ca AVG/trung bình vẫn là giới hạn đã ghi).
- `test_gui.py`: **14/14**; `test_ui_acceptance.py`: **48/48**; `gui-smoke`: đạt.
- Toàn bộ `run_all_tests.ps1` (gồm Excel thật): **31 lệnh, 0 lỗi, 1018,7s,
  exit 0**; `examples/` không đổi, không còn Excel mồ côi, không có thư mục tạm
  mới. Log `.work/qa-logs/run-20261104-105727.log`.
- Ngoại lệ môi trường: `test_real_like_plans.py` mất **703,6s** thay vì khoảng
  327s như Q4; quan sát khoảng trễ khoảng 4 phút giữa hai workbook (11:09:34 →
  11:13:44), chưa xác định nguyên nhân. Không có lỗi, không giấu; các suite còn
  lại vẫn đạt.

## Dừng ở đâu và vì sao

- `_build_app_class` còn 699 dòng vì toàn bộ lớp `App` nằm trong đó. Tách tiếp
  các nhóm widget/state sẽ phải sửa nhiều điểm gắn kết và tăng rủi ro hồi quy
  giao diện; lợi ích không rõ so với chi phí. Dừng tại đây, giữ `gui()` 23 dòng
  làm cửa vào ổn định.
- `_filter_plan` còn 121 dòng cho các nhánh 5d/5e; đây là chuỗi quyết định có
  thứ tự ưu tiên, tách thêm dễ làm đổi thứ tự khớp. Dừng sau khi đã tách 3 nhánh
  lớn và có vàng + mutation bảo vệ.

## Giới hạn / chưa kiểm chứng

- Refactor không đổi hành vi, nên không tăng số phiên bản ở bước này.
- Chưa đo bộ nhớ hoặc thời gian khởi động GUI trước/sau; chỉ đo cấu trúc AST và
  chạy test hành vi.
- Các mục giao diện cần mắt người vẫn giữ ở `NEXT.md` / `CYCLE-08.md`.
