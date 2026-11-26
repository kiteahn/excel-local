# Ảnh giao diện — Q5

Chụp ngày **2026-11-03** từ ứng dụng Tk thật trên Windows, bằng
`capture_gui_screens.py`. Script dùng `examples/real_like/sinh_vien.xlsx`
(dữ liệu tổng hợp), không gọi AI/cloud, không ghi workbook, và thay cấu hình
bằng một dict rỗng để không lộ khóa API thật. Việc đọc file XLSX đi qua bridge
Excel như luồng thật; script chỉ đọc, không thực thi tác vụ.

Màn hình hiện tại: **1536×864, DPI 120 (125%)**.

| Tệp | Nội dung | Kích thước |
|---|---|---|
| `01-main.png` | Màn hình chính sau khi đọc 120 dòng; đang chọn Làm sạch cột `Họ và tên`, bảng xem trước 6 cột | 1160×900 |
| `02-preview.png` | Hộp kiểm tra trước khi chấp nhận: badge `Ô làm sạch: 17`, nguồn, cột, 10 dòng mẫu, nút Hủy/Chấp nhận | 1080×680 |
| `03-settings.png` | Hộp Cài đặt AI & Smart Router với khóa API để trống (không có khóa thật trong ảnh) | 668×560 |

## Đã đo / đã sửa

- Ảnh chính cho thấy tiêu đề ứng dụng ghi `Excel Local Lazy v0.3` trong khi
  tiêu đề cửa sổ là `v0.3.2`. Đây là lệch số phiên bản hiển thị; bản 0.4.0 đã
  đồng bộ cả hai thành `v0.4.0`. Ảnh cũ vẫn giữ nguyên để đối chiếu lịch sử,
  không chụp lại và không chỉnh sửa ảnh.
- Ba ảnh đều nằm trong màn hình 1536×864; không có nút chính nào bị cắt ở
  125% DPI.
- Script chụp lần đầu bị lỗi thứ tự callback (chỉ lưu ảnh chính); đã sửa thành
  chuỗi callback riêng và chụp lại đủ 3 ảnh.

## CẦN NGƯỜI XÁC NHẬN BẰNG MẮT

- Chữ tiếng Việt có dấu, khoảng cách, căn lề và độ tương phản trong cả 3 ảnh.
- Bố cục ở **100%** và **150% DPI** — ảnh này chỉ chứng minh mức 125% của máy
  hiện tại; **CHƯA KIỂM CHỨNG** 100%/150%.
- Màu/trạng thái nút `Chấp nhận và thực hiện`, `Chưa chấp nhận`, và các badge.
- Hộp xem trước nội dung gửi cloud: chưa chụp được vì là `messagebox` hệ thống;
  **CẦN NGƯỜI XÁC NHẬN BẰNG MẮT** khi bật cloud thật (không dùng khóa thật
  trong test).

Ảnh chỉ là bằng chứng thị giác, không thay thế kiểm thử hành vi trong
`test_gui.py` (14/14) và `test_ui_acceptance.py` (48/48).
