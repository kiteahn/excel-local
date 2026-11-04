# Chu kỳ 8 — Q5: giao diện, ảnh chụp và lỗi đo được

Ngày: 2026-11-03. Nhánh `muse-upgrade`; không push; không dùng file thật, không
gọi cloud. Việc đọc file XLSX mẫu đi qua bridge Excel như luồng thật; script chỉ
đọc, không ghi workbook và không thực thi tác vụ.

## Ảnh chụp

`capture_gui_screens.py` mở ứng dụng Tk thật, ép cửa sổ `-topmost` + `lift()`,
dùng `examples/real_like/sinh_vien.xlsx` (120 dòng dữ liệu tổng hợp) và lưu:

| Tệp | Nội dung |
|---|---|
| `docs/screens/01-main.png` | Màn hình chính sau khi đọc dữ liệu, đang chọn làm sạch cột `Họ và tên` |
| `docs/screens/02-preview.png` | Hộp kiểm tra trước khi chấp nhận, có nút Hủy/Chấp nhận |
| `docs/screens/03-settings.png` | Hộp Cài đặt AI, khóa API để trống |

Máy chụp: màn hình 1536×864, DPI 120 (125%). Script thay cấu hình bằng một
dict rỗng để ảnh không chứa khóa API thật. Ảnh chỉ là bằng chứng thị giác,
không thay thế test hành vi.

## Lỗi đo được và đã sửa

- Lần chụp đầu chỉ lưu được ảnh chính vì callback lồng nhau chạy sai thứ tự.
  Đã tách thành chuỗi callback riêng; chụp lại đủ 3 ảnh.
- Ảnh chính cho thấy tiêu đề trong app ghi `Excel Local Expert v0.3` còn tiêu đề
  cửa sổ ghi `v0.3.2`. Đây là lệch số phiên bản hiển thị; ghi nhận để đồng bộ
  khi tăng phiên bản ở Q7, chưa sửa bố cục.

## Kiểm thử

- `test_gui.py`: **14/14**.
- `test_ui_acceptance.py`: **48/48**.
- `gui-smoke`: đạt; cửa sổ tạo và đóng được trong phiên chụp.
- Không đổi bố cục tổng thể, không đổi luồng thao tác.

## CẦN NGƯỜI XÁC NHẬN BẰNG MẮT

- Chữ tiếng Việt có dấu, khoảng cách, căn lề, độ tương phản trong cả 3 ảnh.
- Bố cục ở **100%** và **150% DPI**: **CHƯA KIỂM CHỨNG**; ảnh hiện tại chỉ
  chứng minh mức 125% của máy chụp.
- Màu/trạng thái các nút `Chấp nhận và thực hiện`, `Chưa chấp nhận` và badge.
- Hộp xem trước nội dung gửi cloud: chưa chụp được vì là hộp thoại hệ thống;
  cần người xác nhận khi bật cloud thật.
