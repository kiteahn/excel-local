# USECASES — Bảng use case tầng 2 (PROMPT_MUSE) và trạng thái kiểm chứng

Cập nhật: Vòng 11 (2026-10-30). "Đã chạy" = có check trong
`test_usecases.py` chạy trên máy không có Excel (đường thuần Python).
"CHƯA CHẠY ĐƯỢC" = cần Windows + Excel Desktop thật, chưa từng được ghi đạt.

| # | Use case | Kỳ vọng | Trạng thái |
|---|---|---|---|
| UC-01 | CSV UTF-16 có BOM | Đọc được qua inspect | Đã chạy — đạt |
| UC-02 | CSV phân cách `;` | Nhận đúng 2 cột | Đã chạy — đạt |
| UC-03 | CSV phân cách tab | Nhận đúng 2 cột | Đã chạy — đạt |
| UC-04 | CSV có banner hành chính + tự tìm dòng tiêu đề | Bỏ qua banner, lấy đúng header | Đã chạy — đạt (sau khi vá VĐ-12) |
| UC-05 | CSV 61 cột | Bị từ chối (giới hạn 60) | Đã chạy — đạt |
| UC-06 | CSV đúng 20.000 dòng | Được chấp nhận | Đã chạy — đạt |
| UC-07 | detect_header_row trên rows có banner | Trả về dòng header thật | Đã chạy — đạt |
| UC-08 | Cột trùng tên sau khi cắt khoảng trắng ('ID' vs ' ID ') | Bị từ chối | Đã chạy — đạt |
| UC-09 | Dòng tiêu đề vượt quá dữ liệu | Bị từ chối | Đã chạy — đạt |
| UC-10 | auto_heal với cột trùng/cột khuyết | Tên cột thành duy nhất, đủ số cột | Đã chạy — đạt |
| UC-11 | Khóa thừa khoảng trắng + khác hoa thường | Vẫn khớp khi đối chiếu | Đã chạy — đạt |
| UC-12 | Khóa NFD khớp khóa NFC | Khớp sau chuẩn hóa | Đã chạy — đạt |
| UC-13 | Số 123 (General) vs chuỗi '00123' | KHÔNG khớp | Đã chạy — đạt |
| UC-14 | Số 123 định dạng 00000 vs chuỗi '00123' | Khớp | Đã chạy — đạt |
| UC-15 | Khóa nguồn trống | Đếm riêng blank_keys, không tính khớp | Đã chạy — đạt |
| UC-16 | Khóa chứa ký tự `*` | Hiểu theo nghĩa đen, khớp chính xác | Đã chạy — đạt |
| UC-17 | Lọc theo từ khóa chức vụ (nhiều từ khóa) | Đúng số dòng | Đã chạy — đạt |
| UC-18 | Lọc điều kiện số >= 8 | Đúng số dòng | Đã chạy — đạt |
| UC-19 | Lọc "khác rỗng" | Mọi dòng có giá trị | Đã chạy — đạt |
| UC-20 | Tổng tiền VN '1.234.567' + '2.500.000' (lenient) | Tổng Decimal đúng 3.734.667 | Đã chạy — đạt |
| UC-20b | '500.000' một dấu chấm | Hiểu là thập phân 500 (quy ước hiện tại, xem ghi chú) | Đã chạy — đạt, ghi nhận mơ hồ |
| UC-21 | Số dạng '1,234' ở chế độ chặt | Bị từ chối | Đã chạy — đạt |
| UC-22 | Tổng vượt 10^14 | Bị từ chối | Đã chạy — đạt |
| UC-23 | Tổng có số âm | Cộng đúng | Đã chạy — đạt |
| UC-24 | Ô số trống khi tổng hợp | Bị từ chối, không tự đổi thành 0 | Đã chạy — đạt |
| UC-25 | Khóa serial ngày 45000, allow_dates=True | Đổi thành '2023-03-15' | Đã chạy — đạt |
| UC-26 | Khóa serial ngày, allow_dates=False | Bị từ chối kèm thông báo rõ | Đã chạy — đạt |
| UC-27 | Clean: cắt khoảng trắng + chuẩn NFC | Đúng 1 ô thay đổi, đúng nội dung | Đã chạy — đạt |
| UC-28 | execute với token chấp nhận sai | Bị chặn TRƯỚC khi gọi Excel | Đã chạy — đạt |
| UC-29 | preflight workbook sạch | Được chấp nhận | Đã chạy — đạt |
| UC-30 | Workbook chứa vbaProject (macro) | preflight từ chối | Đã chạy — đạt |
| UC-31 | Workbook chứa externalLink | preflight từ chối | Đã chạy — đạt |
| UC-32 | XLSX qua Excel thật: inspect → preview → execute → đối soát → XLSX/PDF | Toàn luồng | CHƯA CHẠY ĐƯỢC — cần Windows (test_app --integration) |
| UC-33 | File đang mở trong Excel / bị khóa ghi | Báo lỗi dễ hiểu | CHƯA CHẠY ĐƯỢC — cần Windows |
| UC-34 | File bị tráo giữa xem trước và thực thi (đường Excel thật) | Bị chặn bởi hash/token | Một phần: bản mock đã chạy (test_read_boundary); đường thật CHƯA CHẠY ĐƯỢC |
| UC-35 | Ô lỗi Excel (#DIV/0!) sau tính lại | Bị từ chối | CHƯA CHẠY ĐƯỢC — cần Excel thật (có trong integration) |
| UC-36 | Xuất PDF và kiểm tra bố cục | PDF đúng bố cục | CHƯA CHẠY ĐƯỢC — cần Windows + kiểm tra trực quan |

Ghi chú UC-20b: '500.000' mơ hồ giữa "500 nghìn đồng" (quen dùng trong văn
bản tiền VN) và "500,000" thập phân. Parser lenient hiện chọn nghĩa thập
phân cho mẫu MỘT dấu chấm; mẫu nhiều dấu chấm (1.234.567) được hiểu là phân
nhóm nghìn. Đây là lựa chọn thiết kế đã có từ trước, được ghi nhận ở đây để
người dùng/lượt sau quyết định có đổi hay không — vòng 11 không tự ý đổi
hành vi số học.
