# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 27 commit (Bước 0 + Vòng 1–26), CHƯA push —
  người dùng TỰ push, Muse không push.
- Chế độ không cạn việc ĐANG BẬT theo chỉ đạo người dùng 2026-10-30:
  tiếp tục các vòng Linux cho tới khi có STOP.txt hoặc chứng minh được
  không còn việc có giá trị (nêu lý do từng mục).
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.

## Vòng tiếp theo: Vòng 28 — Use case bất thường đợt 2
Chạy qua prepare_sheets/validate_plan/inspect (phần thuần Python):
sheet tên rất dài/ký tự lạ trong snapshot; CSV chỉ có dòng header (0 dòng
dữ liệu → phải báo rõ); nguồn và tham chiếu là CÙNG một snapshot (tự đối
chiếu); filter với selected_columns chứa cột không tồn tại (validate phải
chặn); lookup đồng thời có khóa trùng ở tham chiếu và khóa thiếu ở nguồn
(stats duplicate/missing đếm đúng); clean trên cột toàn số (changed=0,
không phá giá trị); summary khi mọi khóa nhóm đều trống (phải từ chối);
filter điều kiện số trên cột chữ (phải từ chối rõ ràng). Ca nào lộ lỗi
thật → vá tối thiểu + regression; ca đúng thiết kế → ghim vào
test_usecases.py đợt 2 (tên tiền tố UC2-).

## Hàng đợi Linux sau đó (theo thứ tự)
- Vòng 29: mutation đợt 6 (formula_preflight qua read_xml trực tiếp,
  execute_batch phần kiểm tra đầu vào, bridge payload sanitize nếu có
  điểm kiểm tra thuần Python).
- Vòng 30: tổng kết 30 vòng (mốc phụ): chấm lại nhanh, đối chiếu VĐ,
  đếm lại tổng số check.

## NGHIỆM THU WINDOWS (chờ người dùng — danh sách đầy đủ ở bản trước của
## file này và trong WINDOWS_ACCEPTANCE.md của gói bàn giao)
Tóm tắt: Bước 1 chạy 17 lệnh test không cần Excel (số đạt kỳ vọng đã ghi
trong WINDOWS_ACCEPTANCE.md, cộng thêm test_policy_defaults 12,
test_safety_limits 6, test_fuzz_headers 15); Bước 2 test_router +
test_security_privacy với DPAPI thật (kỳ vọng 7/7 và 43/43); Bước 3
test_app.py --integration (97 bài, Excel thật), test_gui.py, benchmark.py;
Bước 4 tám mục kiểm tra thủ công (luồng UC-32, file bị khóa UC-33, ô lỗi
UC-35, PDF UC-36, tooltip/Ctrl+Enter, cài đặt + DPAPI, hộp lỗi batch
VĐ-17, số học vn_dot_thousands); Bước 5 báo lại output + kết quả để ghi
vào UPGRADE_LOG trước khi xét tách gui().

## Điều kiện dừng vòng lặp
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện).
