# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 27 commit (Bước 0 + Vòng 1–26), CHƯA push —
  người dùng TỰ push, Muse không push.
- Chế độ không cạn việc ĐANG BẬT theo chỉ đạo người dùng 2026-10-30:
  tiếp tục các vòng Linux cho tới khi có STOP.txt hoặc chứng minh được
  không còn việc có giá trị (nêu lý do từng mục).
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.

## Vòng tiếp theo: Vòng 27 — Rà soát tĩnh đợt 2 + đối chiếu README ↔ code
1. pyflakes toàn bộ file .py + py_compile; xử lý mọi cảnh báo mới.
2. Đối chiếu README ↔ code: bổ sung 3 file test còn thiếu vào mục Kiểm
   tra (test_fuzz_headers, test_policy_defaults, test_safety_limits,
   test_fuzz_xlsx), kiểm mọi con số đếm còn lại trong README (số bài
   test từng file so với output chạy thật vòng này).
3. Quét các hằng số giới hạn bị lặp giữa app.py/ai_client.py (2000 ký
   tự, 8000 ký tự prompt...) — chỉ ghi nhận, không refactor khi không
   có lỗi.

## Hàng đợi Linux sau đó (theo thứ tự)
- Vòng 28: use case bất thường đợt 2 (sheet tên lạ/rất dài, CSV chỉ có
  header, file nguồn = file tham chiếu cùng một file, filter với
  selected_columns không tồn tại, lookup khóa trùng + thiếu đồng thời,
  clean trên cột số, summary nhóm toàn khóa trống...).
- Vòng 29: mutation đợt 6 cho phần còn lại (formula_preflight, read_xml
  DTD đã có M3-5 ở tầng preflight — thử ở tầng read_xml trực tiếp,
  execute_batch phần thuần logic nếu có điểm kiểm tra được).
- Vòng 30: tổng kết 30 vòng (mốc phụ).

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
