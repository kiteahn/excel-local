# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 26 commit (Bước 0 + Vòng 1–25), CHƯA push —
  người dùng TỰ push, Muse không push.
- Chế độ không cạn việc ĐANG BẬT theo chỉ đạo người dùng 2026-10-30:
  tiếp tục các vòng Linux cho tới khi có STOP.txt hoặc chứng minh được
  không còn việc có giá trị (nêu lý do từng mục).
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.

## Vòng tiếp theo: Vòng 26 — Fuzz đường đọc XLSX thuần Python (load_xlsx_fast)
1. Fuzz `load_xlsx_fast` (đường lenient đọc XLSX không cần Excel): zip
   hỏng/không phải zip, thiếu workbook.xml/sharedStrings, shared string
   index vượt phạm vi, ô kiểu lạ (inlineStr, lỗi #DIV/0!, boolean), sheet
   rỗng, dimension khai khống, XML có DTD (phải bị read_xml chặn), file
   XLSX hợp lệ do openpyxl sinh để đối chiếu với inspect lenient.
   Hợp đồng: đọc đúng hoặc ValueError/None rõ ràng — không traceback lạ.
2. Fuzz `read_xml` trực tiếp: XML lồng sâu, namespace lạ, entity khai báo.
3. Lỗi thật (nếu có) → vá tối thiểu + regression test trong cùng vòng.

## Hàng đợi Linux sau đó (theo thứ tự)
- Vòng 27: rà soát tĩnh đợt 2 (pyflakes toàn bộ + py_compile) và đối
  chiếu README ↔ code lần nữa — gồm danh sách test trong README (đã có
  thêm test_policy_defaults/test_safety_limits/test_fuzz_headers cần bổ
  sung vào mục Kiểm tra) và mọi con số đếm trong README.
- Vòng 28: use case bất thường đợt 2 (sheet tên lạ/rất dài, CSV chỉ có
  header, file nguồn = file tham chiếu cùng một file, filter với
  selected_columns không tồn tại, lookup khóa trùng + thiếu đồng thời,
  clean trên cột số, summary nhóm toàn khóa trống...).
- Vòng 29: mutation đợt 6 cho phần còn lại (execute_batch phần thuần
  logic nếu tách được điểm kiểm tra, formula_preflight, read_xml DTD).
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
