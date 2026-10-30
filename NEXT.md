# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 29 commit (Bước 0 + Vòng 1–28), CHƯA push —
  người dùng TỰ push, Muse không push.
- Chế độ không cạn việc ĐANG BẬT theo chỉ đạo người dùng 2026-10-30.
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.
- VĐ đã phát hiện tới nay: VĐ-01…18 (VĐ-18 vá ở Vòng 28).

## Vòng tiếp theo: Vòng 29 — Mutation đợt 6
Mục tiêu còn lại chưa mutation:
1. formula_preflight: bỏ kiểm tra hàm nguy hiểm (ví dụ INDIRECT/WEBSERVICE)
   trong công thức — kiểm tra test nào bắt (security/preflight qua file
   XLSX chứa công thức lạ, dựng bằng cách sửa XML trong zip).
2. read_xml: bỏ chặn DTD ở tầng read_xml trực tiếp (M3-5 mới phủ tầng
   preflight) — test_fuzz_xlsx X-11 phải bắt.
3. execute_batch: đột biến ở khâu tổng hợp báo cáo (succeeded đếm sai)
   nếu kiểm tra được bằng mock bridge trên Linux; không được thì ghi rõ
   lý do để lại Windows.
4. Điều kiện lọc số trong prepare filter: đổi '>=' thành '>' ở nhánh so
   sánh — test_usecases UC-18/UC2-07b phải bắt.
Đột biến lọt → bổ sung test trong cùng vòng.

## Hàng đợi Linux sau đó
- Vòng 30: TỔNG KẾT 30 VÒNG (mốc phụ): chấm lại 8 hạng mục, đối chiếu
  VĐ-01…18, đếm lại tổng check từ output thật, liệt kê việc còn lại kèm
  lý do giá trị giảm dần (chuẩn bị cho khả năng người dùng dừng vòng lặp).
- Vòng 31+: rà soát excel_bridge.ps1 bằng đọc tĩnh (không chạy được
  PowerShell ở đây — chỉ đọc và ghi nhận, sửa cú pháp hiển nhiên nếu có);
  kiểm tra examples/ còn khớp hành vi hiện tại (chạy create_diverse_examples
  vào thư mục tạm và inspect thử vài file).

## NGHIỆM THU WINDOWS (chờ người dùng)
Danh sách đầy đủ trong WINDOWS_ACCEPTANCE.md của gói bàn giao và bản
trước của file này: 17+ lệnh test theo 3 nhóm (không Excel / DPAPI /
Excel thật với số đạt kỳ vọng), 8 mục kiểm tra thủ công, và phần cần báo
lại. Sau khi có kết quả: ghi vào UPGRADE_LOG, sửa lỗi Windows trước,
rồi mới xét tách gui().

## Điều kiện dừng vòng lặp
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện).
