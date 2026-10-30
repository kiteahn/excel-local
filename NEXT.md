# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 31 commit (Bước 0 + Vòng 1–30), CHƯA push —
  người dùng TỰ push, Muse không push.
- Chế độ không cạn việc ĐANG BẬT theo chỉ đạo người dùng 2026-10-30.
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.
- VĐ đã phát hiện tới nay: VĐ-01…18 (VĐ-18 vá ở Vòng 28).

## Vòng tiếp theo: Vòng 31 — Đọc tĩnh excel_bridge.ps1 + kiểm tra examples/
1. Đọc toàn bộ excel_bridge.ps1 (không chạy được PowerShell trên VM):
   tìm lỗi cú pháp hiển nhiên, tên biến dùng trước khi gán, đường dẫn
   cứng, thông điệp lỗi khó hiểu; chỉ sửa lỗi CHẮC CHẮN, còn lại ghi
   nhận vào UPGRADE_LOG để đối chiếu khi nghiệm thu Windows.
2. Chạy create_diverse_examples.py vào thư mục tạm, inspect thử 2–3
   file ví dụ bằng đường CSV/fast-reader để chắc examples chưa thối
   so với hành vi hiện tại.
3. Sau vòng này: nếu không phát hiện việc giá trị mới, vòng lặp chuyển
   sang trạng thái chờ (nghiệm thu Windows / chỉ đạo mới / STOP.txt),
   theo đúng tiêu chí dừng người dùng đã đặt.

## NGHIỆM THU WINDOWS (chờ người dùng)
Danh sách đầy đủ trong WINDOWS_ACCEPTANCE.md của gói bàn giao và bản
trước của file này: 17+ lệnh test theo 3 nhóm (không Excel / DPAPI /
Excel thật với số đạt kỳ vọng), 8 mục kiểm tra thủ công, và phần cần báo
lại. Sau khi có kết quả: ghi vào UPGRADE_LOG, sửa lỗi Windows trước,
rồi mới xét tách gui().

## Điều kiện dừng vòng lặp
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện).
