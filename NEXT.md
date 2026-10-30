# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 30 commit (Bước 0 + Vòng 1–29), CHƯA push —
  người dùng TỰ push, Muse không push.
- Chế độ không cạn việc ĐANG BẬT theo chỉ đạo người dùng 2026-10-30.
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.
- VĐ đã phát hiện tới nay: VĐ-01…18 (VĐ-18 vá ở Vòng 28).

## Vòng tiếp theo: Vòng 30 — TỔNG KẾT 30 VÒNG (mốc phụ)
1. Chấm lại 8 hạng mục theo thang Bước 0, có đối chiếu trước/sau.
2. Đối chiếu bảng VĐ-01…18: trạng thái từng mục (đã vá / đã ghim test /
   còn chờ Windows).
3. Đếm lại tổng số check từ output chạy thật của toàn bộ 23 file test.
4. Liệt kê việc Linux còn lại kèm đánh giá giá trị giảm dần + lý do,
   để người dùng quyết định dừng hay đi tiếp (Vòng 31+: đọc tĩnh
   excel_bridge.ps1; kiểm tra examples/ khớp hành vi hiện tại).

## NGHIỆM THU WINDOWS (chờ người dùng)
Danh sách đầy đủ trong WINDOWS_ACCEPTANCE.md của gói bàn giao và bản
trước của file này: 17+ lệnh test theo 3 nhóm (không Excel / DPAPI /
Excel thật với số đạt kỳ vọng), 8 mục kiểm tra thủ công, và phần cần báo
lại. Sau khi có kết quả: ghi vào UPGRADE_LOG, sửa lỗi Windows trước,
rồi mới xét tách gui().

## Điều kiện dừng vòng lặp
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện).
