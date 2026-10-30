# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 32 commit (Bước 0 + Vòng 1–31), CHƯA push —
  người dùng TỰ push, Muse không push.
- **Vòng lặp ở trạng thái CHỜ** (kết luận Vòng 30–31): việc Linux có
  giá trị đã làm hết — 23 file test/499 check, mutation 30/30, fuzz 3
  mảng, rà soát tĩnh 2 đợt, đọc tĩnh cả bridge PowerShell, đối chiếu
  README/examples. Việc còn lại trên Linux chỉ là giá trị thấp (mutation
  đợt 7 cho nhánh infer phụ; fuzz CSV seed dài hạn) — không tự chạy tiếp
  trừ khi người dùng yêu cầu.
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.

## Việc đang chờ (theo thứ tự ưu tiên)
1. **Nghiệm thu Windows bởi người dùng** — danh sách đầy đủ trong
   WINDOWS_ACCEPTANCE.md của gói bàn giao (đã xuất lại ở commit mới
   nhất): 3 nhóm lệnh test (không Excel / DPAPI / Excel thật, kèm số
   đạt kỳ vọng), 8 mục kiểm tra thủ công — gồm cả điểm W-1 của Vòng 31
   (nhìn danh sách sheet trong file kết quả của UC-32 xem có sheet
   trống mặc định không). Báo lại output + mục nào không đạt.
2. Sau khi có kết quả Windows: ghi vào UPGRADE_LOG mục "Nghiệm thu
   Windows", sửa lỗi phát hiện TRƯỚC, rồi mới xét tách gui().
3. Người dùng tự push nhánh muse-upgrade (quyết định 2026-10-30).

## Điều kiện dừng vòng lặp
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện). Vòng lặp hiện chờ việc ở mục trên, không tự
  sinh vòng mới.
