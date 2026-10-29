# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái: đã hoàn thành mốc TỔNG KẾT 20 VÒNG (xem UPGRADE_LOG.md)

Nhánh `muse-upgrade`: 22 commit (Bước 0 + Vòng 1–21), CHƯA push.

## Hướng tiếp theo nếu chạy tiếp (theo thứ tự giá trị)

1. **Nghiệm thu Windows (giá trị cao nhất, cần máy Windows của người dùng):**
   chạy trên máy có Excel — test_app.py --integration, UC-32/33/35/36 trong
   USECASES.md, benchmark.py đo lại, test_gui.py + nhìn trực tiếp tooltip
   (Vòng 10) và hộp lỗi batch (VĐ-17), test_security_privacy mục DPAPI.
   Kết quả (đạt/không) cần được ghi bổ sung vào UPGRADE_LOG.
2. **Tách gui() có kiểm chứng** — chỉ sau khi có test_gui chạy thật ở mục 1:
   tách theo panel (nguồn/tham chiếu, đề xuất, xem trước, batch, settings)
   với test vàng hành vi nếu khả thi.
3. **Vòng lặp bảo trì nhẹ từ Linux** (làm được ngay, giá trị giảm dần):
   mutation đợt 3 cho app.py ĐÃ XONG ở Vòng 21 (5/5 sau khi thêm SEC-01/02);
   còn lại: fuzz đợt 2 cho detect_header_row/select_header trên snapshot
   dict; rà soát các except giữ lại ở VĐ-05 khi có tín hiệu lỗi mới.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Push nhánh `muse-upgrade` (22 commit) lên GitHub chưa? Repo công khai
  nên cần xác nhận rõ ràng.
- Chính sách mặc định: Smart Router (có cloud) hay Ollama-only?
- '500.000' một dấu chấm = 500 nghìn hay 500 thập phân? (USECASES UC-20b)

## Điều kiện dừng
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện).
