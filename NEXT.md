# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái: mốc 20 vòng + Vòng 21 (mutation đợt 3) + Vòng 22 (fuzz đợt 2) đã xong

Nhánh `muse-upgrade`: 23 commit (Bước 0 + Vòng 1–22), CHƯA push.
Gói bàn giao cho nghiệm thu Windows đã xuất kèm vòng 22 (git bundle + zip
+ bản hướng dẫn WINDOWS_ACCEPTANCE.md) — nằm ngoài repo, người dùng tải về.

## ĐIỀU KIỆN CHẶN (chỉ đạo người dùng 2026-10-30)
- **KHÔNG tách gui() cho tới khi người dùng báo kết quả test trên Windows.**
- Thứ tự khi có kết quả Windows: (1) ghi kết quả vào UPGRADE_LOG (mục
  nghiệm thu Windows: integration, UC-32/33/35/36, benchmark, tooltip,
  hộp lỗi batch, DPAPI); (2) sửa các lỗi phát hiện trên Windows trước;
  (3) sau đó mới xét tách gui() theo panel.

## Việc làm được từ Linux trong lúc chờ (giá trị giảm dần)
- Rà soát các `except Exception` còn giữ ở VĐ-05 khi có tín hiệu lỗi mới.
- Chuẩn bị sẵn bản nháp tách gui() dạng đọc-hiểu (sơ đồ panel/luồng sự
  kiện) mà KHÔNG sửa code — chỉ khi người dùng yêu cầu.

## Câu hỏi chờ người dùng quyết định (đã trình bày đầy đủ phương án ở lượt 2026-10-30)
1. Chính sách động cơ mặc định: giữ `smart_local` (đề xuất) / cloud-first /
   local-only tuyệt đối.
2. '500.000' trong chế độ lenient: giữ nghĩa thập phân (đề xuất A) / hiểu
   là phân nhóm nghìn kiểu VN có điều kiện (B) / theo ngữ cảnh tên cột (C).

## Câu hỏi khác
- Push nhánh lên GitHub: chưa được xác nhận; thay vào đó người dùng đã
  nhận gói bundle/zip để tự nghiệm thu (lượt 2026-10-30).

## Điều kiện dừng
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện).
