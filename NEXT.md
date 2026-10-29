# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 3 — VĐ-11: dpapi_protect không được âm thầm lưu khóa dạng thô

Vì sao ưu tiên cao: đây là lỗi bảo mật thật vừa phát hiện ở Vòng 2 — nếu DPAPI
thất bại trên Windows, API key của người dùng nằm plaintext trong config.json
mà không hề có cảnh báo, trái với cam kết mã hóa.

Làm:
1. Đọc kỹ `dpapi_protect`/`dpapi_unprotect`/`save_config`/`get_config` trong
   ai_client.py (~dòng 140–302).
2. Thiết kế tối thiểu, không đổi hành vi trên Windows khi DPAPI chạy tốt:
   - `save_config` không ghi khóa thô: nếu `dpapi_protect` không trả về chuỗi
     `dpapi:` cho một khóa khác rỗng thì DỪNG lưu và báo lỗi rõ (raise) thay vì
     ghi plaintext. Kiểm tra nơi GUI gọi save_config để lỗi này hiển thị được
     cho người dùng (không sập app).
   - `get_config`: nếu gặp khóa lưu thô (legacy) thì vẫn đọc được như hiện tại
     (tương thích ngược), không tự ý xóa.
3. Test: bổ sung vào test_router.py hoặc test_security_privacy.py trường hợp
   monkeypatch `dpapi_protect` trả plaintext → `save_config` phải raise và
   file config không chứa khóa. Chạy lại trên Linux: test_router (6 mục),
   test_app (79), test_all_edge_cases (61).
4. Commit, cập nhật UPGRADE_LOG.md, chuyển NEXT sang Vòng 4 (skip theo nền
   tảng cho các test Windows-only — VĐ-04).

## Hàng đợi sau đó
- Vòng 4: VĐ-04 skip rõ ràng cho test cần Excel/DPAPI/display.
- Vòng 5: Mục E — bộ test PrivacyGuard ≥50 mẫu, đo lọt/che nhầm.
- Vòng 6+: VĐ-05 (except hẹp), VĐ-06 (tách hàm), VĐ-07 (đo hiệu năng),
  VĐ-09 (dọn repo), VĐ-10 (giao diện nhẹ).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); Vòng 1 chỉ sửa
  tài liệu cho đúng code, chưa đổi chính sách.
