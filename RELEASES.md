# RELEASES — các phiên bản đã phát hành (cục bộ, không push)

| Phiên bản | Ngày | Nội dung chính | Tài liệu |
|-----------|------|----------------|----------|
| 0.1 | trước 2026-10-30 | Bản đầu: đọc bảng tính, đối chiếu/tổng hợp/làm sạch, xem trước + token | `results/test-results-v0.1.json` |
| 0.2 | trước 2026-10-30 | Sửa lỗi biên, thêm kiểm tra an toàn, biên bản JSON/PDF | `release-notes-0.2.md`, `results/acceptance-v0.2.json` |
| 0.3 | 2026-10 | Lọc & trích xuất, tự chữa lành ô gộp, Smart Router | `release-notes-0.3.md` |
| 0.3.1 | 2026-10-30 | Ổn định Windows, DPAPI, chính sách mặc định Ollama | `release-notes-0.3.1.md` |

## Chu kỳ không tăng phiên bản
- **Chu kỳ 0 (2026-10-31)** — nghiệm thu nền Windows, sửa 4 lỗi bố cục giao diện,
  thêm `test_ui_acceptance.py`, bổ sung hồ sơ nghiệm thu. Đây là bản vá nghiệm thu,
  vẫn mang số 0.3.1; lần tăng phiên bản kế tiếp gắn với thay đổi hành vi người dùng
  (dự kiến Chu kỳ 2: PrivacyGuard chính xác hơn).
