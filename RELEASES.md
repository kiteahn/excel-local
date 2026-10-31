# RELEASES — các phiên bản đã phát hành (cục bộ, không push)

| Phiên bản | Ngày | Nội dung chính | Tài liệu |
|-----------|------|----------------|----------|
| 0.1 | trước 2026-10-30 | Bản đầu: đọc bảng tính, đối chiếu/tổng hợp/làm sạch, xem trước + token | `results/test-results-v0.1.json` |
| 0.2 | trước 2026-10-30 | Sửa lỗi biên, thêm kiểm tra an toàn, biên bản JSON/PDF | `release-notes-0.2.md`, `results/acceptance-v0.2.json` |
| 0.3 | 2026-10 | Lọc & trích xuất, tự chữa lành ô gộp, Smart Router | `release-notes-0.3.md` |
| 0.3.1 | 2026-10-30 | Ổn định Windows, DPAPI, chính sách mặc định Ollama | `release-notes-0.3.1.md` |
| 0.3.2 | 2026-10-31 | Privacy Guard: bịt lọt tên đứng một mình, bỏ che nhầm số tiền/MSSV/mã 13 số, sửa ăn khoảng trắng; bộ đo 62 PII + 37 non-PII | `release-notes-0.3.2.md` |

## Chu kỳ không tăng phiên bản
- **Chu kỳ 0 (2026-10-31)** — nghiệm thu nền Windows, sửa 4 lỗi bố cục giao diện,
  thêm `test_ui_acceptance.py`, bổ sung hồ sơ nghiệm thu. Đây là bản vá nghiệm thu,
  vẫn mang số 0.3.1; lần tăng phiên bản kế tiếp gắn với thay đổi hành vi người dùng
  (Chu kỳ 1 → 0.3.2: PrivacyGuard chính xác hơn).
- **Chu kỳ 2 (2026-10-31)** — tách `open_ai_settings()` thành
  `open_ai_settings_dialog()` cấp module, không đổi giao diện/hành vi. Refactor
  thuần túy nên giữ số 0.3.2; bằng chứng trong `cycles/CYCLE-02.md`.
