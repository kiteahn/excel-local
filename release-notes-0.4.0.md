# Bản 0.4.0 · Nền tảng phát hành sau Bước 0 và Q1–Q7

Ngày: 2026-11-04 · Nhánh `muse-upgrade` (cục bộ, **không push**)

## Vì sao có bản này

0.3.2 là bản hành vi gần nhất trước đợt nâng cấp. Đợt này không chỉ sửa lỗi mà
còn hoàn thiện quy trình kiểm chứng, dữ liệu giống thật, ảnh giao diện và nền
tảng chạy test tự động. Vì vậy số phiên bản tăng từ 0.3.2 lên **0.4.0**.

## Đã làm

### Bước 0 — dọn nợ quy trình

- Test dùng thư mục tạm, không ghi đè `examples/`; có fingerprint tự phát hiện.
- Kết quả test tách theo chế độ: unit, integration, AI.
- Bridge Excel retry có giới hạn: chỉ retry timeout đã dọn được hoặc thiếu
  `result.json`, không retry lỗi Excel thật.
- Một lệnh `run_all_tests.ps1` chạy đủ nhóm, ghi log `.work/qa-logs/`.
- Quét mẫu bí mật/đường dẫn cá nhân trong file được git theo dõi và lịch sử.

### Q1 — tách nhánh lookup

- `prepare_sheets` 262 → 169 dòng; `_prepare_lookup` 106 dòng.
- 27 kịch bản vàng, kiểm chứng hai chiều với bản trước tách.
- Mutation 6/6 chỗ bị bắt; integration Excel 105/105.

### Q2 — PrivacyGuard và xem trước cloud

- Corpus **212 câu** (125 PII + 87 không phải PII).
- **0 rò lọt ngoài 2 giới hạn đã ghi** (tên 2 từ; tên dân tộc không bắt đầu
  bằng họ Kinh), **0% che nhầm** trên 87 câu thường (ngưỡng ≤2%).
- Vá lỗ `"Mật khẩu mới: ..."` không bị che.
- Trước khi gửi cloud, app hiện đúng câu đã che + tên cột đã che, có nút hủy;
  chế độ local/Ollama/heuristic không hiện hộp này.
- README ghi rõ giới hạn, không dùng "tuyệt đối"/"không bao giờ lọt".

### Q3 — hiệu năng

- 5 lần đo mỗi loại, lấy trung vị; lưu `results/benchmark-q3-summary.json`.
- 20.000 dòng: đọc nguồn 5,072s; đọc danh mục 4,810s; đối chiếu 19,886s;
  tổng hợp 16,797s; làm sạch 11,987s.
- Không giữ AutoFit/WrapText (~2,9%) hoặc tái dùng phiên Excel (~8%) vì cải
  thiện dưới 10% mà tăng phức tạp/rủi ro COM.
- Ghi rõ 2 lần Excel treo tạm thời 253–262s; không giấu.

### Q4 — dữ liệu giống thật

- 9 workbook tổng hợp trong `examples/real_like/`, 6 loại bảng nguồn; sinh lại
  bằng `create_real_like_examples.py`.
- 55 câu tiếng Việt: **53/55 = 96,4%** kế hoạch đúng ngay lần đầu.
- Chạy đầy đủ preview → chấp nhận → Excel → verify → preserve: **49 thành
  công, 1 từ chối đúng vì `#N/A`, 0 lỗi**.
- Sửa lỗi thật: ô rỗng `inlineStr`, dòng `TỔNG CỘNG`, chọn cột clean/lookup,
  giá trị categorical, `khác rỗng`, cột `Có mặt`.
- 2 ca còn sai là **AVG/trung bình chưa hỗ trợ**, đã ghim vào vàng kèm cách sửa.

### Q5 — ảnh giao diện

- 3 ảnh trong `docs/screens/`: màn hình chính, hộp kiểm tra trước khi chấp
  nhận, hộp Cài đặt AI. Chụp bằng ứng dụng Tk thật, dữ liệu tổng hợp, cấu hình
  rỗng (không có khóa thật).
- Sửa lỗi chụp callback; ghi nhận lệch tiêu đề `v0.3` / `v0.3.2` và đồng bộ ở
  bản 0.4.0.
- `test_gui.py` 14/14, `test_ui_acceptance.py` 48/48.

### Q6 — tách tiếp

- `_filter_plan` 226 → 121 dòng (3 helper 46/40/38).
- `gui()` 713 → 23 dòng; `App.__init__` 213 → 16 dòng; `_build_ui` 199 dòng.
- Bộ vàng 34 ca; mutation 6/6 bị bắt. Không đổi hành vi.
- Dừng tách lớp `App` còn 699 dòng vì rủi ro giao diện lớn hơn lợi ích.

### Q7 — nền tảng phát hành

- Thêm `.github/workflows/tests.yml` chạy 17 test không cần Excel trên Windows
  và Ubuntu; Windows chạy thêm 2 bộ kiểm tra bảo mật/DPAPI. Ba bộ cần Excel
  trên Windows (`test_headers`, `test_single_file_filter`,
  `test_diverse_scenarios_e2e`) được loại khỏi CI.
- **Không push**, **không tự thêm LICENSE**.
- Đồng bộ tiêu đề app/cửa sổ về `v0.4.0`; cập nhật README/RELEASES; tag cục bộ
  `v0.4.0`.

## Kiểm thử đầy đủ

- `run_all_tests.ps1` lượt cuối: **31 lệnh, 0 lỗi, 1018,7s, exit 0**;
  `examples/` sạch, không Excel mồ côi, không thư mục tạm mới. `test_real_like_plans`
  mất 703,6s, có khoảng trễ ~4 phút giữa hai workbook chưa rõ nguyên nhân; đã ghi.
- Integration Excel: **105/105**; read-boundary 4/4; GUI 14/14; UI acceptance
  48/48; usecases 37/37; usecases2 10/10; PrivacyGuard 212 câu đạt; cloud preview
  20/20; infer golden 34/34.
- `examples/` không bị test ghi đè; quét bí mật 0 finding ngoài allowlist.

## Giới hạn / chưa kiểm chứng

- **AVG/trung bình** chưa hỗ trợ; planner trả `None` thay vì clarify.
- Chưa kiểm chứng giao diện ở **100%/150% DPI** và hộp xem trước cloud bằng
  mắt; ảnh hiện có ở 125% DPI.
- Workflow Ubuntu mới được kiểm tra cú pháp YAML và chạy bộ chọn trên Windows;
  **CHƯA KIỂM CHỨNG** một lượt GitHub Actions thật trên Ubuntu.
- Chưa dùng file thật của người dùng; mọi số đo trên dữ liệu tổng hợp.
- **Không có LICENSE** cho tới khi chủ sản phẩm quyết định.
