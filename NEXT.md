# NEXT — bàn giao sau Chu kỳ 6

## TÓM TẮT CHO CHỦ SẢN PHẨM
- Đã xong **Bước 0 + Q1 + Q2 + Q3** trên nhánh `muse-upgrade`; **không push**, không có `STOP.txt`.
- Test đầy đủ cuối: **30 lệnh, 0 lỗi, 292,0 giây, exit 0**; integration Excel **105/105**.
- PrivacyGuard: **212 câu** (125 PII + 87 thường), **0 rò lọt ngoài 2 giới hạn đã ghi**, **0% che nhầm**.
- Trước khi gửi cloud, app hiện đúng đoạn chữ + tên cột đã che, có nút Cancel; chế độ local không hiện.
- Hiệu năng 20.000 dòng (trung vị 5 lần): đọc nguồn 5,07s, đọc danh mục 4,81s, đối chiếu 19,89s, tổng hợp 16,80s, làm sạch 11,99s.
- Có 2 lần Excel treo tạm thời 253–262s; đã ghi rõ và dùng trung vị, không giấu.
- Gói bàn giao: `handoff/excel-local-muse-upgrade-cycle6.bundle`, `.zip`, `muse-upgrade-cycle6.patch`, `HANDOFF.md`.
- **CẦN NGƯỜI XÁC NHẬN BẰNG MẮT:** giao diện tiếng Việt, hộp xem trước cloud, bố cục 100%/150% DPI.
- **Chờ quyết định:** có thêm `LICENSE` không (giấy phép nào) và có push/tag không. Không tự làm.
- Còn **Q4–Q7**; kế hoạch chi tiết ở dưới.

## Trạng thái
- Nhánh `muse-upgrade`, commit code cuối `0160c6a`, commit tài liệu bàn giao `f98de98`.
- Bundle bảo chứng tới `0160c6a` (`git bundle verify`: complete history, mọi thứ ngoài `handoff/`).
- Hành vi gần nhất v0.3.2 (tag cục bộ `v0.3.2`); refactor không đổi số phiên bản.
- `config.json` chứa khóa DPAPI thật của chủ sản phẩm, đã gitignore; **không đọc/in/commit**; provider test là `heuristic`.

## Bằng chứng chính
- `results/test-results.json`: 105/105 integration.
- `results/benchmark-q3-summary.json`: trung vị 5 lần Python/5k/20k, ngoại lệ Excel.
- `.work/qa-logs/run-20261102-104045.log`: bảng 30 lệnh, 0 lỗi.
- `results/secret-scan.json`: 0 finding ngoài allowlist (15 working tree/12 history, mẫu tổng hợp).
- `cycles/CYCLE-04.md` (Q1), `CYCLE-05.md` (Q2), `CYCLE-06.md` (Q3); `UPGRADE_LOG.md`.
- Ảnh GUI: **chưa có** — sẽ làm ở Q5 vào `docs/screens/`.

## Việc tiếp theo (thứ tự)
1. **Q4 — dữ liệu giống thật:** tạo `examples/real_like/` ≥6 loại (sinh viên, nhân sự, kho, kế toán, điểm danh, lương, bán hàng…), ≥50 câu tiếng Việt; chạy đọc → đề xuất → xem trước → thực thi → đối soát; đo tỷ lệ kế hoạch đúng; thêm ca sai vào `test_infer_golden.py`.
2. **Q5 — ảnh GUI:** chụp màn hình vào `docs/screens/`, chỉ sửa lỗi đo được; ghi **CẦN NGƯỜI XÁC NHẬN BẰNG MẮT**.
3. **Q6 — tách tiếp:** `App.__init__`, `_filter_plan`, `gui()` nếu Q1 thành công; vàng + mutation + test Excel thật.
4. **Q7 — phát hành:** chỉ tạo `.github/workflows/tests.yml`, **không push**; cập nhật README/RELEASES, tăng phiên bản, tag cục bộ; **không tự thêm LICENSE**.

## Câu hỏi chờ chủ sản phẩm
- Có thêm `LICENSE` không? Nếu có, chọn giấy phép nào (MIT/BSD/Apache/GPL/khác)?
- Có muốn push nhánh `muse-upgrade` hoặc tạo tag phát hành trên remote không?
- Xác nhận giao diện và hộp xem trước cloud bằng mắt; ghi đạt/không đạt vào `WINDOWS_ACCEPTANCE.md`.

## Quy tắc đang giữ
- Chỉ làm trên `muse-upgrade`; không push; không đụng nhánh chính.
- Không dùng file thật; chỉ `examples/`/dữ liệu tự sinh; test không gọi cloud.
- Không bịa kết quả: mục không chạy được ghi **CHƯA KIỂM CHỨNG**, mục cần mắt ghi **CẦN NGƯỜI XÁC NHẬN**.
- Mỗi bước nhỏ một commit; dừng hẳn chỉ khi có `STOP.txt`.
