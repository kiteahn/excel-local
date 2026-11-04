# NEXT — bàn giao sau Chu kỳ 9

## TÓM TẮT CHO CHỦ SẢN PHẨM
- Đã xong **Bước 0 + Q1–Q6** trên nhánh `muse-upgrade`; **không push**, không có `STOP.txt`.
- Q4: 9 bảng giống thật, 55 câu; **53/55 = 96,4%** kế hoạch đúng; 49 lượt Excel đạt, 1 từ chối đúng vì `#N/A`, 0 lỗi.
- Q5: 3 ảnh giao diện trong `docs/screens/`; **CẦN NGƯỜI XÁC NHẬN BẰNG MẮT** ở 100%/150% DPI.
- Q6: tách `_filter_plan` 226→121 dòng, `gui()` 713→23, `App.__init__` 213→16; 34 ca vàng, mutation **6/6** bị bắt.
- Test đầy đủ: **31 lệnh, 0 lỗi, 1018,7 giây, exit 0**; integration Excel **105/105**; `examples/` sạch.
- Có một khoảng trễ ~4 phút giữa hai workbook trong `test_real_like_plans` (703,6s); chưa rõ nguyên nhân, đã ghi, không giấu.
- Gói bàn giao Chu kỳ 9: `handoff/excel-local-muse-upgrade-cycle9.bundle`, `.zip`, `muse-upgrade-cycle9.patch`, `HANDOFF.md`.
- Tiếp theo là **Q7**: chỉ tạo `.github/workflows/tests.yml` (không push), cập nhật README/RELEASES, tăng phiên bản, tag cục bộ.
- **Chờ quyết định:** có thêm `LICENSE` không (MIT/Apache-2.0/giữ riêng tư) và có push/tag remote không. Không tự làm.

## Trạng thái
- Nhánh `muse-upgrade`, commit code cuối `29de070` (Q6); commit bàn giao Chu kỳ 9 sẽ tạo sau tài liệu này.
- Hành vi gần nhất trước đợt này là v0.3.2 (tag `v0.3.2`); Q7 sẽ tăng lên v0.4.0 và tag cục bộ.
- `config.json` chứa khóa DPAPI thật, đã gitignore; **không đọc/in/commit**; test dùng heuristic/cấu hình rỗng.

## Bằng chứng chính
- Full suite: `.work/qa-logs/run-20261104-105727.log` — 31 lệnh, 0 lỗi, 1018,7s.
- `results/test-results.json`: 105/105 integration; `results/plan-accuracy-q4.json`: 53/55, 49 chạy đạt/1 từ chối đúng.
- `results/secret-scan.json`: 0 finding ngoài allowlist (sẽ quét lại sau khi commit Q7).
- `cycles/CYCLE-04.md` … `CYCLE-09.md`; `UPGRADE_LOG.md`; `README.md`.
- Ảnh GUI: `docs/screens/01-main.png`, `02-preview.png`, `03-settings.png`.

## Việc tiếp theo (Q7)
1. Tạo `.github/workflows/tests.yml`: 17 test không cần Excel trên Windows + Ubuntu; Windows thêm `test_router.py`, `test_security_privacy.py`. Không đưa 3 bộ cần Excel trên Windows vào CI.
2. Đồng bộ tiêu đề app/cửa sổ thành v0.4.0; cập nhật README/RELEASES/release notes.
3. Chạy lại test GUI/UI sau khi đổi tiêu đề; kiểm tra YAML bằng PyYAML.
4. Tạo tag cục bộ `v0.4.0`; **không push**; **không thêm LICENSE**.
5. Cập nhật `WINDOWS_ACCEPTANCE.md`, xuất handoff cuối, cập nhật lại mục TÓM TẮT này.

## Câu hỏi chờ chủ sản phẩm
- Có thêm `LICENSE` không? Nếu có, chọn MIT/Apache-2.0/BSD/GPL/khác? (Đề xuất: MIT nếu muốn người khác dùng lại dễ; giữ riêng tư nếu chưa muốn.)
- Có muốn push nhánh `muse-upgrade` hoặc tạo tag phát hành trên remote không? Hiện mọi thứ chỉ ở máy.
- Xác nhận giao diện và hộp xem trước cloud bằng mắt; ghi đạt/không đạt vào `WINDOWS_ACCEPTANCE.md`.

## Quy tắc đang giữ
- Chỉ làm trên `muse-upgrade`; không push; không đụng nhánh chính.
- Không dùng file thật; chỉ `examples/`/dữ liệu tự sinh; test không gọi cloud.
- Không bịa kết quả: mục không chạy được ghi **CHƯA KIỂM CHỨNG**, mục cần mắt ghi **CẦN NGƯỜI XÁC NHẬN**.
- Mỗi bước nhỏ một commit; dừng hẳn chỉ khi có `STOP.txt`.
