# NEXT — bàn giao sau Chu kỳ 10 (Q7)

## TÓM TẮT CHO CHỦ SẢN PHẨM
- Đã xong **Bước 0 + Q1–Q7** trên nhánh `muse-upgrade`; **không push**, không có `STOP.txt`.
- Bản phát hành **v0.4.0**: đồng bộ tiêu đề app, cập nhật README/RELEASES, tag cục bộ `v0.4.0`.
- CI mới `.github/workflows/tests.yml`: 17 test không cần Excel trên Windows + Ubuntu; Windows thêm 2 bộ DPAPI/bảo mật.
- Q4: 9 bảng giống thật, **53/55 = 96,4%** kế hoạch đúng; 49 lượt Excel đạt, 1 từ chối đúng `#N/A`, 0 lỗi.
- Q6: `_filter_plan` 226→121 dòng, `gui()` 713→23, `App.__init__` 213→16; 34 ca vàng, mutation **6/6** bị bắt.
- Test đầy đủ: **31 lệnh, 0 lỗi, 1018,7 giây, exit 0**; integration **105/105**; `examples/` sạch.
- Có một khoảng trễ ~4 phút giữa hai workbook trong `test_real_like_plans` (703,6s); chưa rõ nguyên nhân, đã ghi.
- Gói bàn giao cuối: `handoff/excel-local-muse-upgrade-cycle10.bundle`, `.zip`, `muse-upgrade-cycle10.patch`, `HANDOFF.md`.
- **CẦN NGƯỜI XÁC NHẬN BẰNG MẮT:** giao diện, hộp xem trước cloud, bố cục 100%/150% DPI (ảnh hiện ở 125%).
- **CHƯA KIỂM CHỨNG:** một lượt GitHub Actions thật trên Ubuntu; mới kiểm cú pháp YAML và chạy bộ chọn trên Windows.
- **Chờ quyết định:** có thêm `LICENSE` không (MIT/Apache-2.0/giữ riêng tư) và có push/tag remote không. Không tự làm.

## Trạng thái
- Nhánh `muse-upgrade`, tag cục bộ `v0.4.0`; không push, không có tag remote.
- `config.json` chứa khóa DPAPI thật, đã gitignore; **không đọc/in/commit**; test dùng heuristic/cấu hình rỗng.
- `WINDOWS_ACCEPTANCE.md` là gói nghiệm thu cuối; `handoff/HANDOFF.md` mô tả gói cycle10.

## Bằng chứng chính
- Full suite: `.work/qa-logs/run-20261104-105727.log` — 31 lệnh, 0 lỗi, 1018,7s, exit 0.
- `results/test-results.json`: 105/105 integration; `results/plan-accuracy-q4.json`: 53/55, 49 đạt/1 từ chối đúng/0 lỗi.
- `results/secret-scan.json`: 0 finding ngoài allowlist sau khi commit Q7.
- `cycles/CYCLE-04.md` … `CYCLE-10.md`; `UPGRADE_LOG.md`; `README.md`; `RELEASES.md`.
- Ảnh GUI v0.4.0: `docs/screens/01-main.png`, `02-preview.png`, `03-settings.png`.

## Việc tiếp theo (vòng sau Q7)
1. Đọc lại mã như người mới, rà nợ kỹ thuật còn lại (M-4 `prepare_sheets` 177 dòng, M-5 `except Exception`, M-6 `_lookup_plan` 202 dòng).
2. Fuzz/mutation đợt mới; đo lại hiệu năng nếu đổi bridge/giao thức.
3. Chỉ thêm `LICENSE` và push/tag remote khi chủ sản phẩm quyết định.
4. Xác nhận giao diện/cloud preview bằng mắt; thử CI thật trên GitHub khi được phép push.

## Câu hỏi chờ chủ sản phẩm
- Có thêm `LICENSE` không? Nếu có, chọn MIT/Apache-2.0/BSD/GPL/khác? (Đề xuất: MIT nếu muốn người khác dùng lại dễ; giữ riêng tư nếu chưa muốn.)
- Có muốn push nhánh `muse-upgrade` hoặc tag `v0.4.0` lên remote không? Hiện mọi thứ chỉ ở máy.
- Xác nhận giao diện và hộp xem trước cloud bằng mắt; ghi đạt/không đạt vào `WINDOWS_ACCEPTANCE.md`.

## Quy tắc đang giữ
- Chỉ làm trên `muse-upgrade`; không push; không đụng nhánh chính.
- Không dùng file thật; chỉ `examples/`/dữ liệu tự sinh; test không gọi cloud.
- Không bịa kết quả: mục không chạy được ghi **CHƯA KIỂM CHỨNG**, mục cần mắt ghi **CẦN NGƯỜI XÁC NHẬN**.
- Mỗi bước nhỏ một commit; dừng hẳn chỉ khi có `STOP.txt`.
