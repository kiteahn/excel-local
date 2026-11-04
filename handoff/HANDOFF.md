# Bàn giao Chu kỳ 9 — Excel Local Expert (`muse-upgrade`)

Ngày: 2026-11-04. Nhánh `muse-upgrade`, **không push**. Không có `STOP.txt`
tại thời điểm xuất. Commit code Q6: `29de070`.

## Gói trong thư mục này
| File | Nội dung |
|---|---|
| `excel-local-muse-upgrade-cycle9.bundle` | Toàn bộ lịch sử nhánh `muse-upgrade`; đã `git bundle verify` ("complete history") |
| `excel-local-muse-upgrade-cycle9.zip` | Cây file tại đúng commit đầu nhánh, chỉ gồm file git theo dõi |
| `muse-upgrade-cycle9.patch` | Diff `28f5679..HEAD` so với bản gốc |
| `HANDOFF.md` | File này |

Lấy code:
```powershell
git clone handoff/excel-local-muse-upgrade-cycle9.bundle excel-local-check
cd excel-local-check
git checkout muse-upgrade
```
Hoặc giải nén `excel-local-muse-upgrade-cycle9.zip` vào thư mục mới.

## Đã xong từ sau Chu kỳ 6
- **Q4 (Chu kỳ 7):** 9 workbook tổng hợp, 6 loại bảng nguồn, 55 câu tiếng Việt;
  **53/55 = 96,4%** kế hoạch đúng. 49 lượt Excel đạt, 1 từ chối đúng vì `#N/A`,
  0 lỗi. Sửa ô rỗng `inlineStr`, dòng `TỔNG CỘNG`, chọn cột clean/lookup,
  categorical, `khác rỗng`, cột `Có mặt`. Hai ca AVG/trung bình vẫn chưa hỗ trợ.
- **Q5 (Chu kỳ 8):** 3 ảnh Tk thật trong `docs/screens/` ở DPI 125%, cấu hình
  rỗng; sửa lỗi callback chụp ảnh; ghi lệch tiêu đề `v0.3`/`v0.3.2`.
- **Q6 (Chu kỳ 9):** `_filter_plan` 226 → 121 dòng (3 helper), `gui()` 713 → 23,
  `App.__init__` 213 → 16, `_build_ui` 199. Vàng 34 ca; mutation 6/6 bị bắt;
  không đổi hành vi. Dừng tách lớp `App` còn 699 dòng vì rủi ro giao diện.

## Bằng chứng mới nhất
- Full suite: `.work/qa-logs/run-20261104-105727.log` — **31 lệnh, 0 lỗi,
  1018,7s, exit 0**; `examples/` không đổi, không Excel mồ côi, không thư mục
  tạm mới.
- `results/test-results.json`: 105/105 integration.
- `results/plan-accuracy-q4.json`: 53/55; 49 đạt, 1 từ chối đúng, 0 lỗi.
- `results/secret-scan.json`: 0 finding ngoài allowlist (lượt quét trước Q7).
- `test_real_like_plans.py` mất 703,6s, có khoảng trễ ~4 phút giữa hai workbook
  chưa rõ nguyên nhân; không có lỗi, đã ghi.
- Chi tiết: `cycles/CYCLE-07.md`, `CYCLE-08.md`, `CYCLE-09.md`, `UPGRADE_LOG.md`.

## Chạy lại trên Windows
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_all_tests.ps1
```
Cần Excel Desktop cho nhóm 3; GUI cần màn hình. Không cần mạng; không gọi cloud.

## Còn lại (Q7)
- Tạo `.github/workflows/tests.yml` (17 test không cần Excel trên Windows +
  Ubuntu; Windows thêm 2 bộ bảo mật). **Không push.**
- Cập nhật README/RELEASES, tăng phiên bản lên 0.4.0, tag cục bộ `v0.4.0`.
- Cập nhật `WINDOWS_ACCEPTANCE.md` và xuất handoff cuối.

## Cần chủ sản phẩm quyết
- **LICENSE:** có thêm không, chọn MIT/Apache-2.0/giữ riêng tư? Hiện **chưa có**,
  không tự thêm.
- Có muốn push nhánh `muse-upgrade` / tạo tag phát hành trên remote không? Hiện
  mọi thứ chỉ ở máy.
- Xác nhận bằng mắt giao diện, hộp xem trước cloud, bố cục 100%/150% DPI.
