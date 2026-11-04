# Bàn giao Chu kỳ 10 — Excel Local Expert v0.4.0 (`muse-upgrade`)

Ngày: 2026-11-04. Nhánh `muse-upgrade`, **không push**. Không có `STOP.txt`
tại thời điểm xuất. Tag cục bộ: `v0.4.0`.

## Gói trong thư mục này
| File | Nội dung |
|---|---|
| `excel-local-muse-upgrade-cycle10.bundle` | Toàn bộ lịch sử nhánh `muse-upgrade` + tag `v0.4.0`; đã `git bundle verify` ("complete history") |
| `excel-local-muse-upgrade-cycle10.zip` | Cây file tại đúng commit đầu nhánh, chỉ gồm file git theo dõi |
| `muse-upgrade-cycle10.patch` | Diff `28f5679..HEAD` so với bản gốc |
| `HANDOFF.md` | File này |

Lấy code:
```powershell
git clone handoff/excel-local-muse-upgrade-cycle10.bundle excel-local-check
cd excel-local-check
git checkout muse-upgrade
git tag --list v0.4.0
```
Hoặc giải nén `excel-local-muse-upgrade-cycle10.zip` vào thư mục mới.

## Đã xong toàn bộ Bước 0 + Q1–Q7
- **Bước 0:** fixture tạm + fingerprint `examples/`; kết quả test tách theo chế
  độ; retry bridge có giới hạn; một lệnh `run_all_tests.ps1`; quét bí mật.
- **Q1 (Chu kỳ 4):** tách `_prepare_lookup`; `prepare_sheets` 262 → 169 dòng
  (nay 177 sau Q4); 27 kịch bản vàng; mutation 6/6.
- **Q2 (Chu kỳ 5):** PrivacyGuard 212 câu (125 PII + 87 non-PII), 0 rò lọt ngoài
  2 giới hạn đã ghi, 0% che nhầm; `cloud_preview` + hộp xem trước/hủy.
- **Q3 (Chu kỳ 6):** 5 lần đo/loại; trung vị 20k: đọc nguồn 5,07s, danh mục
  4,81s, đối chiếu 19,89s, tổng hợp 16,80s, làm sạch 11,99s; không giữ tối ưu
  dưới 10%.
- **Q4 (Chu kỳ 7):** 9 workbook tổng hợp, 6 loại bảng, 55 câu; **53/55 = 96,4%**
  kế hoạch đúng; 49 lượt Excel đạt, 1 từ chối đúng `#N/A`, 0 lỗi.
- **Q5 (Chu kỳ 8):** 3 ảnh Tk thật trong `docs/screens/`; đã chụp lại cho v0.4.0.
- **Q6 (Chu kỳ 9):** `_filter_plan` 226 → 121 dòng, `gui()` 713 → 23,
  `App.__init__` 213 → 16; 34 ca vàng; mutation 6/6.
- **Q7 (Chu kỳ 10):** `.github/workflows/tests.yml` (17 test không cần Excel trên
  Windows + Ubuntu; Windows thêm 2 bộ DPAPI/bảo mật); đồng bộ tiêu đề v0.4.0;
  README/RELEASES/release notes; tag cục bộ `v0.4.0`; **không push, không LICENSE**.

## Bằng chứng mới nhất
- Full suite: `.work/qa-logs/run-20261104-105727.log` — **31 lệnh, 0 lỗi,
  1018,7s, exit 0**; `examples/` không đổi, không Excel mồ côi, không thư mục
  tạm mới.
- `results/test-results.json`: 105/105 integration.
- `results/plan-accuracy-q4.json`: 53/55; 49 đạt, 1 từ chối đúng, 0 lỗi.
- `results/secret-scan.json`: 0 finding ngoài allowlist sau khi commit Q7.
- `test_real_like_plans.py` mất 703,6s, có khoảng trễ ~4 phút giữa hai workbook
  chưa rõ nguyên nhân; không có lỗi, đã ghi.
- Chi tiết: `cycles/CYCLE-04.md` … `CYCLE-10.md`, `UPGRADE_LOG.md`,
  `WINDOWS_ACCEPTANCE.md`.

## Chạy lại trên Windows
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_all_tests.ps1
```
Cần Excel Desktop cho nhóm 3; GUI cần màn hình. Không cần mạng; không gọi cloud.

## Còn lại / chưa kiểm chứng
- Giao diện, hộp xem trước cloud, bố cục 100%/150% DPI: **CẦN NGƯỜI XÁC NHẬN
  BẰNG MẮT**; ảnh hiện có ở 125% DPI.
- Một lượt GitHub Actions thật trên Ubuntu: **CHƯA KIỂM CHỨNG** (mới kiểm cú
  pháp YAML và chạy bộ chọn trên Windows).
- AVG/trung bình chưa hỗ trợ; 2 ca Q4 vẫn là giới hạn đã ghi.

## Cần chủ sản phẩm quyết
- **LICENSE:** có thêm không, chọn MIT/Apache-2.0/giữ riêng tư? Hiện **chưa có**,
  không tự thêm.
- Có muốn push nhánh `muse-upgrade` / tag `v0.4.0` lên remote không? Hiện mọi
  thứ chỉ ở máy.
