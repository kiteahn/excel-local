# Bàn giao Chu kỳ 6 — Excel Local Expert (`muse-upgrade`)

Ngày: 2026-11-02. Nhánh `muse-upgrade`, **không push**. Không có `STOP.txt`
tại thời điểm xuất. Commit đầu nhánh: `0160c6a`.

## Gói trong thư mục này
| File | Nội dung |
|---|---|
| `excel-local-muse-upgrade-cycle6.bundle` | Toàn bộ lịch sử nhánh `muse-upgrade`; đã `git bundle verify` ("complete history") |
| `excel-local-muse-upgrade-cycle6.zip` | Cây file tại đúng commit `0160c6a`, chỉ gồm file git theo dõi |
| `muse-upgrade-cycle6.patch` | Diff `28f5679..0160c6a` so với bản gốc |
| `HANDOFF.md` | File này |

Lấy code:
```powershell
git clone handoff/excel-local-muse-upgrade-cycle6.bundle excel-local-check
cd excel-local-check
git checkout muse-upgrade
```
Hoặc giải nén `excel-local-muse-upgrade-cycle6.zip` vào thư mục mới.

## Đã xong
- **Bước 0:** B0.1 fixture tạm + fingerprint `examples/`; B0.2 kết quả test theo chế độ; B0.3 retry bridge có kiểm soát; B0.4 một lệnh 3 nhóm + kiểm tra rác; B0.5 quét bí mật/đường dẫn cá nhân, 0 finding ngoài allowlist.
- **Q1 (Chu kỳ 4):** tách `_prepare_lookup`; `prepare_sheets` 262 → 169 dòng; 27 kịch bản vàng; mutation 6/6.
- **Q2 (Chu kỳ 5):** corpus 212 câu (125 PII + 87 non-PII), 0 rò lọt ngoài 2 giới hạn đã ghi, 0% che nhầm; `cloud_preview` + hộp xem trước/cancel; README ghi giới hạn.
- **Q3 (Chu kỳ 6):** 5 lần đo Python stages, 5 lần 5k, 5 lần 20k; trung vị ở `results/benchmark-q3-summary.json`; không giữ tối ưu <10%; test đầy đủ 30 lệnh/0 lỗi/292,0s.

## Bằng chứng mới nhất
- `results/test-results.json`: 105/105 integration.
- `results/benchmark-q3-summary.json`: trung vị và ngoại lệ.
- `.work/qa-logs/run-20261102-104045.log`: bảng 30 lệnh, 0 lỗi.
- `results/secret-scan.json`: 0 finding ngoài allowlist.
- Chi tiết: `cycles/CYCLE-04.md`, `CYCLE-05.md`, `CYCLE-06.md`, `UPGRADE_LOG.md`.

## Chạy lại trên Windows
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_all_tests.ps1
```
Cần Excel Desktop cho nhóm 3; GUI cần màn hình. Không cần mạng; không gọi cloud.

## Còn lại (Q4–Q7)
- **Q4:** bộ dữ liệu giống thật ≥6 loại + ≥50 câu, đo tỷ lệ kế hoạch đúng, thêm ca sai vào golden.
- **Q5:** ảnh GUI vào `docs/screens/`, chỉ sửa lỗi đo được, ghi **CẦN NGƯỜI XÁC NHẬN BẰNG MẮT**.
- **Q6:** tách tiếp `App.__init__`, `_filter_plan`, `gui()` theo quy trình vàng.
- **Q7:** chỉ tạo `.github/workflows/tests.yml` (không push), cập nhật README/RELEASES, tăng phiên bản, tag cục bộ.

## Cần chủ sản phẩm quyết
- **LICENSE:** có thêm không, và chọn giấy phép nào? Hiện **chưa có**, không tự thêm.
- Có muốn push nhánh `muse-upgrade` / tạo tag phát hành không? Hiện mọi thứ chỉ ở máy.
- Xác nhận bằng mắt giao diện tiếng Việt, hộp xem trước cloud, bố cục 100%/150% DPI.
