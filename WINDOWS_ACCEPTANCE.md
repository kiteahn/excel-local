# Gói nghiệm thu Windows — Excel Local Expert, nhánh `muse-upgrade`

Ngày xuất: **2026-11-04** (sau Chu kỳ 10 — Q4, Q5, Q6, Q7).
Nhánh `muse-upgrade`, **không push**. Bản phát hành mới là **v0.4.0** (tag cục
bộ `v0.4.0`); hành vi trước đợt này là v0.3.2.
Toàn bộ phần chạy được trên máy Windows có Excel đã có bằng chứng; các mục cần
mắt người được đánh dấu **CẦN NGƯỜI XÁC NHẬN** trong `NEXT.md`.

## Gói này gồm
- `handoff/excel-local-muse-upgrade-cycle10.bundle` — toàn bộ repo + lịch sử
  nhánh `muse-upgrade` và tag `v0.4.0`.
- `handoff/excel-local-muse-upgrade-cycle10.zip` — cây thư mục tại đúng commit
  đầu nhánh (chỉ file git theo dõi; không chứa `config.json` hay `.work/`).
- `handoff/muse-upgrade-cycle10.patch` — diff đầy đủ so với bản gốc `28f5679`.
- `handoff/HANDOFF.md` — tóm tắt commit, cách chạy test, việc còn lại.
- File này.

## Cập nhật sau Chu kỳ 10 (Q4–Q7)
- **Q4:** 9 workbook tổng hợp trong `examples/real_like/`, 6 loại bảng nguồn;
  55 câu tiếng Việt, **53/55 = 96,4%** kế hoạch đúng ngay lần đầu. Chạy đầy đủ
  preview → Excel → verify → preserve: **49 thành công, 1 từ chối đúng vì
  `#N/A`, 0 lỗi**. Hai ca còn sai là AVG/trung bình chưa hỗ trợ.
- **Q5:** 3 ảnh giao diện trong `docs/screens/` (màn hình chính, hộp kiểm tra,
  Cài đặt AI), chụp từ Tk thật ở DPI 125%, cấu hình rỗng. Đã sửa lỗi callback
  chụp ảnh; ghi nhận lệch tiêu đề `v0.3`/`v0.3.2` (đồng bộ ở bản 0.4.0).
- **Q6:** `_filter_plan` 226 → **121 dòng** (3 helper), `gui()` 713 → **23
  dòng**, `App.__init__` 213 → **16 dòng**, `_build_ui` 199 dòng. Vàng tăng lên
  **34 ca**; mutation **6/6** bị bắt; không đổi hành vi. Dừng tách lớp `App`
  còn 699 dòng vì rủi ro giao diện lớn hơn lợi ích.
- **Q7:** phiên bản **0.4.0**, đồng bộ tiêu đề app/cửa sổ, chụp lại 3 ảnh GUI;
  `.github/workflows/tests.yml` chạy 17 test không cần Excel trên Windows +
  Ubuntu, Windows thêm 2 bộ bảo mật/DPAPI. **Không push, không thêm LICENSE.**
  YAML đã kiểm bằng PyYAML; **CHƯA KIỂM CHỨNG** runner Ubuntu thật.
- **Test đầy đủ cuối:** `run_all_tests.ps1` **31 lệnh, 0 lỗi, 1018,7s, exit 0**
  (`run-20261104-105727.log`); `examples/` không đổi, không còn Excel mồ côi,
  không có thư mục tạm mới. `test_real_like_plans.py` mất **703,6s** thay vì
  khoảng 327s như Q4; có khoảng trễ ~4 phút giữa hai workbook, chưa xác định
  nguyên nhân, nhưng không có lỗi. Không giấu ngoại lệ này.
- **CẦN NGƯỜI XÁC NHẬN BẰNG MẮT:** giao diện tiếng Việt, hộp xem trước nội
  dung cloud, màu/trạng thái nút, bố cục ở 100% và 150% DPI (ảnh hiện có chỉ ở
  125%).
- **Câu hỏi chờ chủ sản phẩm:** có đồng ý thêm `LICENSE` (giấy phép nào) và có
  muốn push nhánh/tạo tag phát hành trên remote không. Không tự thêm, không push.

## Cách lấy code (chọn 1)
```powershell
# Cách A — từ bundle (giữ lịch sử git + tag v0.4.0):
git clone handoff/excel-local-muse-upgrade-cycle10.bundle excel-local-check
cd excel-local-check
git checkout muse-upgrade
git tag --list v0.4.0

# Cách B — giải nén file handoff/excel-local-muse-upgrade-cycle10.zip vào một thư mục mới, mở PowerShell tại đó.
```

## Lệnh chạy test trên Windows

Cách nhanh nhất — chạy đủ 3 nhóm, in bảng tổng hợp, kiểm tra `examples/` và Excel mồ côi:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_all_tests.ps1
```

Script trả exit code 0 chỉ khi mọi suite đạt, `examples/` không bị đổi, và không còn
tiến trình Excel do bridge tạo. Log từng suite ở `.work/qa-logs/`. Có thể bỏ nhóm 3
bằng `-SkipExcel`, bỏ riêng GUI bằng `-SkipGui`, bỏ benchmark bằng `-SkipBenchmark`;
`-Only ten_file.py` để chạy đúng một suite.

### Nhóm 1 — Không cần Excel (chạy ở đâu cũng được, kỳ vọng: tất cả exit 0)
```powershell
python -X utf8 test_app.py
python -X utf8 test_all_edge_cases.py
python -X utf8 test_upgrades.py
python -X utf8 test_cli.py
python -X utf8 test_usecases.py
python -X utf8 test_usecases2.py
python -X utf8 test_fuzz_light.py
python -X utf8 test_fuzz_headers.py
python -X utf8 test_fuzz_xlsx.py
python -X utf8 test_infer_golden.py
python -X utf8 test_sanitize_plan.py
python -X utf8 test_policy_defaults.py
python -X utf8 test_safety_limits.py
python -X utf8 test_prepare_sheets_golden.py
python -X utf8 test_batch_summary.py
python -X utf8 test_privacy_guard_cases.py
python -X utf8 test_cloud_preview.py
python -X utf8 test_batch_processing.py
python -X utf8 benchmark_python_stages.py
python -X utf8 secret_scan.py
```

### Nhóm 2 — Cần Windows (DPAPI thật), không cần Excel mở sẵn
```powershell
python -X utf8 test_router.py            # 7 mục; mục 1 và 7 kiểm tra DPAPI thật
python -X utf8 test_security_privacy.py  # 43 mục trên Windows (38 logic + 5 DPAPI ở mục 16-17)
```

### Nhóm 3 — CẦN EXCEL DESKTOP THẬT (quan trọng nhất của đợt nghiệm thu này)
```powershell
python -X utf8 test_app.py --integration   # 105 bài: inspect/preview/execute/đối soát/PDF qua Excel COM
python -X utf8 test_read_boundary.py       # 4 bài; có lượt đọc Excel thật (chống tráo file)
python -X utf8 test_real_like_plans.py     # 55 câu trên 9 workbook; 50 kế hoạch chạy Excel đầy đủ
python -X utf8 test_headers.py             # 11 bài; phần native Excel chỉ chạy trên Windows
python -X utf8 test_single_file_filter.py  # 30 bài; trên Windows dùng Excel Desktop
python -X utf8 test_diverse_scenarios_e2e.py # 6 kịch bản; trên Windows dùng Excel Desktop
python -X utf8 test_gui.py                 # 14 bài GUI (cần màn hình; đừng khóa máy khi chạy)
python -X utf8 test_ui_acceptance.py       # 48 bài giao diện
python -X utf8 benchmark.py                # đo lại thời gian Excel COM cho code hiện tại
```

Lưu ý:
- `test_app.py` ghi kết quả riêng theo chế độ: `results/test-results-unit.json`
  (chỉ logic), `results/test-results.json` (đủ integration),
  `results/test-results-ai.json`.
- Bridge Excel có retry **có giới hạn**: chỉ thử lại một lần khi timeout mà tiến
  trình Excel của tác vụ đã dọn được, hoặc khi thiếu `result.json`; lỗi Excel
  thật (`ok=false`) không bao giờ được thử lại.
- Hai lần Excel treo tạm thời 253–262s đã ghi trong Q3; báo cáo dùng trung vị,
  không giấu ngoại lệ.

## Kiểm tra thủ công (5–10 phút, ghi kết quả đạt/không)
1. **UC-32:** mở app, chạy trọn luồng với 1 file thật: Đọc → Đề xuất → Xem
   trước → Chấp nhận → kiểm tra sheet kết quả + sheet Kiem_tra + biên bản.
2. **UC-33:** mở sẵn file nguồn trong Excel rồi bấm Đọc — app phải báo lỗi
   dễ hiểu (file đang bị khóa), không treo/sập.
3. **UC-35:** file có ô lỗi #DIV/0! — phải bị từ chối rõ ràng.
4. **UC-36:** bật xuất PDF, mở PDF kiểm tra bố cục.
5. **Tooltip + phím tắt:** rê chuột lên nút chọn động cơ AI / nút Xem trước —
   phải thấy tooltip; trong ô yêu cầu, nhấn **Ctrl+Enter** — phải chạy Đề xuất.
6. **Hộp lỗi Batch:** chạy Xử lý hàng loạt với 1 thư mục chứa file hỏng — phải
   hiện hộp thoại báo lỗi tử tế.
7. **DPAPI:** mở Cài đặt AI, nhập 1 API key, Lưu — mở `config.json` kiểm tra key
   phải có dạng `dpapi:...`, KHÔNG được là khóa thô.
8. **Hai quyết định chính sách:** hàng "Động cơ AI" mặc định chọn Ollama (trừ
   khi config cũ đã lưu lựa chọn khác); checkbox "Số dạng 500.000 hiểu là 500
   nghìn" bật → tổng hợp 500000, tắt → 500.
9. **Q4 — dữ liệu giống thật:** mở một workbook trong `examples/real_like/`,
   chạy một yêu cầu lọc/tổng hợp/đối chiếu và kiểm tra sheet kết quả; thử một
   bảng có banner/dòng tổng.
10. **Q5 — giao diện:** mở `docs/screens/` xem 3 ảnh; kiểm tra bố cục ở 100%
    và 150% DPI trên máy bạn. Ảnh chỉ là bằng chứng thị giác, không thay test.

## Cần báo lại cho Muse
- Kết quả từng nhóm lệnh (đạt / lỗi — nếu lỗi, gửi nguyên văn output).
- 10 mục kiểm tra thủ công ở trên: đạt hay không, kèm ảnh chụp nếu có lỗi.
- Số liệu `benchmark.py` in ra (để đối chiếu với bản cũ).
