# Gói nghiệm thu Windows — Excel Local Expert, nhánh `muse-upgrade`

Ngày xuất: 2026-10-30 (bản cập nhật sau Vòng 31). Commit đầu nhánh:
`a9a1c5a` (Bước 0 + Vòng 1–31 trên nền bản gốc `28f5679` — v0.3.1).
Đã gồm 2 quyết định chính sách của bạn (2026-10-30): mặc định chỉ Ollama
local (cloud tự bật) và tùy chọn `vn_dot_thousands` cho số '500.000'.
Toàn bộ phần việc trên Linux đã xong và có bằng chứng; tài liệu này là
phần CHỈ làm được trên máy Windows có Excel. Danh sách nghiệm thu đầy đủ
từng bước cũng nằm trong `NEXT.md` của repo.

## Gói này gồm
- `excel-local-muse-upgrade.bundle` — toàn bộ repo + lịch sử nhánh
  `muse-upgrade` (đã kiểm chứng: clone từ bundle chạy được test_app đạt).
- `excel-local-muse-upgrade.zip` — cây thư mục tại đầu nhánh (giải nén là
  dùng ngay, không cần git).
- `muse-upgrade.patch` — diff đầy đủ so với bản gốc (để xem nhanh thay đổi).
- File này.

## Cách lấy code (chọn 1)
```powershell
# Cách A — từ bundle (giữ lịch sử git):
git clone excel-local-muse-upgrade.bundle excel-local-check
cd excel-local-check
git checkout muse-upgrade

# Cách B — giải nén file zip vào một thư mục mới, mở PowerShell tại đó.
```

## Danh sách file đã đổi so với bản gốc (46 mục)
- **Code chính:** `app.py`, `ai_client.py` (sửa), `excel_bridge.ps1` (KHÔNG đổi).
- **Tài liệu mới/sửa:** `README.md` (viết lại theo hành vi thật), `UPGRADE_LOG.md`
  (hồ sơ 22 vòng), `NEXT.md`, `USECASES.md` (37 use case), `PROMPT_MUSE.md` +
  `run-muse-loop.ps1` (bản bạn gửi, lưu vào repo).
- **Test mới:** `test_usecases.py`, `test_fuzz_light.py`, `test_fuzz_headers.py`,
  `test_infer_golden.py`, `test_sanitize_plan.py`, `test_privacy_guard_cases.py`,
  `benchmark_python_stages.py`.
- **Test sẵn có đã sửa:** `test_router.py` (bỏ đường dẫn cứng, hermetic, thêm
  regression DPAPI), và cơ chế SKIP theo nền tảng + dọn import trong
  `test_app.py`, `test_all_edge_cases.py`, `test_batch_processing.py`,
  `test_cli.py`, `test_diverse_scenarios_e2e.py`, `test_gui.py`,
  `test_headers.py`, `test_read_boundary.py`, `test_security_privacy.py`,
  `test_single_file_filter.py`, `test_upgrades.py`, `test_cleanup.ps1`.
- **Dọn repo:** 14 file JSON kết quả/benchmark chuyển vào `results/`;
  `benchmark.py`, `benchmark_ai.py` cập nhật đường dẫn; `create_diverse_examples.py`,
  `generate_project_pdf.py` bỏ đường dẫn cứng/import thừa.

## Lệnh chạy test trên Windows

Cách nhanh nhất — chạy đủ 3 nhóm, in bảng tổng hợp, kiểm tra `examples/` và Excel mồ côi:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run_all_tests.ps1
```

Script trả exit code 0 chỉ khi mọi suite đạt, `examples/` không bị đổi, và không còn
tiến trình Excel do bridge tạo. Log từng suite ở `.work/qa-logs/`. Có thể bỏ nhóm 3
bằng `-SkipExcel`, bỏ riêng GUI bằng `-SkipGui`, bỏ benchmark bằng `-SkipBenchmark`;
`-Only ten_file.py` để chạy đúng một suite.

Nếu muốn chạy tay từng nhóm (nội dung y hệt script):

### Nhóm 1 — Không cần Excel (chạy ở đâu cũng được, kỳ vọng: tất cả exit 0)
```powershell
python -X utf8 test_app.py
python -X utf8 test_all_edge_cases.py
python -X utf8 test_upgrades.py
python -X utf8 test_cli.py
python -X utf8 test_headers.py
python -X utf8 test_usecases.py
python -X utf8 test_usecases2.py
python -X utf8 test_fuzz_light.py
python -X utf8 test_fuzz_headers.py
python -X utf8 test_fuzz_xlsx.py
python -X utf8 test_infer_golden.py
python -X utf8 test_sanitize_plan.py
python -X utf8 test_policy_defaults.py
python -X utf8 test_safety_limits.py
python -X utf8 test_batch_summary.py
python -X utf8 test_privacy_guard_cases.py
python -X utf8 test_single_file_filter.py
python -X utf8 test_diverse_scenarios_e2e.py
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
python -X utf8 test_app.py --integration  # 105 bài: inspect/preview/execute/đối soát/PDF qua Excel COM
python -X utf8 test_read_boundary.py      # có lượt đọc Excel thật (chống tráo file)
python -X utf8 test_gui.py                # 14 bài GUI (cần màn hình; đừng khóa máy khi chạy)
python -X utf8 benchmark.py               # đo lại thời gian Excel COM cho code hiện tại
```

Lưu ý mới từ Bước 0:
- `test_app.py` ghi kết quả riêng theo chế độ: `results/test-results-unit.json` (chỉ
  logic), `results/test-results.json` (đủ integration), `results/test-results-ai.json`.
- Bridge Excel có retry **có giới hạn**: chỉ thử lại một lần khi timeout mà tiến trình
  Excel của tác vụ đã dọn được, hoặc khi thiếu `result.json`; lỗi Excel thật
  (`ok=false`) không bao giờ được thử lại. Đo 20 lần thật: 19/20 đạt, 1 lần hỏng cả hai
  attempt; một lần timeout 60s đã được retry thành công (xem
  `results/bridge-retry-20.json`).

## Kiểm tra thủ công (5–10 phút, ghi kết quả đạt/không)
1. **UC-32:** mở app, chạy trọn luồng với 1 file thật: Đọc → Đề xuất → Xem
   trước → Chấp nhận → kiểm tra sheet kết quả + sheet Kiem_tra + biên bản.
   Nhìn thêm danh sách sheet: nếu có một sheet trống mặc định (Sheet1)
   không thuộc dữ liệu của bạn thì báo lại (điểm W-1, Vòng 31).
2. **UC-33:** mở sẵn file nguồn trong Excel rồi bấm Đọc — app phải báo lỗi
   dễ hiểu (file đang bị khóa), không treo/sập.
3. **UC-35:** file có ô lỗi #DIV/0! — phải bị từ chối rõ ràng.
4. **UC-36:** bật xuất PDF, mở PDF kiểm tra bố cục.
5. **Tooltip + phím tắt (Vòng 10):** rê chuột lên các nút chọn động cơ AI /
   nút Xem trước — phải thấy tooltip; trong ô yêu cầu, nhấn **Ctrl+Enter**
   — phải chạy Đề xuất.
6. **Hộp lỗi Batch (VĐ-17):** chạy Xử lý hàng loạt với 1 thư mục chứa file
   hỏng — phải hiện hộp thoại báo lỗi tử tế (bản gốc sẽ sập âm thầm bằng
   NameError ở đường này).
7. **DPAPI (VĐ-11):** mở Cài đặt AI, nhập 1 API key, Lưu — mở `config.json`
   kiểm tra key phải có dạng `dpapi:...`, KHÔNG được là khóa thô.
8. **Hai quyết định mới:** hàng "Động cơ AI" phải có nút Ollama và mặc
   định chọn Ollama (trừ khi config cũ của bạn đã lưu lựa chọn khác);
   trong Cài đặt AI có checkbox "Số dạng 500.000 hiểu là 500 nghìn" —
   thử tổng hợp lenient với ô '500.000': bật → tính 500000, tắt → tính 500.

## Cần báo lại cho Muse
- Kết quả từng nhóm lệnh (đạt / lỗi — nếu lỗi, gửi nguyên văn output).
- 7 mục kiểm tra thủ công ở trên: đạt hay không, kèm ảnh chụp nếu có lỗi.
- Số liệu `benchmark.py` in ra (để đối chiếu với bản cũ).
Sau khi có kết quả này mới xét tiếp việc tách `gui()` (đang đóng băng theo
chỉ đạo của bạn) và xử lý các lỗi phát hiện trên Windows nếu có.
