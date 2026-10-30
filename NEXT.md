# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Trạng thái tổng (2026-10-30)
- Nhánh `muse-upgrade`: 25 commit (Bước 0 + Vòng 1–24), CHƯA push —
  người dùng quyết định TỰ push (2026-10-30), Muse không push.
- Hai quyết định chính sách đã áp dụng ở Vòng 23: (1) mặc định Ollama
  local, cloud tự bật; (2) tùy chọn `vn_dot_thousands` cho '500.000'.
- **ĐÓNG BĂNG tách gui()** cho tới khi người dùng báo kết quả Windows.
- Việc làm được trên Linux đã cạn (fuzz 2 đợt, mutation 4 đợt/20 đột
  biến, rà soát tĩnh pyflakes — tất cả xong). Còn lại: nghiệm thu Windows.

## DANH SÁCH NGHIỆM THU WINDOWS ĐẦY ĐỦ (làm trên máy Windows có Excel)

Chuẩn bị: lấy code bằng bundle/zip trong gói bàn giao
(`excel-local-muse-upgrade.bundle` hoặc `.zip`), mở PowerShell tại thư
mục code, xác nhận đang ở đúng commit:

```
git checkout muse-upgrade   (nếu dùng bundle)
git log --oneline -1        → phải là commit mới nhất của nhánh
```

### Bước 1 — Nhóm test không cần Excel (kỳ vọng: mọi lệnh exit 0 / báo đạt)
```
python -X utf8 test_app.py                  → 79 đạt
python -X utf8 test_all_edge_cases.py       → 61 đạt
python -X utf8 test_upgrades.py             → 28 đạt
python -X utf8 test_cli.py                  → 6 đạt
python -X utf8 test_headers.py              → 11 đạt
python -X utf8 test_usecases.py             → 37 đạt (UC + SEC)
python -X utf8 test_fuzz_light.py           → 20 đạt
python -X utf8 test_fuzz_headers.py         → 15 đạt
python -X utf8 test_infer_golden.py         → 28 đạt
python -X utf8 test_sanitize_plan.py        → 8 đạt
python -X utf8 test_policy_defaults.py      → 12 đạt
python -X utf8 test_safety_limits.py        → 6 đạt
python -X utf8 test_privacy_guard_cases.py  → 53 mẫu: 0 lọt bất ngờ, 0 che nhầm bất ngờ
python -X utf8 test_single_file_filter.py   → 24 đạt
python -X utf8 test_diverse_scenarios_e2e.py → 6 kịch bản đạt
python -X utf8 test_batch_processing.py     → 10 đạt (phần logic)
python -X utf8 benchmark_python_stages.py   → in bảng thời gian, preview < ~0,3s ở 20.000 dòng
```

### Bước 2 — Cần Windows/DPAPI thật (không cần mở Excel)
```
python -X utf8 test_router.py            → 7/7 mục, gồm mục 1 và 7 chạy DPAPI THẬT (không còn SKIP)
python -X utf8 test_security_privacy.py  → 43/43 (trên Linux chỉ đạt 38 + 5 SKIP DPAPI ở mục 16–17)
```
Kỳ vọng riêng cho VĐ-11: cố tình làm DPAPI lỗi (nếu làm được) → save_config
phải TỪ CHỐI lưu và báo lỗi rõ, config.json không chứa khóa thô.

### Bước 3 — Cần Excel Desktop thật (quan trọng nhất)
```
python -X utf8 test_app.py --integration  → 97 bài qua Excel COM: inspect/preview/execute/đối soát/PDF
python -X utf8 test_gui.py                → 14 bài GUI (cần màn hình đang mở, đừng khóa máy)
python -X utf8 test_read_boundary.py      → đạt cả lượt đọc Excel thật (chống tráo file giữa preview và execute)
python -X utf8 benchmark.py               → ghi lại số giây in ra (đối chiếu bản cũ trong results/)
```

### Bước 4 — Kiểm tra thủ công trong app (mỗi mục ghi đạt/không)
1. UC-32: trọn luồng 1 file thật: Đọc → Đề xuất (mặc định phải là Ollama)
   → Xem trước → Chấp nhận → mở kết quả: đúng sheet kết quả + sheet
   Kiem_tra + biên bản JSON.
2. UC-33: mở sẵn file nguồn trong Excel rồi bấm Đọc → báo lỗi dễ hiểu, không treo.
3. UC-35: file có ô #DIV/0! → bị từ chối rõ ràng.
4. UC-36: bật Xuất PDF → mở PDF, bố cục đọc được.
5. Tooltip (Vòng 10): rê chuột lên 4 nút động cơ (Ollama/Siêu tốc/Smart
   Router/Offline) và nút Xem trước → thấy tooltip; ô yêu cầu: Ctrl+Enter
   → chạy Đề xuất. Hàng động cơ phải có nút Ollama và được chọn sẵn theo
   mặc định mới (nếu config cũ của bạn chưa lưu lựa chọn khác).
6. Cài đặt AI: thấy checkbox 'vn_dot_thousands' (500.000) và combobox động
   cơ mặc định 'ollama'; nhập thử API key → Lưu → mở config.json: key dạng
   `dpapi:...`, không có khóa thô.
7. Batch (VĐ-17): Xử lý hàng loạt với thư mục có 1 file hỏng → hiện hộp
   thoại báo lỗi tử tế (không im lặng/sập).
8. Số học quyết định 2: tổng hợp lenient với cột tiền có '500.000' → tổng
   tính 500000 khi checkbox bật, 500 khi tắt.

### Bước 5 — Báo lại cho Muse
- Output nguyên văn của mọi lệnh ở Bước 1–3 (hoặc ít nhất các lệnh lỗi).
- Kết quả 8 mục ở Bước 4 (đạt/không + ảnh chụp nếu lỗi).
- Số liệu benchmark.py.
Sau đó: Muse ghi kết quả vào UPGRADE_LOG (mục "Nghiệm thu Windows"),
sửa lỗi phát hiện (nếu có) TRƯỚC, rồi mới xét tách gui().

## Điều kiện dừng vòng lặp
- File `STOP.txt` ở thư mục gốc repo (kiểm tra sau mỗi vòng — tới nay
  chưa từng xuất hiện).
