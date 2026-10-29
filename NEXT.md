# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 4 — Skip rõ ràng theo nền tảng cho test (VĐ-04, mục G phần đầu)

Mục tiêu: trên máy không có Excel/PowerShell/DPAPI/display, các test phải in
"SKIP: <lý do>" cho đúng phần phụ thuộc thay vì sập traceback, để tín hiệu hồi
quy thuần logic không bị che. Trên Windows có đủ thành phần, mọi check vẫn
chạy như cũ — không được làm yếu test.

Làm:
1. `test_security_privacy.py`: mục 16–17 (DPAPI + lưu config) chỉ chạy khi
   `dpapi_protect('probe')` cho kết quả `dpapi:`; nếu không, in SKIP và không
   tính là fail. Các mục còn lại (PII, pseudonym, sanitize...) vẫn chạy và
   phải đạt trên Linux.
2. `test_headers.py`, `test_single_file_filter.py`,
   `test_diverse_scenarios_e2e.py`, `test_read_boundary.py`,
   `test_batch_processing.py`: kiểm tra `shutil.which('powershell.exe')` (và
   Excel bridge) ở đầu; thiếu thì in SKIP rõ lý do và exit 0, thay vì
   FileNotFoundError.
3. `test_gui.py`: bọc phần tạo cửa sổ — nếu TclError/no display thì in SKIP.
   Nếu môi trường có `xvfb` thì ưu tiên chạy thật; ghi rõ đã dùng cách nào.
4. Kiểm tra: chạy toàn bộ test_*.py trên Linux — kỳ vọng: mọi file exit 0,
   các mục SKIP được in rõ; tổng check thuần logic đạt không giảm
   (79 + 61 + 28 + 6 + router 7 + security phần logic).
5. Commit, cập nhật UPGRADE_LOG.md, chuyển NEXT sang Vòng 5 (PrivacyGuard).

## Hàng đợi sau đó
- Vòng 5: Mục E — bộ test PrivacyGuard ≥50 mẫu, đo lọt/che nhầm, sửa theo số liệu.
- Vòng 6+: VĐ-05 (except hẹp), VĐ-06 (tách hàm), VĐ-07 (đo hiệu năng),
  VĐ-09 (dọn repo), VĐ-10 (giao diện nhẹ).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
