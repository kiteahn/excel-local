# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 11 — Tầng kiểm tra 2: danh sách ≥30 use case, tự chạy phần chạy được

Theo tầng 2 trong PROMPT_MUSE.md. Môi trường này không có Excel nên chỉ chạy
được use case chạm đường thuần Python: preflight (từ chối macro/liên kết
ngoài/công thức ngoài allowlist), inspect CSV (UTF-8/UTF-16, phân cách
`,`/`;`/tab, ragged), inspect lenient qua fast reader, detect_header_row/
select_header (banner, ô gộp mô phỏng bằng snapshot dict), preview_plan các
action, validate_plan, key_value/normalized_keys (mã 00123, khóa trùng/trống,
NFC/NFD), giới hạn 20.000 dòng/60 cột/400.000 ô.

Làm:
1. Lập bảng ≥30 use case trong file mới `USECASES.md` (hoặc mục trong
   UPGRADE_LOG): mỗi cái ghi đường chạy (hàm nào), kỳ vọng, trạng thái.
2. Viết `test_usecases.py` chạy các use case chạy được; cái nào cần Excel
   COM thì ghi rõ "CHƯA CHẠY ĐƯỢC — cần Windows" trong bảng, không bỏ qua
   lặng lẽ.
3. Lỗi thật phát hiện được → sửa tối thiểu + regression, như các vòng trước.
4. Chạy toàn bộ suite cũ để bảo đảm không hồi quy. Commit, cập nhật log.

## Hàng đợi sau đó
- Vòng 12+: tách Lookup/Filter trong infer_smart_plan; fuzz CSV/header;
  mutation test bộ test hiện có.
- Hàng đợi Windows (chỉ làm được trên máy Windows có Excel): đo lại
  benchmark.py cho code hiện tại; chạy test_gui.py thật + kiểm tra trực quan
  tooltip/phím tắt (Vòng 10) và messagebox DPAPI (Vòng 3); nghiệm thu DPAPI
  thật (security mục 16–17, router mục 1/7).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
- Nhánh `muse-upgrade` hiện có 11 commit (Bước 0 + Vòng 1–10) CHƯA được push
  lên GitHub. Việc push cần người dùng xác nhận (đây là repo công khai).
