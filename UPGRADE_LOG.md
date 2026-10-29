# UPGRADE_LOG — Excel Local Expert (nhánh muse-upgrade)

Môi trường kiểm chứng mặc định của các vòng do Muse chạy: VM Linux, Python 3.12.3,
KHÔNG có Excel Desktop, PowerShell, DPAPI, màn hình. Mọi test cần các thành phần
đó đều ghi "CHƯA CHẠY ĐƯỢC" kèm lý do, không bao giờ ghi đạt.
Môi trường đích của sản phẩm: Windows + Excel Microsoft 365 + Python 3.14.

## BƯỚC 0 — HIỆN TRẠNG (2026-10-30)

### Kết quả test gốc (chạy 2026-10-30, lệnh `python3 -X utf8 <file>` cho từng test)

| Test | Kết quả thật | Ghi chú |
|---|---|---|
| test_app.py | ĐẠT 79/79 | logic + trust boundary, không cần Excel |
| test_all_edge_cases.py | ĐẠT 61/61 | |
| test_upgrades.py | ĐẠT 28/28 | |
| test_cli.py | ĐẠT 6/6 | README ghi 4 — lệch tài liệu (xem VĐ-08) |
| test_batch_processing.py | CHƯA CHẠY ĐƯỢC | `execute_batch` → `app.bridge()` gọi `powershell.exe` (app.py:268) → FileNotFoundError trên Linux; cần Windows+Excel để kết luận |
| test_router.py | HỎNG (lỗi thật, mọi máy trừ máy tác giả) | test_router.py:30 đọc cứng `d:\Excel\excel-local\config.json` → FileNotFoundError |
| test_security_privacy.py | CHƯA CHẠY ĐƯỢC trọn vẹn | fail tại assert DPAPI (Windows-only) mà không skip; các check trước đó chạy được |
| test_headers.py | CHƯA CHẠY ĐƯỢC | cần powershell.exe |
| test_single_file_filter.py | CHƯA CHẠY ĐƯỢC | cần powershell.exe |
| test_diverse_scenarios_e2e.py | CHƯA CHẠY ĐƯỢC | cần powershell.exe |
| test_read_boundary.py | CHƯA CHẠY ĐƯỢC | cần powershell.exe |
| test_gui.py | CHƯA CHẠY ĐƯỢC | `_tkinter.TclError: no display` (headless) |

Tổng kiểm chứng được trên Linux: 79 + 61 + 28 + 6 = 174 check đạt.

### Điểm hiện trạng (1–10)

| Hạng mục | Điểm | Lý do |
|---|---|---|
| Độ đúng | 7 | 174 check logic đạt; luồng Excel thật chưa kiểm chứng được trong môi trường này |
| Bảo mật / quyền riêng tư | 6 | Có DPAPI + PrivacyGuard + pseudonym cột; nhưng PrivacyGuard regex chưa từng được đo tỷ lệ lọt/che nhầm, tài liệu dùng ngôn ngữ tuyệt đối vượt bằng chứng |
| Cấu trúc code | 4 | app.py 2.264 dòng, hàm `gui` ~890 dòng, `infer_smart_plan` ~478 dòng; khó bảo trì |
| Xử lý lỗi | 5 | 30 chỗ `except Exception` rộng (app.py 20, ai_client.py 10); vài test sập bằng traceback thô thay vì skip/báo rõ |
| Test | 6 | Nhiều test (tổng ~285 check theo tài liệu) nhưng là script rời, lẫn test thuần logic với test cần Excel/DPAPI, 1 test hardcode đường dẫn máy tác giả, file kết quả *.json lệch hash code hiện tại |
| Hiệu năng | 5 | benchmark-*.json ghi hash app.py/ai_client.py KHÔNG khớp bản hiện tại → số đo là của code cũ; nút thắt Excel COM chưa đo lại |
| Giao diện | 5 | Chưa kiểm tra trực quan được (headless); test_gui chưa chạy được ở đây |
| Tài liệu | 4 | README tự mâu thuẫn về cloud, lệch phiên bản v0.3.0/v0.3.1, lệch số thao tác 3/4, lệch số test CLI 4/6 |

### Danh sách vấn đề (đánh số, theo mức nghiêm trọng)

- VĐ-01 (Nghiêm trọng — tài liệu/bảo mật): README mâu thuẫn về cloud: dòng 51 nói "chỉ gọi API loopback, không chuyển sang cloud khi lỗi"; dòng 94 nói `request_plan` "chỉ gọi API Ollama local"; trong khi code có Smart Router gọi Gemini/Groq/OpenAI khi người dùng chọn provider smart/cloud. Người dùng có thể hiểu sai dữ liệu của mình đi đâu.
- VĐ-02 (Nghiêm trọng — tài liệu/bảo mật): Ngôn ngữ tuyệt đối không có bằng chứng đo: "Zero-Leakage", "tuyệt đối", "xóa bỏ mọi PII", "Offline 100%". PrivacyGuard là regex, chắc chắn có tỷ lệ lọt/che nhầm chưa đo.
- VĐ-03 (Trung bình — test/portability): Đường dẫn cứng `d:\Excel\excel-local\` trong test_router.py:30 và generate_project_pdf.py:7; test_router còn phụ thuộc config.json thật của máy tác giả (không hermetic).
- VĐ-04 (Trung bình — test): Test cần Windows/Excel/DPAPI/display không có cơ chế skip → sập traceback trên máy khác, che khuất tín hiệu hồi quy thật.
- VĐ-05 (Trung bình — xử lý lỗi): ~30 `except Exception` rộng trong app.py/ai_client.py; nguy cơ nuốt lỗi và khó truy vết.
- VĐ-06 (Trung bình — cấu trúc): Hàm quá dài: `gui` ~890 dòng, `infer_smart_plan` ~478, `prepare_sheets` ~261, `execute_batch`, `open_ai_settings` (lồng trong gui).
- VĐ-07 (Trung bình — hiệu năng/bằng chứng): benchmark-*.json lệch hash so với code hiện tại; chưa có đo phân giai đoạn mới.
- VĐ-08 (Nhẹ — tài liệu): README: đoạn cuối ghi v0.3.0 (đúng: v0.3.1); "ba thao tác được hỗ trợ" (đúng: 4 gồm clarify); test_cli ghi 4 (thật: 6); bước 1 cho chọn XLS trong khi mục Giới hạn nói chưa hỗ trợ XLS.
- VĐ-09 (Nhẹ — dọn repo): File kết quả/benchmark/script phụ nằm rải ở root (12+ file *.json kết quả, generate_project_pdf.py) gây khó đọc.
- VĐ-10 (Nhẹ — giao diện): Chưa rõ nút khóa khi đang chạy, tooltip, phím tắt, thứ tự Tab; chưa kiểm tra trực quan được.

### Backlog (thứ tự ưu tiên)

1. Vòng 1 — VĐ-01+VĐ-02+VĐ-08 (mục A+H của PROMPT): sửa README theo hành vi THẬT của code, bỏ ngôn ngữ tuyệt đối, sửa các lệch số/phiên bản.
2. Vòng 2 — VĐ-03 (mục B): bỏ đường dẫn cứng, test_router hermetic (config tạm).
3. Vòng 3 — VĐ-04 (mục G phần đầu): cơ chế skip rõ ràng cho test cần Windows/Excel/DPAPI/display.
4. Vòng 4 — Mục E: bộ test PrivacyGuard ≥50 mẫu, đo lọt/che nhầm, sửa nếu có số liệu.
5. Tiếp: VĐ-05 (except hẹp dần, từng chỗ có test), VĐ-06 (tách hàm), VĐ-07 (đo hiệu năng), VĐ-09 (dọn repo), VĐ-10 (giao diện nhẹ).

## NHẬT KÝ VÒNG

### Vòng 0 — Bước 0 chuẩn bị (2026-10-30)
- Tạo nhánh `muse-upgrade` từ main (28f5679).
- Đưa PROMPT_MUSE.md và run-muse-loop.ps1 vào root repo (script lặp yêu cầu prompt nằm trong repo).
- Chạy toàn bộ 12 file test_*.py, kết quả gốc ở bảng trên (174 check đạt; 1 test hỏng thật: test_router hardcode; còn lại CHƯA CHẠY ĐƯỢC vì thiếu Excel/PowerShell/DPAPI/display).
- Tạo UPGRADE_LOG.md, NEXT.md. Chưa sửa code ở vòng này.

### Vòng 1 — README đúng hành vi thật (VĐ-01, VĐ-02, VĐ-08) (2026-10-30)
- Kiểm chứng code trước khi sửa: đọc `ai_client.request_plan` (dòng 1261+) và `get_active_candidates` (1163+). Hành vi THẬT: mặc định config là `smart_local` (heuristic cục bộ); `ollama` chỉ loopback; `smart`/cloud cụ thể → cascade Gemini→Groq→OpenAI→OpenRouter/endpoint tùy chỉnh theo khóa đã cấu hình, lỗi thì fallback Ollama nếu `auto_fallback` (mặc định bật), cuối cùng heuristic. Prompt gửi đi = tên cột (bí danh nếu nhạy cảm) + yêu cầu đã che PII; không có dữ liệu hàng. `app.validate_plan` (app.py:634): action hợp lệ là lookup/summary/clean/filter/clarify; execute_batch chỉ nhận 4 action thực thi.
- Sửa README (chỉ tài liệu, không đổi code/chính sách): mở đầu và mục Runtime ghi rõ 2 chế độ local/cloud và dữ liệu gì đi cloud; đổi tiêu đề "Zero-Leakage" thành "Bảo mật & quyền riêng tư (mô tả đúng hành vi code, kèm giới hạn)"; PrivacyGuard ghi rõ là regex, có thể lọt/che nhầm, khuyên không gõ PII thật; DPAPI bỏ chữ "tuyệt đối", ghi rõ phạm vi bảo vệ; mục API/CLI viết lại đúng cascade; sửa v0.3.0→v0.3.1, "ba thao tác"→ danh sách đúng, test_cli 4→6, bước 1 bỏ XLS, bảng tính năng "bốn thao tác"→ liệt kê đúng.
- Kiểm tra sau: `grep` các cụm "Zero-Leakage / chỉ gọi API Ollama local / v0.3.0 / ba thao tác" trong README = sạch (các chữ "tuyệt đối"/"100%" còn lại đều là ngữ cảnh đúng: phủ định, dung sai toán học, stdlib). Chạy lại test_app.py: vẫn 79/79.

### Vòng 2 — Bỏ đường dẫn cứng, test_router hermetic (VĐ-03) (2026-10-30)
- Trước khi sửa: test_router.py HỎNG trên mọi máy trừ máy tác giả (đọc cứng `d:\Excel\excel-local\config.json`), và nguy hiểm hơn: test ghi thẳng vào config.json THẬT cạnh code rồi đặt lại `provider_keys` về rỗng ở cuối — chạy test có thể làm mất cấu hình khóa của người dùng.
- Sửa: (1) test_router chuyển `ai_client.CONFIG_FILE` sang thư mục tạm ngay đầu `run_tests()`, đọc raw file từ chính đường dẫn đó; cuối test chỉ dọn thư mục tạm, không đụng config thật. (2) Check "không có khóa thô trên đĩa" chỉ chạy khi DPAPI thật sự khả dụng (`dpapi_protect('probe')` bắt đầu bằng `dpapi:`), ngoài Windows in SKIP rõ lý do. (3) generate_project_pdf.py: `OUTPUT_DIR` đổi từ `d:\Excel\excel-local\output` sang `Path(__file__).parent / 'output'`. Các đường dẫn `C:\Program Files\...` còn lại là danh sách vị trí cài đặt chuẩn của Excel/Edge trên máy đích Windows (dữ kiện môi trường, không phải đường dẫn repo) nên giữ nguyên.
- Sau khi sửa (Linux): test_router ĐẠT cả 6 mục (trước: exit 1), không còn file config.json bị tạo trong repo; test_app 79/79, test_all_edge_cases 61/61, py_compile OK.
- Việc mới phát hiện khi soi lại (thêm vào backlog): **VĐ-11 (Nghiêm trọng — bảo mật): `ai_client.dpapi_protect` nuốt mọi lỗi và TRẢ VỀ PLAINTEXT khi DPAPI không khả dụng/thất bại** (ai_client.py:~140-167, `except Exception: return plaintext`). Trên Windows, nếu CryptProtectData lỗi, khóa API sẽ nằm thô trên đĩa mà không có cảnh báo nào. Cần vòng riêng: không lưu khóa khi mã hóa thất bại, hoặc cảnh báo rõ cho người dùng; có test.

### Vòng 3 — save_config không bao giờ ghi khóa thô (VĐ-11) (2026-10-30)
- Trước khi sửa: `save_config` ghi thẳng kết quả `dpapi_protect` xuống đĩa; khi DPAPI lỗi/không có, kết quả đó là plaintext → khóa nằm thô trong config.json, không cảnh báo. GUI `save_settings` (app.py) cũng không bắt lỗi lưu.
- Sửa: (1) `ai_client.save_config`: trước khi ghi, kiểm tra mọi khóa khác rỗng (openai_api_key + từng provider_keys) đều có dạng mã hóa `dpapi:`; nếu không, raise RuntimeError tiếng Việt nói rõ nguyên nhân và cách xử lý, và KHÔNG ghi file. Tương thích ngược: `get_config` vẫn đọc được config cũ dạng thô như trước. (2) `app.py save_settings`: bắt RuntimeError → `messagebox.showerror` và giữ hộp thoại mở, không báo "đã lưu" sai. (3) test_router: trên nền tảng không có DPAPI, cài cipher mô phỏng đảo ngược được (base64 + tiền tố `dpapi:`) để vẫn kiểm được hợp đồng lưu/đọc; thêm mục 7 — regression test: giả lập DPAPI hỏng (protect trả plaintext) thì save_config phải raise và file config nguyên vẹn, không chứa khóa.
- Sau khi sửa (Linux): test_router ĐẠT 7/7 mục (trước 6/6, chưa có regression); test_app 79/79; test_all_edge_cases 61/61; test_cli 6/6; py_compile app/ai_client/test_router OK; không tạo config.json trong repo.
- Hạn chế ghi rõ: hành vi DPAPI thật vẫn chỉ kiểm chứng được trên Windows; test_gui chưa chạy được ở môi trường headless này nên thay đổi messagebox trong save_settings mới kiểm ở mức compile + đọc code, CHƯA kiểm tra trực quan/tương tác.

### Vòng 4 — Skip rõ ràng theo nền tảng (VĐ-04) (2026-10-30)
- Trước: 7/12 file test sập traceback trên máy không có Excel/PowerShell/DPAPI/display (xem bảng Bước 0), che tín hiệu hồi quy.
- Sửa (chỉ file test, không đổi code sản phẩm, không bỏ check nào trên Windows):
  - test_security_privacy: mục 16–17 (DPAPI) chỉ chạy khi probe DPAPI thật cho kết quả `dpapi:`, ngược lại in SKIP; các mục logic còn lại chạy mọi nơi.
  - test_single_file_filter, test_batch_processing, test_diverse_scenarios_e2e: guard `sys.platform != 'win32'` ở đầu hàm chạy, in SKIP kèm lý do (Excel bridge), exit 0.
  - test_headers: phần logic (7 check) chạy mọi nơi; phần native Excel được guard + in số check logic đã đạt.
  - test_read_boundary: 2 check đối kháng (mock bridge) chạy mọi nơi; chỉ lượt đọc Excel thật cuối file được guard.
  - test_gui: bắt `tk.TclError` khi không có display → in SKIP + hướng dẫn, exit 0.
- Sau khi sửa (Linux): **cả 12/12 file test exit 0**. Check thuần logic đã kiểm chứng: test_app 79 + edge 61 + upgrades 28 + cli 6 + router 7 + security 38 = **219 check đạt**; phần bị SKIP đều in rõ lý do (5 check DPAPI của security khớp tổng 43 như README ghi cho Windows).

### Vòng 5 — Đo PrivacyGuard bằng bộ 53 mẫu (mục E) (2026-10-30)
- Tạo `test_privacy_guard_cases.py`: 33 mẫu PII phải che (SĐT mọi đầu số/dạng viết +84, 84 trần, chấm/gạch/khoảng trắng; CCCD/CMND; MST 10/13; STK 13/16/khoảng trắng; email; khóa sk-/Bearer; mật khẩu; tên kèm xưng hô; tên đứng một mình) + 20 mẫu không-phải-PII phải giữ nguyên (mã 00123, MSSV, mã lớp/môn, tiền, điểm, năm, số dòng/cột, số nội bộ, mã đơn/mã nghiệp vụ dài).
- Số đo TRƯỚC khi sửa: lọt bất ngờ 2/33 (SĐT dạng `84xxxxxxxxx` không dấu +; mật khẩu theo mẫu "… là <mật khẩu>"), che nhầm bất ngờ 0/20.
- Sửa tối thiểu trong ai_client: (1) PHONE_REGEX thêm nhánh tiền tố `84` trần; (2) PASSWORD_REGEX thêm mẫu "là/la" ngay sau từ khóa và cho phép tối đa 3 từ chen giữa từ khóa và "là".
- Số đo SAU: lọt bất ngờ 0/33, che nhầm bất ngờ 0/20. Chạy kèm: security 38, router 7, app 79, edge 61, upgrades 28, cli 6 — tất cả vẫn đạt.
- Giới hạn còn lại (đã đo, ghi trong test và README, không giấu): tên riêng đứng một mình không xưng hô — có dấu lẫn không dấu — đều lọt (regex không phân biệt được với cụm từ viết hoa thường); che nhầm theo thiết kế với dãy 9 số (trùng mẫu CMND) và mã 13 số (trùng mẫu STK). Muốn xử lý tên đứng một mình cần cơ chế khác (từ điển/danh sách tên hoặc NER) — ngoài phạm vi regex, ghi vào backlog dài hạn.

### Vòng 6 — Thu hẹp `except Exception` đợt 1 (VĐ-05) (2026-10-30)
- Phân loại đủ 30 chỗ (app.py 20, ai_client.py 10) thành: nuốt lỗi im lặng / bắt rộng nhưng có xử lý / ranh giới ngoài có thông báo. Đã thu hẹp 14 chỗ:
  - app.py + ai_client.py: guard `sys.stdout/stderr.reconfigure` (4 chỗ) → `(AttributeError, OSError, ValueError)`.
  - app.py `cleanup_stale_work_dirs`: → `OSError`.
  - app.py `terminate_active_bridges` (3 chỗ): → `(OSError, subprocess.SubprocessError)` và `(OSError, ValueError, subprocess.SubprocessError)` cho vòng đọc bản ghi tiến trình.
  - app.py GUI: 4 guard `pack/pack_forget` của nút Dừng → `tk.TclError`.
  - ai_client `get_config`: → `(OSError, ValueError)` (đọc file + parse JSON hỏng thì về mặc định).
  - ai_client phân tích thân lỗi HTTP: → `(ValueError, TypeError, AttributeError)`.
- Giữ nguyên 16 chỗ sau khi soi, có lý do: fallback DPAPI (đã có save_config chặn ghi thô từ Vòng 3), `load_xlsx_fast` trả None để rơi xuống đường Excel bridge, rollback xuất bản có re-raise, cách ly lỗi từng file trong batch (ghi vào báo cáo), callback tiến độ không được giết batch, ranh giới worker GUI/CLI (đều đưa lỗi ra giao diện), cascade router (ghi lỗi đã sanitize rồi chuyển provider), fallback Ollama → heuristic, `test_ai_connection` trả (False, thông báo).
- Kiểm tra sau: toàn bộ suite đạt như trước — app 79, edge 61, upgrades 28, cli 6, router 7, security 38, privacy corpus 53 mẫu (0 lọt/0 che nhầm bất ngờ). `grep -c "except Exception"`: app.py 20→10, ai_client.py 10→6.
- Lưu ý: 4 guard GUI mới chỉ kiểm ở mức compile + soi code (headless, test_gui SKIP như các vòng trước) — CHƯA kiểm tra trực quan.

### Vòng 7 — Tách hàm dài đợt 1 (VĐ-06) (2026-10-30)
- Mục tiêu: `ai_client.infer_smart_plan` 478 dòng. Tách thuần cơ học, 2 bước, test sau từng bước:
  - Bước 1: khối Q&A đầu hàm (4 mẫu clarify: giải thích lỗi #N/A, đếm cột, hướng dẫn, tính năng chưa hỗ trợ) → hàm `_qa_clarify_plan(p_lower, source_columns, reference_columns)`.
  - Bước 2: khối Summary Heuristics (tự chứa, chỉ phụ thuộc p_lower + source_columns) → hàm `_summary_plan(p_lower, source_columns)`.
- Kết quả: infer_smart_plan còn 375 dòng; hai hàm mới 47 và 70 dòng. Không đổi thứ tự đánh giá điều kiện, không đổi nội dung message/plan.
- Kiểm tra sau mỗi bước: router 7/7, security 38, app 79; sau bước 2 chạy thêm edge 61, upgrades 28, cli 6, privacy corpus — tất cả đạt như trước. Các test này chạm trực tiếp đường Q&A (router mục 3) và summary (router mục 5–6, heuristic rescue).
- Còn lại cho đợt sau: khối Lookup (~150 dòng, nhiều biến trạng thái) và Filter (~170 dòng) trong infer_smart_plan; `app.gui` ~890 dòng và `prepare_sheets` ~261 dòng — tách khi có test GUI/Excel chạy được để giảm rủi ro, hoặc tách cơ học từng khối nhỏ như đợt này.

### Vòng 8 — Đo hiệu năng phân giai đoạn, phần thuần Python (VĐ-07) (2026-10-30)
- Thêm `benchmark_python_stages.py`: đo lặp lại được các giai đoạn không cần Excel — preflight xlsx, đọc nhanh openpyxl (đường inspect lenient), inspect CSV, preview_plan lookup/summary/clean — ở 100/1.000/5.000/20.000 dòng, dữ liệu tổng hợp. Chạy 2 lần, số ổn định (chênh lệch nhỏ). Môi trường: VM Linux, Python 3.12, openpyxl 3.1.2.
- Số đo (giây; lần chạy lưu trong log này): tại 20.000 dòng — preflight 0,51; inspect nguồn (fast) 1,89; inspect danh mục (fast) 1,26; preview lookup 0,28; summary 0,18; clean 0,13; inspect CSV 0,11. Tại 5.000 dòng: preflight 0,12; inspect 0,44/0,28; preview 0,07/0,04/0,03. Tại 1.000 dòng tất cả ≤0,09.
- Kết luận có số liệu: các giai đoạn thuần Python tăng ~tuyến tính theo số dòng, preview ở quy mô tối đa (20.000 dòng) tốn <0,3 giây — KHÔNG có nút thắt thuần Python đáng tối ưu ở vòng này, nên không sửa code hiệu năng nào (đúng nguyên tắc: không có số đo cải thiện thì không nhận thay đổi).
- CHƯA ĐO ĐƯỢC ở môi trường này: toàn bộ giai đoạn qua Excel COM (inspect chuẩn, execute đối chiếu ~16 giây/5.000 dòng theo tài liệu cũ) — chi phí lớn nằm ở đó theo README (khởi động Excel ~5 giây bất kể kích thước). Muốn đo/tối ưu phải chạy `benchmark.py` trên Windows; các file benchmark-*.json cũ đã lệch hash (VĐ-07) và cần được đo lại trên Windows rồi mới so sánh. Ghi vào hàng đợi Windows ở NEXT.

### Vòng 9 — Dọn repo: gom file kết quả vào `results/` (VĐ-09) (2026-10-30)
- Chuyển toàn bộ 14 file *.json kết quả/benchmark từ root vào `results/` bằng `git mv` (giữ lịch sử): acceptance, benchmark-before/after/final/ai, test-results (+v0.1), cli/gui/header/read-boundary/cleanup-test-results, logic-regression.
- Cập nhật mọi nơi GHI các file này sang `results/` (tạo thư mục nếu thiếu): test_app.py, test_cli.py, test_gui.py, test_headers.py, test_read_boundary.py, benchmark.py, benchmark_ai.py, test_cleanup.ps1. README mục kiểm tra cũng trỏ sang `results/` và ghi chú benchmark cũ đo trên code cũ.
- Quyết định giữ nguyên vị trí: `create_diverse_examples.py` và `generate_project_pdf.py` ở root — chuyển vào tools/ sẽ làm vỡ `import app`/`app.ROOT` và đường dẫn output theo `__file__`, rủi ro lớn hơn lợi ích cho 2 script phụ trợ; ghi rõ để vòng sau cân nhắc lại nếu muốn.
- Kiểm tra sau: chạy lại toàn bộ 13 file test — tất cả exit 0; root không còn file *.json; file kết quả mới sinh ra nằm trong `results/` (đã khôi phục nội dung kết quả đã commit để vòng này chỉ gồm việc di chuyển, không lẫn nội dung kết quả chạy trên Linux).
