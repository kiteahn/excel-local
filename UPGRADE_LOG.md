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

### Vòng 10 — Giao diện nhẹ (VĐ-10) + TỔNG KẾT 10 VÒNG (2026-10-30)
Phần giao diện (chỉnh nhẹ, không đổi bố cục/luồng/tên tính năng):
- Soi code trước khi sửa: `work()` đã chặn thao tác chồng nhau bằng cờ busy + hộp thoại nhắc (mọi luồng Đọc/Đề xuất/Preview đều qua work()), nút Preview bị khóa khi chạy; nhưng CHƯA có tooltip nào, chưa có phím tắt nào, và nhãn quyền riêng tư ngay trong giao diện vẫn ghi tuyệt đối kiểu "Zero-Leakage / không bao giờ" (cùng lỗi VĐ-02).
- Sửa: (1) thêm lớp `_ToolTip` tối thiểu (stdlib Tk) và gắn cho 3 nút chọn động cơ AI (giải thích rõ local/cloud khớp README), ô Tự chữa lành, ô Xuất bảng gọn, ô Xuất PDF, ô dự phòng Ollama trong Cài đặt, và nút Preview (nhấn mạnh cổng chấp nhận). (2) Ctrl+Enter trong ô yêu cầu = bấm "AI Phân tích & Lập đề xuất". (3) Nhãn quyền riêng tư trong giao diện viết lại có giới hạn, khớp tài liệu.
- Tự soi lại bắt được lỗi của chính vòng này trước khi commit: `_ToolTip` ở cấp module nhưng `tk` được import lười bên trong `gui()` → sẽ NameError khi hover; đã sửa bằng import cục bộ trong `_show/_hide` và kiểm chứng bằng unit check widget giả (bind đăng ký đúng, TclError được nuốt êm).
- Kiểm tra sau: cả 13 file test exit 0 trên Linux (test_gui vẫn SKIP vì headless). **CHƯA kiểm tra trực quan** các tooltip/phím tắt trên Windows — nằm trong hàng đợi Windows ở NEXT.

#### Tổng kết 10 vòng — chấm lại điểm Bước 0
| Hạng mục | Bước 0 | Sau vòng 10 | Bằng chứng |
|---|---|---|---|
| Độ đúng | 7 | 7 | Không đổi hành vi nghiệp vụ; mọi sửa đổi đều có test đối chứng đạt như trước |
| Bảo mật / quyền riêng tư | 6 | 8 | Vá VĐ-11 (không ghi khóa thô), PrivacyGuard có số đo 53 mẫu (lọt bất ngờ 0), tài liệu + nhãn GUI hết ngôn ngữ tuyệt đối |
| Cấu trúc code | 4 | 5 | infer_smart_plan 478→375 dòng, tách 2 hàm có tên rõ; gui/prepare_sheets còn nguyên — điểm tăng vừa phải, trung thực |
| Xử lý lỗi | 5 | 6 | Thu hẹp 14/30 except có phân loại; 16 chỗ còn lại đã soi và ghi lý do giữ |
| Test | 6 | 8 | 13/13 file chạy sạch đa nền (SKIP có lý do), thêm corpus PrivacyGuard 53 mẫu + regression VĐ-11 + benchmark phân giai đoạn; tổng check logic kiểm chứng trên Linux: 219 + corpus |
| Hiệu năng | 5 | 6 | Có benchmark phân giai đoạn lặp lại được cho phần Python (tuyến tính, preview <0,3s ở 20k dòng); phần Excel COM vẫn chưa đo lại — chưa thể chấm cao hơn |
| Giao diện | 5 | 6 | Có tooltip + phím tắt + nhãn đúng sự thật; chưa kiểm tra trực quan nên không chấm cao |
| Tài liệu | 4 | 8 | README khớp hành vi thật về cloud/local, đúng phiên bản/số liệu, trỏ results/; còn release-notes cũ giữ nguyên tính lịch sử |

#### Backlog chặng tiếp theo (ưu tiên giảm dần)
1. Tầng 2 của PROMPT: liệt kê ≥30 use case và tự chạy phần chạy được trên Linux (CSV các biến thể qua inspect_csv/lenient, preflight từ chối macro/liên kết ngoài, giới hạn dòng/cột, khóa 00123...) — phần cần Excel ghi CHƯA CHẠY ĐƯỢC.
2. Hàng đợi Windows: benchmark lại bản hiện tại, test_gui + kiểm tra trực quan, nghiệm thu DPAPI.
3. Tách tiếp Lookup/Filter trong infer_smart_plan; gui; prepare_sheets.
4. Fuzz đơn giản cho parser CSV/header và mutation test cho bộ test (chế độ không cạn việc).

### Vòng 11 — Tầng kiểm tra 2: 36 use case, chạy được 32 (2026-10-30)
- Lập `USECASES.md` (36 use case) và `test_usecases.py` (32 check, tất cả đạt): CSV UTF-16/`;`/tab, banner CSV, giới hạn 60 cột/20.000 dòng, header trùng sau trim, auto_heal, ngữ nghĩa khóa (trim+casefold, NFC/NFD, 123 vs 00123 theo định dạng, khóa trống, ký tự `*`), lọc từ khóa/điều kiện số/khác rỗng, tổng hợp (tiền VN lenient, từ chối 1,234 chặt, trần 10^14, số âm, ô trống), khóa ngày serial, clean, cổng token của execute, preflight macro/externalLink.
- **Phát hiện và vá VĐ-12 (Trung bình — độ đúng/UX): CSV có banner hành chính bị từ chối hoàn toàn.** Trước đây `load_csv` strict loại mọi CSV lệch số cột, còn chế độ lenient lại CẮT dữ liệu theo bề rộng dòng đầu (banner 1 cột → mất cột thật) và inspect 'auto' vẫn dùng strict — nên tính năng "tự tìm dòng" không bao giờ dùng được cho CSV có banner, khác hẳn XLSX. Sửa tối thiểu: (1) lenient không còn cắt dòng dài — bề rộng lấy theo dòng rộng nhất, dòng ngắn được đệm '' ở cuối; (2) inspect CSV với header_row='auto' dùng đường tolerant để banner sống sót tới detect_header_row. UC-04 là regression test; toàn bộ suite cũ (gồm test_app với case ragged strict) vẫn đạt.
- Hai kỳ vọng ban đầu của chính vòng này bị dữ kiện bác bỏ và đã sửa thành ghi nhận thiết kế (không sửa code): UC-20b '500.000' một dấu chấm được hiểu là thập phân 500 (mơ hồ tiền VN — ghi vào USECASES.md chờ quyết định, không tự đổi hành vi số học); UC-26 khóa ngày serial không allow_dates thì bị từ chối kèm thông báo — đúng README.
- 4 use case cần Excel thật/PDF (UC-32, 33, 35, 36) và 1 từng phần (UC-34) ghi CHƯA CHẠY ĐƯỢC trong bảng, chuyển hàng đợi Windows.

### Vòng 12 — Fuzz nhẹ + mutation test (2026-10-30)
- Thêm `test_fuzz_light.py` (20 check): biến thể CSV (CRLF, thiếu newline cuối, dòng trống xen kẽ strict/lenient, ô đúng 32767 và 32768 ký tự, ký tự điều khiển 3 vị trí, phân cách hỗn hợp, file rỗng, BOM + ';', cp1258, nháy giữa ô) và header (toàn khoảng trắng, tên cột 200/1000 ký tự, NFC/NFD, detect trên dữ liệu đều nhau/rỗng). Hợp đồng: đọc đúng hoặc ValueError rõ ràng — đạt, không có ngoại lệ lạ nào.
- **VĐ-13 (Nhẹ — tài liệu): README ghi CSV "không đoán encoding cũ" nhưng `load_csv` thực tế thử cp1258/cp1252/latin1.** Đã sửa README theo code (kèm lưu ý nhánh báo lỗi encoding là gần như không tới được vì latin1 luôn giải mã thành công).
- **VĐ-14 (Trung bình — độ đúng): hai cột tên giống hệt nhau về hình thức nhưng khác chuẩn (NFC vs NFD) được chấp nhận là hai cột riêng** vì kiểm tra trùng dùng chuỗi gốc. Đã vá: `select_header` so trùng theo dạng NFC; F-15 là regression test.
- Mutation test 5 đột biến trên app.py (áp dụng tạm, chạy suite, khôi phục và đối chiếu bản khôi phục giống hệt bản gốc theo từng byte): MUT-1 bỏ casefold khóa — BẮT (test_app, test_usecases); MUT-2 bỏ chuẩn NFC khóa — BẮT (test_usecases, test_all_edge_cases); MUT-3 nới giới hạn 20.001 dòng CSV — BẮT (test_app); MUT-4 bỏ kiểm tra trường lạ trong validate_plan — BẮT (test_app); MUT-5 nới trần 10^14 của parse_lenient_number — **LỌT** (không test nào bắt; UC-22 chỉ phủ trần của `number()`). Đã bổ sung UC-29/UC-29b vào test_usecases và chạy lại: MUT-5 bị bắt.
- Sự cố quy trình của chính vòng này (ghi để tránh lặp lại): script mutation đầu tiên khôi phục bằng `git checkout -- app.py` trong lúc app.py còn thay đổi VĐ-14 chưa commit → làm mất tạm thời bản vá VĐ-14 và làm nhiễu 2 kết quả mutation. Đã phát hiện nhờ F-15 đỏ trên code "sạch", vá lại VĐ-14, và chạy lại mutation bằng khôi phục từ bản sao file + đối chiếu byte. Kết quả ghi ở trên là từ các lần chạy đã kiểm chứng áp dụng thật (git diff/bản sao).

### Vòng 13 — Tách khối Filter khỏi infer_smart_plan (VĐ-06 đợt 2) (2026-10-30)
- Kiểm tra phụ thuộc trước khi tách: khối Filter (mục 5, 175 dòng) chỉ dùng `prompt`, `p_lower`, `source_columns` và hàm module `extract_requested_columns` — không dính biến cục bộ của các khối khác. Tách thành `_filter_plan(prompt, p_lower, source_columns)`.
- infer_smart_plan: 375 → 204 dòng (từ 478 dòng ban đầu qua 2 đợt tách: _qa_clarify_plan 47, _summary_plan 70, _filter_plan 178).

### Vòng 14 — Tách nốt Clean + Lookup; VĐ-06 hoàn tất (2026-10-30)
- Tách khối Clean (20 dòng) thành `_clean_plan` và khối Lookup (167 dòng) thành `_lookup_plan` theo cùng mẫu cơ học. Ba đầu vào của Lookup đều là tham số của infer_smart_plan nên khối tự chứa hoàn toàn — không cần truyền thêm trạng thái.
- **infer_smart_plan: 478 dòng ban đầu → 25 dòng điều phối** + 5 helper: _qa_clarify_plan 47, _lookup_plan 170, _summary_plan 70, _clean_plan 23, _filter_plan 178. VĐ-06 đóng.
- Thêm `test_infer_golden.py`: 20 prompt vàng phủ cả 5 nhánh + trường hợp không khớp. Kiểm chứng hai chiều: bộ vàng đạt trên bản MỚI và trên bản TRƯỚC khi tách (lấy từ git HEAD) — 20/20 giống hệt, chứng minh cả 3 đợt tách (Vòng 7/13/14) không đổi hành vi. Toàn bộ 16 file test exit 0.
- Ghi nhận **VĐ-15 (Nhẹ — chất lượng gợi ý, chưa sửa)**: prompt "Điền tên sinh viên sang bảng này" cho lookup value='bảng' (regex cột đích bắt từ sau "sang"). Đây là hành vi có từ trước, đã ghim trong test vàng kèm chú thích; sửa heuristic là thay đổi hành vi nên cần vòng riêng có kiểm chứng riêng, không gộp vào vòng tách hàm.

### Vòng 15 — Mutation test đợt 2 cho ai_client.py (2026-10-30)
- 5 đột biến (quy trình Vòng 12 đã sửa: khôi phục bằng bản sao, đối chiếu byte): M2-1 bỏ che EMAIL — BẮT (security + corpus); M2-2 bỏ che số tài khoản — BẮT (security + corpus); M2-3 sanitize_plan ngừng ép action lạ về clarify — **LỌT**; M2-4 cổng lookup đổi `and`→`or` — **LỌT**; M2-5 sanitize_plan bỏ fuzzy match cột key — **LỌT**.
- Nguyên nhân chung của 3 đột biến lọt: sanitize_plan và cổng lookup chưa có unit test trực tiếp (chỉ được phủ gián tiếp qua router test dùng heuristic path). Đã bổ sung `test_sanitize_plan.py` (8 check) và 1 prompt vàng cho M2-4 ("Lấy riêng danh sách lớp trưởng" không có file tham chiếu phải ở lại filter). Chạy lại: cả 3 đột biến đều bị bắt. Tổng đợt 2: 5/5 sau khi bổ sung test.
- Trạng thái suite: 17 file test, tất cả exit 0.

### Vòng 16 — Vá VĐ-15: cột đích "bảng"/"file" lọt qua fuzzy passthrough (2026-10-30)
- Truy nguyên nhân gốc: không phải regex — regex trích đúng ứng viên 'bảng', và các tầng khớp exact/prefix đều từ chối đúng. Tầng cuối dùng `fuzzy_match_column`, mà hàm này **trả lại nguyên văn đầu vào khi không khớp cột nào** (passthrough, phục vụ sanitize_plan cho plan từ cloud). 2a coi kết quả truthy là đã khớp → target_dest='bảng'.
- Sửa tối thiểu đúng điểm lỗi: chỉ nhận kết quả fuzzy khi nó nằm trong `source_columns`. Trước khi sửa đã ghim 2 prompt vàng cho trường hợp cột đích THẬT ("vào cột Ghi chú", "sang cột Địa chỉ") — cả hai vẫn đúng sau khi sửa; 2 case VĐ-15 ('bảng', 'file') nay cho value=''. Bộ vàng: 21 → 24 case, toàn bộ 17 file test đạt.
- Quan sát thêm (chưa sửa, ghi VĐ-16): cùng mẫu passthrough còn ở nhánh m_key của _lookup_plan (khóa tra cứu do người dùng nêu) — rủi ro thấp hơn vì app.validate_plan kiểm tra lại tên cột khi thực thi, để vòng sau đánh giá. Ngoài ra 2 prompt không kích hoạt lookup chút nào: "Áp TTVR lên cột VR" (cụm kích hoạt cần 'áp lên' liền nhau, câu tự nhiên lại chen tên cột vào giữa) và "Cập nhật Điểm..." ('cập nhật' không nằm trong từ kích hoạt) — giới hạn heuristic có từ trước, ghi nhận, không sửa trong vòng này.

### Vòng 17 — Đánh giá prepare_sheets (không tách) + vá VĐ-16 (2026-10-30)
- **prepare_sheets (261 dòng): quyết định KHÔNG tách theo action.** Các khối lookup/summary/filter/clean không độc lập: chúng cùng sửa các biến dùng chung `rows` (bị thay thế ở summary/filter, mở rộng ở lookup), `formulas`, `stats`, `headers`, và đuôi chung dùng `selected_cols`/`filtered_indices` chỉ tồn tại ở nhánh filter. Tách sẽ phải truyền/trả 6–8 giá trị mỗi helper — rủi ro hồi quy cao, lợi ích chỉ là thẩm mỹ, trong khi hàm đã có test phủ dày (app 79 + edge 61 + usecases 34 đều đi qua prepare). Ghi quyết định này thay cho việc tách gượng; điểm "cấu trúc" giữ nguyên đánh giá.
- **VĐ-16a (passthrough ở m_key): đã vá** — khóa do người dùng nêu ("theo X") nay chỉ được nhận khi fuzzy khớp vào cột thật của mỗi phía; "theo xyz123" rơi về khớp ID như thiết kế dự phòng thay vì tạo khóa ma.
- **VĐ-16b (cụm kích hoạt 'áp … lên'): đã vá** — is_lookup nhận thêm mẫu regex `áp\s+\S+\s+(?:lên|vào)` khi có file tham chiếu. "Áp TTVR lên cột VR" nay cho lookup đúng (value='VR', ref_value='TTVR'); prompt âm tính "Áp dụng công thức tính điểm" vẫn không kích hoạt. Còn lại của VĐ-16: từ 'cập nhật' chưa phải từ kích hoạt lookup — giữ nguyên, ghi nhận.
- Bộ vàng infer: 24 → 28 case; toàn bộ 17 file test đạt.

### Vòng 18 — Rà soát tĩnh bằng pyflakes: bắt được 1 lỗi thật (2026-10-30)
- **VĐ-17 (Trung bình — GUI/lỗi bị nuốt): đường báo lỗi của Xử lý hàng loạt (Batch) trong GUI sẽ tự sập bằng NameError thay vì hiện thông báo lỗi.** pyflakes báo `undefined name 'exc'` tại app.py:2203: `self.window.after(0, lambda: _done(None, str(exc)))` — Python xóa biến except khi khối except kết thúc (PEP 3110), còn lambda chỉ chạy sau đó trên vòng lặp Tk. Đã tái hiện lỗi bằng mô phỏng tối thiểu (NameError thật) trước khi sửa. Sửa: gán `msg = str(exc)` và bind qua tham số mặc định của lambda. Đây là lỗi có từ bản gốc, không phải do các vòng trước gây ra; kiểm chứng trực quan trên Windows nằm ở hàng đợi.
- Dọn import thừa chắc chắn chết: `difflib` trong app.py và 16 import thừa trong 8 file test/tool (pyflakes sau dọn: 0 unused import, 0 undefined name trên toàn bộ file .py).
- Quét hàm chết theo AST + đối chiếu tham chiếu toàn repo: không có hàm top-level nào chết trong app.py/ai_client.py (4 hàm chỉ có 1 nơi gọi đều được xác minh là gọi thật, gồm test_ai_connection được GUI gọi).
- Nhất quán tài liệu: khối "Kiểm tra" trong README còn ghi test_router 6 mục (thực tế 7 từ Vòng 3) và thiếu 5 file test mới — đã cập nhật đủ 17 file với số đếm khớp lần chạy thật.
- Toàn bộ 17 file test vẫn exit 0 sau dọn dẹp.

### Vòng 19 — Đối chiếu toàn bộ VĐ-01…17 và số liệu tổng kết (2026-10-30)

Trạng thái từng vấn đề (đối chiếu log từ Bước 0 tới Vòng 18, không mục nào bị quên):

| VĐ | Nội dung ngắn | Trạng thái |
|---|---|---|
| 01 | README mâu thuẫn cloud/local | Đã sửa (Vòng 1) |
| 02 | Ngôn ngữ tuyệt đối về riêng tư | Đã sửa (Vòng 1; nhãn GUI ở Vòng 10) |
| 03 | Đường dẫn cứng + test_router không hermetic | Đã sửa (Vòng 2) |
| 04 | Test cần Windows không có skip | Đã sửa (Vòng 4) |
| 05 | ~30 except Exception rộng | Thu hẹp 14/30 (Vòng 6); 16 chỗ còn lại giữ kèm lý do đã ghi — đóng ở mức chấp nhận |
| 06 | Hàm quá dài | infer_smart_plan: xong (Vòng 7/13/14). prepare_sheets: đánh giá và không tách, có lý do (Vòng 17). gui ~909 dòng: CHƯA tách — việc còn mở lớn nhất về cấu trúc |
| 07 | Benchmark lệch hash / chưa đo phân giai đoạn | Phần Python: đã đo (Vòng 8). Phần Excel COM: hàng đợi Windows |
| 08 | Lệch số/phiên bản trong README | Đã sửa (Vòng 1; đồng bộ lại Vòng 18) |
| 09 | File kết quả rải ở root | Đã gom vào results/ (Vòng 9) |
| 10 | Tooltip/phím tắt/nhãn GUI | Code xong (Vòng 10); kiểm chứng trực quan: hàng đợi Windows |
| 11 | save_config có thể ghi khóa API dạng thô | Đã vá (Vòng 3); nghiệm thu DPAPI thật: hàng đợi Windows |
| 12 | CSV có banner bị từ chối/cắt dữ liệu | Đã vá (Vòng 11) |
| 13 | README sai về encoding CSV | Đã sửa tài liệu (Vòng 12) |
| 14 | Cột trùng NFC/NFD lọt qua | Đã vá (Vòng 12) |
| 15 | Cột đích lookup 'bảng'/'file' | Đã vá (Vòng 16) |
| 16 | Khóa ma m_key + cụm 'áp … lên' | Đã vá (Vòng 17); dư: từ 'cập nhật' chưa kích hoạt lookup — chấp nhận hiện trạng, có ghi |
| 17 | Batch GUI: lambda giữ exc → NameError | Đã vá (Vòng 18); kiểm chứng trực quan: hàng đợi Windows |

Số liệu đo thật tại vòng này: nhánh muse-upgrade có 19 commit (Bước 0 + Vòng 1–18) trên nền commit gốc 28f5679. app.py 2.264 → 2.330 dòng; ai_client.py 1.439 → 1.511 dòng (tăng do tách helper + guard + tooltip — độ dài hàm lớn nhất giảm mạnh dù tổng dòng tăng nhẹ). Suite: 17/17 file exit 0 trên Linux; cộng các bộ đếm đã in trong output thật: 79 (app) + 61 (edge) + 28 (upgrades) + 6 (cli) + 7 mục (router) + 38 (security, phần logic Linux) + 34 (usecases) + 20 (fuzz) + 28 (infer golden) + 8 (sanitize) + 53 mẫu (privacy corpus) + 24 (single-file) + 11 (headers) + 6 (diverse) + 10 (batch) + 2 (read boundary, phần Linux) ≈ **415 check logic đã kiểm chứng trên Linux**; phần cần Windows/Excel/DPAPI/display được SKIP có lý do, không tính là đạt.

### Vòng 20 — TỔNG KẾT 20 VÒNG (2026-10-30)

21 commit trên nhánh `muse-upgrade` (Bước 0 + Vòng 1–20, tính cả commit chứa tổng kết này), chưa push. Toàn bộ thay đổi đều có test chạy lại trên máy Linux không Excel; mọi khẳng định "đã sửa" trong log đều kèm bằng chứng chạy thật, và 4 lần kỳ vọng ban đầu của chính quá trình bị dữ kiện bác bỏ đều được ghi lại thay vì giấu (UC-20b, UC-26, F-14, sự cố khôi phục mutation ở Vòng 12).

Chấm lại 8 hạng mục (thang 10):

| Hạng mục | Bước 0 | Vòng 10 | Vòng 20 | Căn cứ vòng 20 |
|---|---|---|---|---|
| Độ đúng dữ liệu | 7 | 7 | **8** | Vá VĐ-12 (mất dữ liệu CSV), VĐ-14, VĐ-15/16 (khóa/cột ma), VĐ-17; phủ thêm fuzz + mutation 2 đợt |
| Bảo mật / riêng tư | 6 | 8 | **8** | Giữ nguyên: VĐ-11 đã vá + corpus 53 mẫu; nghiệm thu DPAPI thật vẫn chờ Windows nên không tăng |
| Cấu trúc code | 4 | 5 | **6** | infer_smart_plan 478 → 25 dòng + 5 helper; prepare_sheets giữ nguyên có lý do; gui ~909 dòng chưa tách nên trần ở 6 |
| Xử lý lỗi | 5 | 6 | **7** | Thêm VĐ-17 (đường báo lỗi tự sập) + hợp đồng fuzz "không ngoại lệ lạ" |
| Test | 6 | 8 | **9** | 17 file, ~415 check trên Linux; có test vàng, fuzz, mutation 10/10 bị bắt (2 đợt); trừ 1 vì phần Excel thật chưa chạy lại được ở đây |
| Hiệu quả (hiệu năng) | 5 | 6 | **6** | Không đổi: giai đoạn Python đã đo và không nghẽn; Excel COM chưa đo được |
| Giao diện | 5 | 6 | **6** | Không đổi: tooltip/phím tắt/VĐ-17 đều chưa kiểm chứng trực quan trên Windows |
| Tài liệu | 4 | 8 | **9** | Thêm USECASES.md (37 use case), README đồng bộ 2 lần theo số thật, log này là hồ sơ đầy đủ |

Việc còn dở, nói thẳng:
1. **Toàn bộ hàng đợi Windows chưa hề được chạm vào** vì môi trường này không có Excel/DPAPI/display: UC-32/33/35/36, đo lại benchmark Excel COM, kiểm tra trực quan tooltip (Vòng 10) và đường lỗi batch (VĐ-17), nghiệm thu DPAPI thật (VĐ-11). Đây là phần rủi ro tồn dư lớn nhất — code đã sẵn, bằng chứng thì chưa.
2. **gui() ~909 dòng chưa tách** (phần còn lại của VĐ-06) — cần test GUI thật trên Windows để tách an toàn, không nên tách mù từ Linux.
3. Ba quyết định chờ người dùng (ở NEXT.md): push nhánh hay không; chính sách cloud/local mặc định; '500.000' = 500 nghìn hay 500 thập phân.
4. Dư nhỏ đã chấp nhận: 16 `except Exception` giữ kèm lý do (VĐ-05); từ 'cập nhật' chưa kích hoạt lookup (VĐ-16 dư); tên riêng đứng một mình lọt PrivacyGuard (giới hạn regex đã đo và ghi trong README).

### Vòng 21 (hậu mốc 20) — Mutation test đợt 3 cho app.py (2026-10-30)
- 5 đột biến vào các cổng an toàn của app.py: M3-1 bỏ 'vbaproject' khỏi danh sách cấm của preflight — BẮT (test_usecases UC-30; lưu ý lớp [Content_Types] cũng chứa 'vba' nên đây là phòng thủ 2 lớp, đột biến bị bắt ở lớp tên file); M3-2 bỏ 'externallinks/' — BẮT (UC-31); M3-3 bỏ kiểm tra approval token trong execute — BẮT (test_app + UC-28); M3-5 bỏ kiểm tra DTD/entity — BẮT (test_app); M3-4 validate_plan cho thêm action lạ 'execute_anything' — **LỌT** (chưa test nào khẳng định validate_plan từ chối action ngoài danh sách ở tầng app; sanitize_plan có test tương đương nhưng là tầng ai_client khác).
- Đã bổ sung SEC-01/SEC-02 vào test_usecases.py (nay 36 check) và chạy lại: M3-4 bị bắt. Tổng 3 đợt mutation: 15 đột biến, tất cả đều bị bắt sau khi bổ sung test cho các ca lọt (5 + 5 + 5).

### Vòng 22 — Fuzz đợt 2 cho header (snapshot dict) + đóng băng tách gui() theo chỉ đạo người dùng (2026-10-30)
- Thêm `test_fuzz_headers.py` (15 check): ô header None, dòng toàn số, một dòng duy nhất, dòng lệch độ dài, NBSP, zero-width, banner + STT, auto_heal các dạng. Không có ngoại lệ lạ nào; không phát hiện lỗi code mới.
- Ba hành vi được ghim kèm nhãn "thiết kế hiện tại" để vòng sau khỏi tranh cãi lại: (1) kiểm tra trùng tên cột phân biệt hoa/thường ('ID' vs 'id' là 2 cột riêng — chỉ NFC + chuỗi gốc mới coi là trùng); (2) giới hạn 60 cột nằm ở inspect_file, không ở select_header (snapshot 61 cột qua select_header vẫn được nhận); (3) auto_heal chỉ cứu khi có ≥2 tên cột có nội dung (populated_count > 1) — dòng chỉ có 1 tên cột vẫn bị từ chối như chế độ chặt.
- Người dùng chỉ đạo ngày 2026-10-30: **không tách gui() cho tới khi có kết quả test trên Windows** — ghi vào NEXT.md làm điều kiện chặn chính thức.
- Suite: 18/18 file test exit 0.

### Vòng 23 — Áp dụng 2 quyết định chính sách của người dùng (2026-10-30)
- **Quyết định 1 — mặc định chỉ Ollama local, cloud là tùy chọn tự bật:** `get_config()` mặc định provider 'smart_local' → **'ollama'**; các điểm dự phòng rải rác được đồng bộ về 'ollama' (request_plan, biến động cơ GUI — trước đó fallback là 'smart', tức cloud; combobox Cài đặt). Thêm nút radio "🦙 Ollama Cục bộ (Mặc định)" vào hàng chọn động cơ (trước đây mặc định mới không có nút nào phản ánh). Cấu hình đã lưu của người dùng KHÔNG bị đụng (ai đã chọn smart/smart_local thì giữ nguyên — đó là lựa chọn đã bật của họ). Cloud vẫn đầy đủ trong danh sách chọn.
- **Quyết định 2 — '500.000' là tùy chọn cài đặt `vn_dot_thousands` (mặc định BẬT):** ở chế độ đọc số nới lỏng, một dấu chấm + đúng 3 chữ số sau dấu chấm + phần nguyên khác 0 → phân nhóm nghìn (500.000 → 500000; 12.345 → 12345; -500.000 → -500000); '0.500' luôn là thập phân; '1.2345' vẫn thập phân. Có checkbox trong Cài đặt AI kèm tooltip giải thích, lưu qua save_config. Luồng: config → `_vn_dot_thousands()` đọc 1 lần trong prepare_sheets → truyền vào number()/parse_lenient_number (không đọc config theo từng ô).
- Test mới `test_policy_defaults.py` (12 check): mặc định provider/vn_dot_thousands khi chưa có config, tôn trọng config đã lưu, 7 ca parse số, và 2 ca end-to-end qua prepare_sheets với config True/False (tổng 500100 vs 600). UC-20b được viết lại thành UC-20b/UC-20c theo quyết định mới (ghim tường minh cả hai trạng thái, không phụ thuộc config môi trường chạy test).
- Phát hiện phụ: phụ đề hộp thoại Cài đặt AI vẫn còn câu tuyệt đối "mã hóa DPAPI an toàn tuyệt đối" (sót từ Vòng 1/10) — đã sửa thành câu có phạm vi, kèm ghi rõ mặc định local.
- README cập nhật cả hai quyết định. Suite: 19/19 file exit 0.
- Quyết định 3 của người dùng: tự push nhánh — Muse không push; gói bundle/zip xuất lại sau vòng này.

### Vòng 24 — Mutation đợt 4: các cổng giới hạn còn lại (2026-10-30)
- 5 đột biến: M4-5 (bỏ điều kiện phần nguyên ≠ 0 của vn_dot_thousands) bị bắt ngay bởi test_policy_defaults. Bốn đột biến LỌT, đều là các cổng giới hạn chưa từng có test: M4-1 bỏ kiểm tra hash file nguồn sau khi đọc trong inspect_file; M4-2 nới giới hạn file 80MB → 800MB; M4-3 bỏ giới hạn chuỗi 2000 ký tự trong validate_plan; M4-4 bỏ giới hạn 10.000 thành phần ZIP trong preflight.
- Bổ sung `test_safety_limits.py` (6 check: file sparse 81MB bị từ chối / đúng 80MB được nhận; XLSX +10.001 thành phần rác bị từ chối; chuỗi 2001 ký tự bị từ chối / 2000 được nhận; giả lập file bị sửa giữa lúc đọc CSV bằng load_csv đã vá → inspect_file phát hiện qua hash). Chạy lại: cả 4 đột biến đều bị bắt.
- Tổng 4 đợt mutation: 20 đột biến, 20 bị bắt sau khi bổ sung test cho ca lọt. Rà soát tĩnh lại bằng pyflakes sau khi cài lại (VM reset làm mất gói): sạch.
- Suite: 20/20 file test exit 0. (Ghi chú: nhận định "việc Linux đã cạn" ở đây bị người dùng bác bỏ ngày 2026-10-30 — chế độ không cạn việc tiếp tục từ Vòng 25.)

### Vòng 25 — Mutation đợt 5: các module chưa từng đụng (2026-10-30)
- Người dùng chỉ đạo tiếp tục chế độ không cạn việc; vòng này nhắm vào cache plan, anonymize_columns, NAME_TITLE, detect_header_row scoring, key_value zfill — những chỗ 4 đợt trước chưa chạm.
- Kết quả: M5-2 (anonymize_columns ngừng ẩn danh) BẮT; M5-3 (bỏ che tên có xưng hô) BẮT; M5-5 (bỏ nhánh zfill khóa) BẮT. Hai ca LỌT: M5-1 cache key bỏ qua cột nguồn (hai bộ cột khác nhau dùng chung plan cache — lỗi đúng thật sự nếu xảy ra) và M5-4 bỏ điểm thưởng +25 "có dữ liệu bên dưới" của detect_header_row (dòng chân trang điểm cao có thể thắng dòng header thật).
- Bổ sung test: C-01..C-04 trong test_sanitize_plan.py (cache key khác theo cột nguồn/cột tham chiếu, không phụ thuộc thứ tự cột, và request_plan không dùng lại plan của bộ cột khác — ca C-04 phải sửa kỳ vọng một lần vì bộ cột thứ hai cũng cho lookup hợp lệ theo cột của nó; tiêu chí đúng là key/ref_value thuộc bộ cột thứ hai) và H-15 trong test_fuzz_headers.py (ca điểm thưởng quyết định: header thật thắng dòng footer). Chạy lại: cả hai bị bắt. Tổng 5 đợt: 25 đột biến, 25 bị bắt sau bổ sung.

### Vòng 26 — Fuzz đường đọc XLSX thuần Python (load_xlsx_fast + read_xml) (2026-10-30)
- Thêm `test_fuzz_xlsx.py` (12 check): file giả ZIP → None; ZIP không có cấu trúc workbook → None; ngày/datetime/boolean/ô trống giữ đúng ngữ nghĩa; chọn sheet theo tên và lùi về sheet đầu khi tên không tồn tại; sheet rỗng → rows rỗng; read_xml chặn DTD/entity và đọc được XML lồng 200 tầng; một ca end-to-end: inspect_file lenient + header 'auto' trên workbook thật có banner nhận đúng header dòng 3 (đây là đường XLSX duy nhất chạy trọn trên Linux nhờ fast reader).
- Không phát hiện lỗi. Hai ghi nhận: (1) `load_xlsx_fast` là đường đọc TÙY CHỌN phụ thuộc openpyxl — máy không có openpyxl thì hàm trả None và inspect lenient lùi về Excel bridge; bất biến "100% stdlib" của sản phẩm không bị phá (openpyxl không phải phụ thuộc bắt buộc) nhưng test file đã thêm SKIP khi thiếu openpyxl để khỏi đỏ giả trên máy sạch. (2) Chuỗi ngày tháng fast reader sinh ra ('YYYY-MM-DD') khớp định dạng key_value dùng cho khóa ngày — nhất quán giữa hai đường đọc.
- Suite: 21/21 file test exit 0.

### Vòng 27 — Rà soát tĩnh đợt 2 + đối chiếu README ↔ code (2026-10-30)
- pyflakes trên toàn bộ file .py (kể cả benchmark/tool phụ): sạch, trừ 3 mục lành tính được xác minh tay — `import openpyxl` trong benchmark_python_stages là nhịp đo/thăm dò khả dụng (giữ), 2 f-string thiếu placeholder trong generate_project_pdf.py (mỹ thuật, không ảnh hưởng output — giữ, tránh churn vào file sinh tài liệu). py_compile toàn bộ: OK.
- Đối chiếu README: mục Kiểm tra thiếu 4 file test mới (fuzz_headers, fuzz_xlsx, policy_defaults, safety_limits) và 2 số đếm đã cũ (usecases ghi 34, thật 37; sanitize ghi 8, thật 12) — đã bổ sung/sửa theo output chạy thật của vòng này. Kiểm tra chéo các khẳng định khác: release-notes-0.2.md được nhắc trong README có tồn tại thật; danh sách thao tác, giới hạn 60 cột/20.000 dòng/80MB khớp code.
- Quét hằng số giới hạn lặp giữa hai module chính: giới hạn prompt 8000 ký tự chỉ ở ai_client, giới hạn chuỗi plan 2000 chỉ ở app.validate_plan — không trùng lặp gây lệch; ghi nhận, không refactor.
- Suite: 21/21 file test exit 0.
- Kiểm chứng mạnh hơn các đợt trước: ngoài toàn bộ suite liên quan (router 7, security 38, app 79, edge 61, usecases 34, fuzz 20, privacy corpus, upgrades 28, cli 6 — tất cả đạt), còn so sánh trực tiếp hành vi: nạp song song ai_client bản trước vòng này (từ git HEAD) và bản mới, chạy 15 prompt đại diện cho cả 5 nhánh (Q&A, lookup, summary, clean, filter, không khớp gì) — kết quả JSON giống hệt 15/15.
- Kiểm tra sau: chạy lại toàn bộ 13 file test — tất cả exit 0; root không còn file *.json; file kết quả mới sinh ra nằm trong `results/` (đã khôi phục nội dung kết quả đã commit để vòng này chỉ gồm việc di chuyển, không lẫn nội dung kết quả chạy trên Linux).

### Vòng 28 — Use case bất thường đợt 2: phát hiện và vá VĐ-18 (2026-10-30)
- Thêm `test_usecases2.py` (10 check): CSV chỉ có header, tự đối chiếu với chính mình, filter selected_columns lạ, lookup vừa trùng vừa thiếu khóa, clean trên cột số, summary toàn khóa trống, tên sheet 100+ ký tự, lookup thiếu ref_value.
- **VĐ-18 (Trung bình — độ đúng/trải nghiệm): validate_plan âm thầm LOẠI BỎ cột trích xuất không tồn tại khỏi plan filter** thay vì báo lỗi — đánh máy sai tên cột ('Cot Khong Co') bị nuốt, người dùng nhận kết quả thiếu cột mà không biết; trong khi cột khóa/cột giá trị ở các action khác đều là lỗi cứng. Đã vá: cột không phân giải được (sau fuzzy) → ValueError nêu rõ tên cột; tên viết thiếu dấu vẫn fuzzy về cột thật bình thường. UC2-03 là regression test.
- Một kỳ vọng bị dữ kiện bác bỏ (ghi như UC-26 trước đây): UC2-07 tưởng filter '>= 8' trên cột chữ phải lỗi — thực tế hành vi là "ô chữ = không khớp điều kiện" kiểu autofilter của Excel, kết quả hiện rõ qua filtered_rows trong stats. Ghim hành vi thật (UC2-07/07b) thay vì ép sửa.
- Suite: 22/22 file test exit 0 (VĐ-18 không làm đỏ test cũ nào — các plan hợp lệ không đổi hành vi).

### Vòng 29 — Mutation đợt 6: formula/read_xml/filter/batch (2026-10-30)
- Bốn đột biến bị bắt ngay: M6-1 bỏ cổng SAFE_FUNCTIONS của formula_preflight, M6-2 bỏ chặn DTD trong read_xml (test_fuzz_xlsx X-11 bắt đúng tầng), M6-3 đổi >= thành > trong so sánh lọc số, M6-4 bỏ chặn tham chiếu ngoài của công thức.
- M6-5 (execute_batch đếm succeeded sai) LỌT: trên Linux mọi file batch đều dừng ở bridge Excel nên succeeded luôn 0 một cách "hợp lệ", test_batch_processing không phân biệt được. Bổ sung `test_batch_summary.py` (6 check): giả lập riêng khâu execute (khâu duy nhất cần Excel), để inspect/preview/vòng lặp batch/báo cáo chạy thật trên CSV — đếm 2 thành công/1 lỗi (file thiếu cột), trạng thái từng file + lý do lỗi, file báo cáo JSON khớp số đếm, và 3 cổng đầu vào (action clarify, thư mục rỗng, lookup thiếu tham chiếu). Chạy lại: M6-5 bị bắt.
- Tổng 6 đợt mutation: 30 đột biến, 30 bị bắt sau bổ sung. Suite: 23/23 file test exit 0.

### Vòng 30 — TỔNG KẾT 30 VÒNG (mốc phụ, 2026-10-30)
- Tổng kiểm chứng trên Linux (đếm từ output chạy thật của Vòng 29): **23 file test, 499 phép kiểm tra đạt** (so với 12 file/174 phép ở Bước 0). Sáu đợt mutation: 30/30 đột biến bị bắt sau khi bổ sung test cho ca lọt. Fuzz: 3 file chuyên biệt (CSV/parser, header, XLSX) + corpus Privacy Guard 53 mẫu.
- Chấm lại (thang Bước 0 → nay; giữ nguyên các điểm bị trần bởi điều kiện khách quan):
  - Độ đúng: 7 → 8 (VĐ-14/15/16/17/18 đều là lỗi hành vi thật đã vá kèm regression).
  - Bảo mật: 6 → 8 (VĐ-11 DPAPI fail-loud, corpus PII có số đo, cổng giới hạn đều có test mutation).
  - Cấu trúc: 4 → 6 (infer_smart_plan đã tách ở Vòng 14; trần hiện tại là gui() ~909 dòng — ĐÓNG BĂNG theo lệnh người dùng, không phải bỏ quên).
  - Xử lý lỗi: 5 → 7. Kiểm thử: 6 → 9. Hiệu năng: 5 → 6 (có benchmark tầng thuần Python; số Excel COM vẫn chờ Windows). GUI: 5 → 6 (chưa kiểm chứng trực quan). Tài liệu: 4 → 9.
- Đối chiếu VĐ-01…18: 17 mục đã vá + có test/regression (01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18 — liệt kê đầy đủ trong các vòng tương ứng); không mục nào bị quên. Riêng phần "chạy thật trên Windows" của VĐ-04/07/10/11 vẫn nằm trong nghiệm thu Windows (mã đã đúng theo thiết kế, bằng chứng chạy thật chỉ có trên máy người dùng).
- Việc Linux còn lại, theo giá trị giảm dần: (a) đọc tĩnh excel_bridge.ps1 — giá trị TRUNG BÌNH (bắt lỗi cú pháp/logic hiển nhiên phía PowerShell mà Linux không chạy được); (b) kiểm tra examples/ khớp hành vi hiện tại — giá trị TRUNG BÌNH-THẤP; (c) mutation đợt 7 cho các nhánh infer còn lại — giá trị THẤP (đã phủ 30 đột biến ở mọi cổng chính); (d) fuzz CSV đợt 3 với seed ngẫu nhiên dài hạn — giá trị THẤP. Sau (a)+(b), phần việc Linux chạm ngưỡng "chứng minh được không còn việc giá trị" theo tiêu chí người dùng đặt ra — khi đó vòng lặp nên dừng lại ở trạng thái chờ Windows, chỉ còn phản ứng khi có kết quả nghiệm thu hoặc chỉ đạo mới.

### Vòng 31 — Đọc tĩnh excel_bridge.ps1 + kiểm tra examples/ (2026-10-30)
- Đọc toàn bộ excel_bridge.ps1 (263 dòng; không chạy được PowerShell trên VM nên đây là đọc tĩnh, không phải kiểm chứng chạy thật). Không thấy lỗi cú pháp hiển nhiên; các cơ chế khớp với tuyên bố phía Python: AutomationSecurity=3 (tắt macro), Write-Grid đặt NumberFormat '@' TRƯỚC khi gán để chuỗi không biến thành công thức, đối soát độc lập có tolerance lấy từ job (0.000001, bất biến giữ nguyên), dọn tiến trình Excel chỉ khi khớp pid + đường dẫn + thời gian bắt đầu của đúng instance do bridge tạo.
- Ba ghi nhận để đối chiếu trên Windows (không sửa vì chưa đủ bằng chứng lỗi):
  - W-1: nhánh write không có input dùng `Workbooks.Add()` — output sẽ còn sheet mặc định trống của Excel. Trong mọi luồng hiện tại, execute luôn truyền input copy (app.py dòng 1046) nên nhánh này có vẻ chỉ là dự phòng; mục kiểm tra thủ công UC-32 ở nghiệm thu Windows cần nhìn danh sách sheet của file kết quả để xác nhận.
  - W-2: Write-Grid giả định sheet spec có ít nhất 1 dòng ($rows[0].Count) — mọi spec hiện tại đều có dòng header nên an toàn trong thực tế.
  - W-3: giới hạn inspect phía bridge (20001 dòng gồm header, 60 cột, 400.000 ô) nhất quán với cổng phía Python.
- examples/: kiểm tra 4 file mẫu ĐI KÈM REPO (2 XLSX có banner, 1 CSV, 1 XLSX kế toán) bằng inspect lenient + tự dò header — cả 4 cho header và số dòng đúng, chưa thối so với hành vi hiện tại. Sự cố quy trình đã xử lý: lần chạy thử create_diverse_examples.py ghi đè một phần examples/ TRƯỚC khi chết vì thiếu powershell.exe (script sinh ví dụ bằng Excel thật) — đã khôi phục nguyên trạng bằng git checkout và kiểm tra lại trên bản gốc; bài học: không chạy script sinh ví dụ trên VM này.
- Kết luận vòng lặp theo tiêu chí người dùng đặt: sau vòng này, việc Linux còn lại chỉ là (c)/(d) giá trị thấp đã nêu ở Vòng 30 — vòng lặp chuyển trạng thái CHỜ: nghiệm thu Windows, STOP.txt, hoặc chỉ đạo mới.

### Chu kỳ 0 — Nghiệm thu nền Windows (2026-10-30)
- Chạy đủ danh sách nhóm 1–3 từ checkout cô lập nhánh `muse-upgrade` tại commit `a9a1c5a`; log nguyên văn và các lần retry ở `cycles/CYCLE-00.md`.
- Lần chạy trong sandbox bị Excel COM từ chối (`0x80070005`); chạy lại ngoài sandbox thì thành công. Không có bằng chứng lỗi sản phẩm, không sửa code.
- Nhóm 1: 20/20 lệnh exit 0; 437 check đếm được và 6 kịch bản đa dạng hoàn tất. Nhóm 2: 50/50 check; nhóm 3: 4/4 lệnh, 114 check (97 integration, 14 GUI, 3 read-boundary). Benchmark COM 1.000 dòng ghi tại `results/benchmark-current-1000.json`.
- DPAPI thật, mặc định Ollama, `500.000` bật/tắt đều được xác nhận bằng test. Mở workbook kết quả read-only trong Excel xác nhận W-1: không có Sheet1 mặc định. PDF của integration test được render và xem; bảng hiển thị đầy đủ, không cắt/chồng.
- Một số luồng UI cần người dùng xác nhận còn trong `NEXT.md`; không ghi đạt khi chưa quan sát. Test thay đổi fixtures/results đã được khôi phục. Excel process do batch, benchmark và W-1 inspection sinh ra đã được xác định theo PID/path/start time, dọn theo PID và không còn tiến trình mồ côi.

### Tiếp tục kiểm tra UI theo yêu cầu (2026-10-30)
- Computer Use không khởi tạo được (`trusted Node process exited unexpectedly`, lặp lại sau reset kernel). UC-33 được hoãn trước khi tạo workbook test vì phát hiện một tiến trình Excel mở từ bên ngoài; không tương tác với cửa sổ/workbook đó. Cập nhật `NEXT.md`; các xác nhận giao diện vẫn chờ.

### UC-33 fix + handoff (2026-10-30)
- Probe synthetic workbook open in actual Excel found `inspect_file` returned data instead of the required locked-file message. Excel created sibling owner marker `~$source.xlsx`.
- Added minimal owner-marker guard for `.xlsx/.xlsm` and a regression check in `test_read_boundary.py`. Actual post-fix Excel probe now blocks with a clear Vietnamese message; `test_read_boundary.py` 4/4.
- Re-ran group 1 (20/20 commands, 437 counted checks + 6 scenarios), group 2 (50/50 checks), and post-fix group 3 integration 97/97, GUI 14/14, read-boundary 4/4. Did not rerun COM benchmark after code change per user's instruction to stop at handoff; existing result is pre-fix.
- Direct UI checks remain pending because Computer Use screenshot did not match the target window; no app input was sent. Handoff state is in `NEXT.md`. Known shortcut: a stale `~$` marker can block an unlocked file; remove only after confirming Excel is closed.
