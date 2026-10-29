# Excel Local Expert (v0.3.1)

Ứng dụng Windows xử lý bảng tính với Smart Multi-Model Router (luân phiên các nhà cung cấp người dùng đã cấu hình khóa: Google Gemini, Groq Cloud, OpenAI/OpenRouter hoặc endpoint tương thích, và Ollama Local), tính năng tự chữa lành ô gộp/tiêu đề và đối soát độc lập bằng native Excel Desktop. Người dùng tự chọn động cơ: chế độ cục bộ (heuristic/Ollama) không gửi gì ra ngoài; chế độ cloud gửi **tên cột và câu yêu cầu đã qua bước che PII bằng regex** tới nhà cung cấp đã chọn — không gửi dữ liệu hàng (xem mục "Bảo mật & quyền riêng tư" để biết giới hạn).

## Mở ứng dụng

Nhấp đúp `start.cmd` hoặc chạy `python app.py`. Máy hiện tại đã có Python 3.14 và Excel Microsoft 365. Không cần cài thêm bất kỳ thư viện ngoài nào (100% Python Standard Library).

1. Chọn **File nguồn** XLSX/CSV (định dạng XLS cũ chưa được hỗ trợ — xem mục Giới hạn). Bấm **🔍 Tự tìm dòng** nếu file có các dòng tiêu đề/banner phía trên. Bật tùy chọn **Tự chữa lành ô gộp & cột khuyết** để tự động xử lý ô gộp và cột trùng tên.
2. Với đối chiếu, chọn thêm **Danh mục đối chiếu** và các cột khóa/giá trị.
3. Chọn động cơ AI:
   - **⚡ Smart Multi-Model Router (Khuyên dùng):** Tự động luân phiên giữa các model AI miễn phí & trả phí chất lượng cao (Google Gemini Free Tier, Groq Llama 3.3 70B, OpenAI). Tự động failover sang model kế tiếp khi gặp giới hạn hạn mức (quota) và tự động fallback về Ollama Local khi mất mạng internet.
   - **🏠 Ollama Cục bộ (Offline):** Khi chọn chế độ này, ứng dụng chỉ gọi Ollama tại 127.0.0.1 trên máy, không cần internet.
4. Bấm **⚙ Cài đặt AI & Smart Router** để nhập API Key (được mã hóa bằng Windows DPAPI; chỉ tài khoản Windows hiện tại mới giải mã được — đây là lớp bảo vệ khi file cấu hình bị sao chép, không phải bảo vệ tuyệt đối nếu chính tài khoản/máy đã bị xâm nhập). Có thể dùng API Key hoàn toàn miễn phí từ Google AI Studio hoặc Groq Cloud.
5. Nhập yêu cầu tự nhiên rồi bấm **✨ AI Phân tích & Lập đề xuất**, hoặc chọn cột thủ công.
6. Bấm **3. Kiểm tra kết quả dự kiến (Preview)** để xem trước dữ liệu mẫu và các thẻ thống kê trực quan.
7. Bấm **Chấp nhận và thực hiện** để Excel native thực thi, tính toán và đối soát độc lập.


| Thao tác | Kết quả |
|---|---|
| Lọc & Trích xuất | Lọc dòng dữ liệu theo từ khóa/chức vụ (cán sự, lớp trưởng, tổ trưởng...), điều kiện số (>=, <=, >, <) hoặc không trống từ **1 file đơn lẻ** (KHÔNG CẦN file đối chiếu); hỗ trợ trích xuất riêng các cột mong muốn (Column Projection / Trích chọn cột); tự động bỏ qua banner hành chính ("Cộng hòa...", "Độc lập...") |
| Đối chiếu | Khớp khóa chính xác, công thức INDEX/MATCH từ mẫu cố định, liệt kê `KHOP`, `KHONG_TIM_THAY`, `KHOA_TRUNG`, `KHOA_TRONG` |
| Tổng hợp | Cộng một cột số theo nhóm; SUM/COUNT trên vùng riêng của từng nhóm trong Excel được so với tổng Decimal và số dòng tính độc lập bằng Python |
| Làm sạch | Bỏ khoảng trắng đầu/cuối và chuẩn NFC ở một cột, trên sheet kết quả; không xóa dòng trùng |
| AI tiếng Việt | Tự động phân tích yêu cầu ngôn ngữ tự nhiên thành một trong các thao tác lookup/summary/clean/filter (hoặc clarify để hỏi lại); tự động nhận diện ý định lọc cán sự/điểm số |
| Xem trước và chấp nhận | Dự kiến tính trong RAM; thay đổi file, dữ liệu, cột, PDF hoặc thư mục đầu ra làm mất hiệu lực bản chấp nhận |
| Xuất bản | Sau khi chấp nhận và đối soát: workbook XLSX mới, biên bản JSON, sheet kiểm tra và PDF tùy chọn |

Sheet nguồn trong XLSX vẫn nằm trong workbook kết quả. Sheet kết quả chứa ảnh chụp giá trị đã tính lại; các công thức nguồn vẫn được giữ tại sheet gốc. Công thức đối chiếu/tổng hợp mới tham chiếu dữ liệu nội bộ của workbook kết quả. Tác vụ không cập nhật tự động từ một file danh mục đã đổi: nạp lại dữ liệu rồi chạy tác vụ mới.

Khóa được chuẩn NFC, bỏ khoảng trắng đầu/cuối, không phân biệt hoa thường. Mã chuỗi `00123` giữ số 0 đầu; mã số với định dạng toàn số 0 như `00000` được đối chiếu theo mã hiển thị. Số `123` ở định dạng General khớp chuỗi `123`; không khớp `00123`. Khóa ngày/giờ dạng số phải chuyển thành chuỗi theo quy tắc rõ ràng trước.

Khóa danh mục xuất hiện nhiều lần luôn được báo trùng, kể cả các giá trị giống nhau; ứng dụng không tự lấy dòng đầu. Ký tự `*`, `?`, `~` được xử lý như ký tự thường khi tra cứu. Ô không tìm thấy giữ trống ở cột giá trị và có trạng thái riêng.

## Giới hạn v1

- XLSX/CSV tối đa 80 MB; XLSX tối đa 256 MB dữ liệu giải nén; tối đa 20.000 dòng dữ liệu, 60 cột và 400.000 ô được đọc trên sheet chọn.
- Một hàng tiêu đề, tên cột đầy đủ và không trùng (so trùng theo dạng chuẩn NFC, nên cùng một tên viết kiểu NFC và NFD cũng bị coi là trùng). CSV ưu tiên UTF-8 hoặc UTF-16 có BOM, dấu phân cách phẩy/chấm phẩy/tab; nếu không giải mã được bằng hai encoding đó, bộ đọc sẽ thử lần lượt cp1258/cp1252/latin1 (hành vi thực tế của `load_csv`, bản cũ của tài liệu ghi "không đoán encoding cũ" là không đúng với code).
- Số tổng hợp: ô số Excel hoặc chuỗi số dùng dấu chấm thập phân và không phân nhóm nghìn. Không tự đoán `1,234`, không tự đổi ô trống thành 0; từ chối số/tổng nhóm vượt 10^14 hoặc có hơn 15 chữ số có nghĩa. Dung sai đối soát số học tuyệt đối 0,000001; tác vụ dừng khi vượt dung sai.
- V1 từ chối sheet có giá trị lỗi Excel sau khi tính lại; không biến mã lỗi COM thành số. CSV không nhận ô dài quá 32767 ký tự hoặc ký tự điều khiển không hợp lệ.
- Chưa hỗ trợ XLS/XLSM, macro, OCR, Word, công thức tự do, cập nhật PivotTable, ghép nhiều-nhiều hay chọn bản ghi trùng theo quy tắc nghiệp vụ.
- Từ chối workbook có VBA/macro Excel 4.0, ActiveX, đối tượng nhúng, kết nối dữ liệu hoặc quan hệ ngoài; chính sách hiện tại cũng từ chối hyperlink ngoài. Kiểm tra cả công thức trong defined names. Chỉ nhận các hàm trong allowlist ở `app.py`; INDIRECT và tham chiếu bảng có ngoặc vuông chưa hỗ trợ.
- Bảo toàn được kiểm tra tự động ở mức giá trị/công thức của sheet gốc và hash file nguồn. Chưa chứng minh bảo toàn mọi kiểu định dạng, biểu đồ, PivotTable, chữ ký hoặc đối tượng trên mọi workbook; cần nghiệm thu bằng file đại diện trước khi mở rộng cam kết.
- PDF được xuất từ Excel, có trạng thái **chưa kiểm tra trực quan** trong biên bản. Bảng quá rộng vẫn cần điều chỉnh vùng in/bố cục.

## Runtime AI

Runtime portable: Ollama **0.35.1**, cổng **127.0.0.1:11435**; model **qwen3:4b-q4_K_M**, digest **2bfd38a7daaf4b1037efe517ccb73d1a3bbd4822cf89f1a82be1569050a114e0**. `runtime/runtime-lock.json` ghi hash binary và model đã chuẩn bị. Ứng dụng kiểm tra lock trước khi dùng.

`setup-runtime.ps1` dành cho giai đoạn chuẩn bị có mạng: tải bản phát hành chính thức đã khóa SHA-256 và model. **Đoạn này chỉ mô tả chế độ Ollama cục bộ:** ở chế độ đó, ứng dụng chỉ gọi API loopback (127.0.0.1:11435), không tự tải model, không dùng proxy/chuyển hướng và không tự chuyển sang cloud khi lỗi; `OLLAMA_NO_CLOUD=1` được đặt khi khởi động server Ollama. **Ở chế độ Smart Router hoặc khi chọn một nhà cung cấp cloud, ứng dụng CÓ gửi dữ liệu ra internet** — cụ thể là tên cột và câu yêu cầu của người dùng, sau khi đã qua bước che PII/bí danh cột (xem mục Bảo mật) — tới đúng nhà cung cấp đã cấu hình khóa; khi cloud lỗi/hết quota và tùy chọn `auto_fallback` đang bật (mặc định bật), ứng dụng mới rơi xuống Ollama cục bộ. Ở mọi chế độ, AI chỉ nhận tên cột và yêu cầu để lập kế hoạch; dữ liệu hàng không được đưa vào prompt.

Máy đã kiểm tra: RAM khoảng 16 GB, i5-12450H, RTX 2050 VRAM 4 GB, Excel Microsoft 365 16.0.20430.20092. Trong lần triển khai, `nvidia-smi` báo **GPU is lost**; Ollama đang dùng CPU. Không thay đổi driver hoặc khởi động lại máy trong tác vụ này. Sau khi người dùng khởi động lại và GPU hoạt động, runtime có thể cần khởi động lại để nhận diện GPU.

## Bảo mật & quyền riêng tư (mô tả đúng hành vi code, kèm giới hạn)

1. **Che PII bằng regex (Privacy Guard):** Trước khi gọi bất kỳ động cơ AI nào (cloud hoặc Ollama), hệ thống quét câu yêu cầu bằng biểu thức chính quy và che các mẫu nhận diện được: CCCD (12 số), CMND (9 số), SĐT VN (+84, 03x-09x), email, số tài khoản ngân hàng, mã số thuế, mật khẩu, một số mẫu tên riêng. **Giới hạn (đã đo bằng bộ 53 mẫu trong `test_privacy_guard_cases.py`):** đây là regex nên KHÔNG bảo đảm che hết — tên riêng đứng một mình, không kèm từ xưng hô (sinh viên/cô/anh...) thì không nhận diện được; định dạng lạ có thể lọt; ngược lại chuỗi số 9 chữ số (MSSV/số tiền) có thể bị che nhầm thành CMND và mã 13 chữ số thành số tài khoản. Đừng gõ thông tin định danh thật vào ô yêu cầu.
2. **Bảo vệ tiêu đề cột (Column Pseudonymization):** Tiêu đề cột chứa dữ liệu nhạy cảm được thay thế bằng bí danh `[COT_BAO_MAT_i]` khi gọi API và hoàn nguyên cục bộ khi nhận kết quả.
3. **Mã hóa khóa API bằng Windows DPAPI:** Khóa OpenAI API Key được mã hóa bằng chuẩn Windows Data Protection API (DPAPI). Không lưu văn bản thô trên ổ đĩa; chỉ tài khoản Windows hiện tại mới có thể giải mã.
4. **Bộ nhớ đệm kế hoạch cục bộ (Local Plan Cache):** Yêu cầu lặp lại cùng ý định được phục vụ ngay từ RAM trong 0ms, không phát sinh bất kỳ kết nối mạng nào ra ngoài (Zero Egress).
5. **Cách ly chỉ thị & Chống Prompt Injection:** Sử dụng ranh giới XML (`<danh_sach_cot_nguon>`, `<yeu_cau_nguoi_dung>`) và chỉ thị hệ thống cấm mô hình thực thi shell hoặc phá vỡ cấu trúc schema.
6. **Dữ liệu hàng không vào prompt:** Theo thiết kế, prompt gửi cho AI chỉ gồm tên cột (đã thay bí danh nếu nhạy cảm) và câu yêu cầu đã che PII; giá trị các hàng dữ liệu không được đưa vào prompt ở bất kỳ chế độ nào. AI chỉ trả JSON kế hoạch thuộc 5 loại: 4 thao tác thực thi (`lookup`, `summary`, `clean`, `filter`) và `clarify` (hỏi lại, không thực thi). Mọi công thức được sinh từ template cố định và đối soát độc lập trong native Excel Desktop.
7. **Tự động dọn dẹp file tạm:** Hàm `cleanup_stale_work_dirs()` tự động dọn sạch các thư mục tạm trong `.work/` khi ứng dụng khởi động.

## Kiểm tra và mã nguồn

```powershell
python -X utf8 test_single_file_filter.py # 24 bài test Lọc & Trích xuất file đơn lẻ, tự bỏ qua banner, cán sự & điểm số
python -X utf8 test_router.py             # 7 mục kiểm tra Smart Multi-Model Router (cascade & DPAPI; mục DPAPI cần Windows)
python -X utf8 test_all_edge_cases.py     # 61 bài test stress-test biên (Malformed, CSV, Unicode, Injection)
python -X utf8 test_security_privacy.py   # 43 bài test bảo mật PII, DPAPI, cache, số âm, ngày tháng
python -X utf8 test_upgrades.py           # 28 bài test tự chữa lành ô gộp, cột trùng, tiền tệ VN
python -X utf8 test_app.py                # 79 bài test logic, trust boundary và console UTF-8
python -X utf8 test_app.py --integration  # 97 bài test Excel Desktop COM thật, tính toán và xuất PDF
python -X utf8 test_gui.py                # 14 bài test kiểm tra giao diện Stepper, treeview, xem trước/hủy/chấp nhận
python -X utf8 test_cli.py                # 6 bài test kiểm tra cổng chấp nhận an toàn của CLI
python -X utf8 test_headers.py            # 11 bài test chẩn đoán tiêu đề, lề trống, định dạng
python -X utf8 test_read_boundary.py      # 3 bài test chống tráo file; lượt đọc Excel thật
python -X utf8 test_diverse_scenarios_e2e.py # 6 kịch bản thực tế đại học/kế toán/kho vận E2E
python -X utf8 test_batch_processing.py       # 10 bài test Xử lý hàng loạt (Batch Mode) 0-token
python -X utf8 test_usecases.py           # 34 use case tầng 2 (USECASES.md): CSV/banner, khóa, tổng hợp, token, preflight
python -X utf8 test_privacy_guard_cases.py # 53 mẫu đo Privacy Guard (33 PII + 20 không phải PII)
python -X utf8 test_infer_golden.py       # 28 prompt vàng cho infer_smart_plan (ghim hành vi heuristic)
python -X utf8 test_sanitize_plan.py      # 8 bài test sanitize_plan (ép clarify, fuzzy match cột)
python -X utf8 test_fuzz_light.py         # 20 biến thể fuzz nhẹ cho parser CSV/header
powershell -NoProfile -File test_cleanup.ps1 # 4 bài test dọn tiến trình Excel cô lập
```

Các file kết quả kiểm tra và benchmark nằm trong thư mục `results/` (ví dụ `results/test-results.json` ghi từng phép kiểm tra đã chạy). Đây không phải bằng chứng về mọi loại workbook hay khả năng chống mọi tấn công. Kết quả là dữ liệu giả, đáp án số học được tính độc lập.

Lượt triển khai 0.1 có **79 kiểm tra đạt**, được giữ ở `results/test-results-v0.1.json`. Kết quả nghiệm thu 0.2 và số đo hiệu năng nằm trong `results/acceptance-v0.2.json`, `results/test-results.json`, `results/gui-test-results.json`, `results/cli-test-results.json` và các file `results/benchmark-*.json` (lưu ý: các benchmark này đo trên mã nguồn cũ hơn, xem UPGRADE_LOG vòng 8). Các phép kiểm tra giao diện/CLI dùng mock ở bước thực hiện để xác nhận cổng chấp nhận; bộ integration riêng chạy Excel desktop thật. Xem `release-notes-0.2.md` để biết lỗi đã vá, số đo và giới hạn của phép đo.

Mã nguồn chính: `app.py`; API AI chỉ lập đề xuất: `ai_client.py`; bridge Excel: `excel_bridge.ps1`; chuẩn bị runtime: `setup-runtime.ps1`; mở giao diện: `start.cmd`. Không bổ sung thư viện ngoài, huấn luyện model hoặc dịch vụ cloud cho v1.

Tài liệu tham chiếu: [Ollama Windows](https://docs.ollama.com/windows), [Structured outputs](https://docs.ollama.com/capabilities/structured-outputs), [Ollama local-only](https://docs.ollama.com/faq), [Microsoft Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open).

## API AI và chấp nhận trên CLI

`ai_client.request_plan(prompt, source_columns, reference_columns)` trả JSON đề xuất và không mở/ghi workbook. Động cơ được dùng tùy lựa chọn của người dùng và cấu hình: heuristic cục bộ (không gọi mạng), Ollama local (chỉ loopback), hoặc Smart Router gọi lần lượt các nhà cung cấp cloud đã cấu hình khóa (Gemini → Groq → OpenAI → OpenRouter/endpoint tùy chỉnh), lỗi thì rơi xuống Ollama nếu `auto_fallback` bật. `app.ai_plan` kiểm tra đề xuất theo tên cột và các thao tác được hỗ trợ (lookup, summary, clean, filter; clarify chỉ để hỏi lại). `app.preview_plan` tính mẫu và số liệu dự kiến trong bộ nhớ. `app.execute` yêu cầu token khớp bản xem trước hiện tại; giao diện chỉ truyền token sau khi bạn bấm chấp nhận.

Các đề xuất trùng yêu cầu và tên cột được cache tối đa 16 mục trong RAM; mỗi lượt vẫn kiểm tra runtime/model. Cache không ghi xuống đĩa. Dữ liệu hàng không gửi vào model.

CLI mặc định in JSON với `status: awaiting_approval`, chưa tạo kết quả:

```powershell
python -X utf8 app.py --source "examples/ho-so-mau.xlsx" --reference "examples/danh-muc-mau.xlsx" --ai-prompt "Đối chiếu Mã hồ sơ, lấy Họ tên."
```

Sau khi kiểm tra JSON, lưu đúng đối tượng `preview.plan` thành `plan.json`. Chạy lại cùng nguồn/sheet/hàng tiêu đề và tùy chọn PDF, dùng `--plan plan.json --approve TOKEN`, với TOKEN là `preview.approval_token` đã xem. Thay đổi bất kỳ thành phần được chấp nhận sẽ bị từ chối và cần xem trước lại. `--ref-header-row` hỗ trợ hàng tiêu đề riêng cho danh mục.

Khởi động ứng dụng bằng `start.cmd` hoặc `python app.py` để trải nghiệm phiên bản **v0.3.1** với đầy đủ các nâng cấp mới nhất.
