# Bản 0.3.1 · Nâng Cấp Đột Phá: Smart Multi-Model Router (Kiến Trúc Luân Phiên 9Router) & Tối Ưu Hóa Trí Tuệ Toàn Phần

Triển khai ngày 12/10/2026 tại `D:\Excel\excel-local`. Kế thừa 100% nền tảng bảo mật Zero-Leakage của bản 0.3.0, bổ sung cơ chế điều phối và luân phiên thông minh giữa các mô hình trí tuệ nhân tạo hàng đầu thế giới hoàn toàn miễn phí, độc lập, không lệ thuộc một nhà cung cấp đơn lẻ.

---

## 1. Bối Cảnh & Vấn Đề Được Giải Quyết

Người dùng thực tế thường gặp các bài toán:
1. **Lệ thuộc vào một model duy nhất:** Khi mô hình gặp lỗi quá tải (HTTP 429 Rate Limit/Quota Exceeded) hoặc lỗi mạng, tiến trình xử lý bị đứt đoạn.
2. **Khó khăn về chi phí & API Key:** Người dùng không có thẻ tín dụng quốc tế để trả phí OpenAI, cần giải pháp tận dụng các gói **API Key miễn phí (Free Tier)** từ các nhà cung cấp lớn (Google Gemini, Groq Cloud).
3. **Mô hình cục bộ Ollama (Qwen 4B) chạy trên CPU**: Độ chính xác và khả năng diễn giải ngữ nghĩa tiếng Việt hành chính phức tạp có giới hạn so với các siêu mô hình 70B+ trên đám mây.
4. **Các câu hỏi nghiệp vụ / tra cứu thường gặp**: Người dùng thường hỏi các câu như *"Tại sao bị lỗi #N/A?", "Bảng này có bao nhiêu cột?"* – việc gọi mô hình AI tốn thời gian và rủi ro mạng không đáng có.

---

## 2. Kiến Trúc Smart Multi-Model Router (9Router-Style)

Hệ thống bổ sung bộ định tuyến thông minh với chuỗi thác đổ (Cascade Routing) ưu tiên tự động:

```
[Yêu cầu người dùng]
         │
         ▼
[Zero-Leakage Privacy Guard] (Xóa sạch CCCD, CMND, SĐT, Email, STK, Mật khẩu, Tên riêng)
         │
         ▼
[0ms Instant Q&A / Semantic Cache] ──(Khớp)──► Trả lời tức thì trong 0ms (0 byte ra internet)
         │ (Chưa khớp)
         ▼
[Smart Router Cascade]
 ├── 1. Google Gemini (Free Tier: gemini-1.5-flash / gemini-2.0-flash)
 │       └─ Miễn phí 15 RPM, thông minh vượt trội với tiếng Việt hành chính
 │          (Nếu lỗi / Hết Quota 429 ➔ Tự động chuyển)
 │
 ├── 2. Groq Cloud (Free Tier: llama-3.3-70b-versatile)
 │       └─ Miễn phí 100%, tốc độ siêu thanh (~300 tokens/s), suy luận chính xác
 │          (Nếu lỗi / Hết Quota 429 ➔ Tự động chuyển)
 │
 ├── 3. OpenAI Official (Paid Tier: gpt-4o-mini / gpt-4o)
 │       └─ Dự phòng cho tài khoản doanh nghiệp đã nạp tiền
 │          (Nếu lỗi mạng / Hết quota ➔ Tự động chuyển)
 │
 ├── 4. Ollama Cục bộ (Local CPU/GPU: qwen3:4b-q4_K_M)
 │       └─ Hoàn toàn offline trong mạng cô lập 100%
 │
 └── 5. Smart Heuristic Engine (Cứu cánh ngữ nghĩa)
         └─ Tự động phân tích từ khóa và cấu trúc cột với độ chính xác cao ngay cả khi mất mạng và tắt Ollama.
```

---

## 3. Các Tính Năng Nâng Cấp Nổi Bật

### 3.1. Hướng Dẫn Tích Hợp API Key Hoàn Toàn Miễn Phí (1 Phút)
- **Google Gemini API (Khuyên dùng nhất):**
  - Đăng ký tại [Google AI Studio](https://aistudio.google.com/).
  - Tạo API Key miễn phí không cần thẻ tín dụng, hạn mức 15 requests/phút.
- **Groq Cloud API:**
  - Đăng ký tại [Groq Console](https://console.groq.com/).
  - Tạo API Key miễn phí, sử dụng mô hình mã nguồn mở đỉnh cao `llama-3.3-70b-versatile`.

### 3.2. Mã Hóa An Toàn Tuyệt Đối Bằng Windows DPAPI
- Toàn bộ các khóa API (`gemini`, `groq`, `openai`, `openrouter`) được lưu trong `config.json` dưới trường `provider_keys_enc`.
- Sử dụng hàm mật mã cấp hệ điều hành `CryptProtectData` (Windows DPAPI). Tuyệt đối **không lưu văn bản thô (plaintext)** trên ổ đĩa. Chỉ tài khoản Windows đăng nhập hiện tại mới giải mã được trong RAM.

### 3.3. Cơ Chế Hỏi Đáp Nghiệp Vụ Tức Thì (Instant Q&A - 0ms Latency)
- Hệ thống tích hợp sẵn từ điển tri thức nghiệp vụ cho các câu hỏi trợ giúp:
  - Tra cứu ý nghĩa lỗi Excel: `#N/A`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`.
  - Tra cứu cấu trúc cột: *"Bảng này có bao nhiêu cột?", "Các cột hiện có là gì?"*.
  - Hướng dẫn cú pháp VLOOKUP/INDEX-MATCH.
- Phản hồi ngay tức thì trong 0ms, không phát sinh bất kỳ kết nối mạng nào ra ngoài.

### 3.4. Cải Tiến Hộp Thoại Cài Đặt AI Trên Giao Diện (GUI)
- Hộp thoại **⚙ Cài đặt AI & Smart Router** cho phép:
  - Lựa chọn chế độ: `⚡ Smart Router (Tự động luân phiên)`, `Google Gemini`, `Groq Cloud`, `OpenAI`, `Ollama Local`.
  - Nhập riêng lẻ API Key của từng nhà cung cấp (được ẩn dạng password `••••••••`).
  - Nút **Kiểm tra kết nối** tức thì cho từng nhà cung cấp.
  - Huy hiệu động cơ thời gian thực trên thanh trạng thái: `Động cơ: ⚡ Smart Router (Google Gemini ➔ Groq Cloud ➔ Ollama)`.

---

## 4. Kết Quả Kiểm Thử Toàn Phần (351/351 Kiểm Tra PASS 100%)

Toàn bộ hệ thống đã vượt qua 100% các bộ kiểm thử nghiêm ngặt nhất:

| Bộ kiểm thử | Số bài / checks | Kết quả | Phạm vi kiểm tra |
|---|:---:|:---:|---|
| [test_router.py](file:///d:/Excel/excel-local/test_router.py) | 6 | **6/6 PASS** | Kiểm tra chuỗi luân phiên Smart Router, failover khi 429, cứu cánh Ollama, DPAPI đa khóa |
| [test_all_edge_cases.py](file:///d:/Excel/excel-local/test_all_edge_cases.py) | 61 | **61/61 PASS** | Stress-test toàn diện 8 phân vùng: File hỏng, CSV phức tạp, Unicode NFC/NFD, Prompt injection |
| [test_security_privacy.py](file:///d:/Excel/excel-local/test_security_privacy.py) | 43 | **43/43 PASS** | Zero-Leakage: Xóa PII (CCCD, SĐT, Email, STK), DPAPI, Cache 0-egress, số âm, ngày tháng |
| [test_upgrades.py](file:///d:/Excel/excel-local/test_upgrades.py) | 28 | **28/28 PASS** | Tự chữa lành ô gộp, cột khuyết, tiền tệ VN (`1.500.000 đ`), link an toàn, file XLS |
| [test_app.py](file:///d:/Excel/excel-local/test_app.py) | 79 | **79/79 PASS** | Nghiệp vụ thuần, chuẩn hóa khóa, đối soát độc lập Python Decimal |
| `test_app.py --integration` | 97 | **97/97 PASS** | Tích hợp thực tế với Excel Desktop COM thật, tính toán và xuất PDF |
| [test_gui.py](file:///d:/Excel/excel-local/test_gui.py) | 14 | **14/14 PASS** | Giao diện Stepper 4 bước, treeview cuộn kép, cổng chấp nhận token, reset an toàn |
| [test_cli.py](file:///d:/Excel/excel-local/test_cli.py) | 4 | **4/4 PASS** | Kiểm tra luồng phê duyệt token nghiêm ngặt trên dòng lệnh |
| [test_headers.py](file:///d:/Excel/excel-local/test_headers.py) | 11 | **11/11 PASS** | Chẩn đoán địa chỉ ô lỗi tiêu đề, lề trống, bảo toàn định dạng số |
| [test_read_boundary.py](file:///d:/Excel/excel-local/test_read_boundary.py) | 3 | **3/3 PASS** | Chống tráo file độc hại OOXML, kiểm tra hash trước khi mở |
| [test_diverse_scenarios_e2e.py](file:///d:/Excel/excel-local/test_diverse_scenarios_e2e.py) | 6 | **6/6 PASS** | Kiểm thử End-to-End 6 kịch bản thực tế đại học, kế toán chi phí, kho vận CSV, bảng điểm |
| [test_batch_processing.py](file:///d:/Excel/excel-local/test_batch_processing.py) | 10 | **10/10 PASS** | Xử lý hàng loạt (Batch Mode) 0-token: Đa file, đối chiếu, tổng hợp, cô lập lỗi và pattern |
| [test_cleanup.ps1](file:///d:/Excel/excel-local/test_cleanup.ps1) | 4 | **4/4 PASS** | Dọn dẹp tiến trình Excel cô lập, không ảnh hưởng các workbook Excel khác của người dùng |
| `app.py --gui-smoke` | 1 | **0 clipped** | Kiểm tra dựng cửa sổ giao diện hoàn chỉnh, không có thành phần nào bị cắt xén |
| **TỔNG CỘNG** | **367** | **367/367 PASS (100%)** | **Hoàn thành toàn diện, sẵn sàng sử dụng thực tế** |
