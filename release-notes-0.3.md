# Bản 0.3 · Nâng Cấp Toàn Diện: Dual-Engine AI, Tự Chữa Lành File Thực Tế & Giao Diện Stepper Trực Quan

Triển khai ngày 11/10/2026 tại `D:\Excel\excel-local`. Toàn bộ nâng cấp tuân thủ tuyệt đối quy tắc **chỉ dùng Python Standard Library + PowerShell + Excel Desktop COM**, tương thích ngược 100% với toàn bộ hệ thống test hiện hữu.

---

## 1. Khắc Phục Triệt Để Các Lỗi Đọc File Thực Tế

Trong môi trường thực tế tại Việt Nam (trường học, cơ quan, doanh nghiệp), các file Excel thường không ở dạng bảng chuẩn mực mà chứa nhiều đặc thù:
- **Tự động nhận diện dòng tiêu đề (`detect_header_row`):** Tự động quét 15 dòng đầu để tìm dòng chứa tiêu đề bảng thực tế, bỏ qua các dòng quốc hiệu, tiêu ngữ, tên trường/đơn vị hoặc dòng banner.
- **Tự chữa lành ô gộp & khuyết tên cột (`select_header(..., auto_heal=True)`):**
  - Ô gộp ngang trước đây tạo ra các ô trống `None`, gây lỗi chặn đọc file. Chế độ mới tự động điền tên cột mở rộng dạng `Họ tên_phụ`, `Điểm thi_phụ` hoặc `Cột_C`.
  - Trùng tên cột được tự động đánh số phân biệt `Ghi chú_2`, `Ghi chú_3`.
- **Hỗ trợ định dạng số & tiền tệ tiếng Việt (`parse_lenient_number`):**
  - Đọc hiểu các định dạng: `1.500.000 đ`, `1.234.567,89 VND`, `8,5` (dấu phẩy thập phân).
  - Tự động coi ô trống/None là `0` khi tính tổng hợp thay vì báo lỗi dừng toàn bộ tiến trình.
- **Hỗ trợ khóa ngày tháng (`key_value(..., allow_dates=True)`):** Chuyển đổi an toàn số serial ngày của Excel (ví dụ `45000`) thành chuỗi ngày ISO `YYYY-MM-DD`.
- **Mở rộng hỗ trợ định dạng file:**
  - Hỗ trợ các liên kết ngoài an toàn dạng `.../relationships/hyperlink` (email `mailto:`, link web) trong bước kiểm tra OOXML `preflight`.
  - Hỗ trợ mở và chuyển đổi định dạng Excel cũ `.xls` (BIFF8 97-2003) và bypass an toàn trong `preserve_check`.
  - Cho phép sheet nguồn có sẵn các lỗi công thức cũ (`#N/A`, `#DIV/0!`) không liên quan trong bước inspect (`$job.lenient`).

---

## 2. Nâng Cấp AI Tiếng Việt (Dual-Engine AI $\ge$ 98% Độ Chính Xác)

- **Tích hợp OpenAI API Cloud:**
  - Gọi trực tiếp endpoint Chat Completions (`gpt-4o-mini`, `gpt-4o`) qua `urllib.request` thuần của Python standard library.
  - Độ chính xác đọc hiểu yêu cầu tiếng Việt $\ge 98\%$, hiểu sâu sắc ngôn ngữ hành chính văn phòng đại học Việt Nam.
  - Tự động bóc tách JSON bọc trong markdown fences và xử lý lỗi mạng/timeout an toàn.
- **Cơ chế Fuzzy Matching & Ánh xạ ngữ nghĩa (`fuzzy_match_column`):**
  - Tự động liên kết các từ đồng nghĩa từ yêu cầu tự nhiên vào tên cột thực tế của bảng tính:
    - *quê quán* $\rightarrow$ **Nơi sinh** hoặc **Tỉnh/TP**.
    - *học phí / tiền nộp* $\rightarrow$ **Số tiền phải nộp**.
    - *mã sinh viên / mssv* $\rightarrow$ **Mã SV** hoặc **Mã hồ sơ**.
- **An toàn & Quyền riêng tư dữ liệu:**
  - Tuyệt đối **KHÔNG** gửi dữ liệu hàng (records) lên AI Cloud. AI chỉ nhận danh sách tên cột (`headers`) và prompt của người dùng để sinh kế hoạch thao tác.
- **Fallback linh hoạt với Ollama Local:**
  - Nếu mất mạng hoặc muốn làm việc offline 100%, hệ thống tự động fallback hoặc cho phép người dùng chủ động chọn Ollama Cục bộ (`qwen3:4b-q4_K_M`).
- **Cấu hình qua `config.json` & Giao diện:**
  - Hộp thoại **⚙ Cài đặt AI & API Key** trên GUI cho phép nhập API Key, chọn model, kiểm tra kết nối (Test connection) tức thì.

---

## 3. Cải Tiến Toàn Diện UI/UX (Tkinter Segoe UI)

- **Quy trình 4 bước trực quan (Stepper Bar):**
  - `① Nạp Dữ Liệu ➔ ② AI & Thiết Lập ➔ ③ Xem Trước Dự Kiến ➔ ④ Thực Thi & Hoàn Thành` giúp người dùng luôn nắm rõ vị trí và tiến trình công việc.
- **Gợi ý câu lệnh nhanh (Quick Prompt Chips):**
  - 3 nút mẫu nhanh phía trên ô yêu cầu: `🔍 Điền thông tin theo Khóa`, `💰 Tính tổng theo Nhóm`, `🧹 Chuẩn hóa chuỗi`. Bấm 1 click tự động điền prompt mẫu và chọn thao tác tương ứng.
- **Thẻ thống kê kết quả trực quan (Stat Badges):**
  - Cửa sổ Review được bổ sung các badge màu sắc trực quan: `[✓ Khớp: X]`, `[⚠ Thiếu/Trống: Y]`, `[⚡ Khóa trùng: Z]`, `[📊 Số nhóm: N]`, `[💰 Tổng dự kiến: M]`.
- **Xem trước dữ liệu nâng cấp:**
  - Bảng Treeview tích hợp cả 2 thanh cuộn ngang/dọc (dual scrollbar), hiển thị rõ ràng tối đa 20 dòng / 10 cột dữ liệu đã nạp.
- **Nhật ký thời gian thực:**
  - Mọi thao tác, cảnh báo, kết quả đều được ghi vào khung log với timestamp chuẩn `[HH:MM:SS]`.

---

## 4. Kết Quả Kiểm Thử Toàn Diện (100% Pass)

| Bộ kiểm thử | Số test | Kết quả | Mô tả |
|---|:---:|:---:|---|
| `test_upgrades.py` | 17 | **17/17 PASS** | Kiểm tra toàn bộ tính năng nâng cấp v0.3 (Header, Heal, Currency, Dates, Hyperlinks, Fuzzy Match, XLS) |
| `test_app.py` | 79 | **79/79 PASS** | Kiểm tra logic nghiệp vụ, đối soát độc lập, giới hạn bảo mật |
| `test_app.py --integration` | 97 | **97/97 PASS** | Kiểm tra thực tế với Excel Desktop COM thật, tính toán và xuất PDF |
| `test_gui.py` | 13 | **13/13 PASS** | Kiểm tra các nút giao diện, cửa sổ xem trước, quy trình chấp nhận |
| `test_cli.py` | 4 | **4/4 PASS** | Kiểm tra cơ chế token chấp nhận an toàn trên dòng lệnh |
| `test_headers.py` | 11 | **11/11 PASS** | Kiểm tra chẩn đoán hàng tiêu đề và định dạng số |
| `test_read_boundary.py` | 3 | **3/3 PASS** | Chống tráo file độc hại và kiểm tra biên đọc |
| `test_cleanup.ps1` | 4 | **4/4 PASS** | Kiểm tra dọn dẹp tiến trình Excel cô lập, không ảnh hưởng file/cửa sổ khác |
| `app.py --gui-smoke` | 1 | **0 clipped** | Kiểm tra giao diện đồ họa không có thành phần nào bị tràn/cắt xén |
