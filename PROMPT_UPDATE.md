# PROMPT_UPDATE - PHẦN TIẾP THEO CỦA BỘ MÁY CÔNG TY
Bạn là kỹ sư trưởng của "Excel Local Expert", chạy trên máy Windows của chủ sản phẩm (có Excel thật). Bạn đã hoàn thành Chu kỳ 0-3 theo `PROMPT_COMPANY.md`.

**Quy tắc nền:** ĐỌC LẠI `PROMPT_COMPANY.md` và áp dụng NGUYÊN VĂN mọi phần trong đó: nguyên tắc sản phẩm bất biến, 7 pha của chu kỳ (Sản phẩm, Phát triển, Kiểm thử, Thử nghiệm, Nghiệm thu, Phát hành, Rà soát), quy tắc chống làm qua loa, cơ chế chạy liên tục (`UPGRADE_LOG.md`, `NEXT.md`, `ROADMAP.md`, `RELEASES.md`, `cycles/`), và `STOP.txt` là cách dừng duy nhất. File này CHỈ bổ sung: (1) việc dọn dẹp bắt buộc, (2) hàng đợi công việc chi tiết cho các chu kỳ tiếp theo, (3) cách bàn giao cho chủ sản phẩm. Khi hai file khác nhau về THỨ TỰ CÔNG VIỆC thì theo file này; về nguyên tắc an toàn thì theo `PROMPT_COMPANY.md`.

Vẫn giữ: chỉ làm trên nhánh `muse-upgrade`, KHÔNG push, mặc định Ollama local (cloud chỉ chạy khi người dùng tự bật), không dùng file thật của người dùng, không gọi cloud trong test, chỉ đóng Excel do mình/test mở.

# BẮT ĐẦU MỖI LƯỢT
Đọc `UPGRADE_LOG.md`, `NEXT.md`, `ROADMAP.md`, `cycles/CYCLE-04.md` (kế hoạch tách nhánh `lookup` khỏi `prepare_sheets`). Chạy `git status`, xác nhận nhánh `muse-upgrade` và cây làm việc sạch, rồi chạy nhóm test không cần Excel làm mốc. Làm tiếp đúng mục dở trong hàng đợi dưới đây.

# BƯỚC 0 - DỌN NỢ QUY TRÌNH (làm trước mọi thứ, trong Chu kỳ 4, mỗi mục một commit)
Các lỗi quy trình sau do chính các test gây ra; sửa tận gốc, không dùng cách "nhớ chạy lệnh khôi phục":
- **B0.1** Test ghi đè file mẫu Excel trong `examples/` (hiện phải `git checkout -- examples`). Sửa để test chép fixture sang thư mục tạm rồi mới ghi; thêm kiểm tra cuối test: `git status --porcelain examples` phải rỗng.
- **B0.2** `test_app.py` chạy không kèm `--integration` làm `results/test-results.json` còn 79 check. Sửa để file kết quả mỗi chế độ ghi ra một tên riêng (hoặc chỉ ghi khi chạy đủ), và README ghi rõ cách chạy.
- **B0.3** Excel COM thỉnh thoảng timeout một lần. Tìm nguyên nhân bằng cách tái hiện 20 lần liên tiếp, đo tần suất. Thêm cơ chế thử lại có giới hạn (ví dụ tối đa 2 lần, có chờ), chỉ cho các lỗi tạm thời, KHÔNG nuốt lỗi thật. Có test mô phỏng timeout và test xác nhận lỗi thật vẫn nổi lên rõ ràng.
- **B0.4** Viết `run_all_tests.ps1`: chạy tuần tự 3 nhóm (không cần Excel / DPAPI / Excel thật), in bảng tổng hợp đạt/hỏng/bỏ qua kèm số check, trả mã thoát khác 0 nếu có hỏng, kiểm tra không còn tiến trình Excel mồ côi do test sinh ra, và dọn thư mục tạm. Cập nhật README và `WINDOWS_ACCEPTANCE.md` dùng lệnh này.
- **B0.5** Quét bí mật trước mỗi lần phát hành: tìm mẫu khóa API (`AIza`, `sk-`, `gsk_`, `ghp_`...), đường dẫn cá nhân (tên người dùng Windows, `C:\Users\...`) trong repo và `git log -p`. Ghi kết quả vào `UPGRADE_LOG.md`.

# HÀNG ĐỢI CÔNG VIỆC (làm theo thứ tự; mỗi chu kỳ tối đa MỘT hạng mục lớn + các lỗi phát sinh)

## Q1 - Chu kỳ 4: tách nhánh `lookup` khỏi `prepare_sheets`
Điểm nóng đo bằng AST: `prepare_sheets` 262 dòng, `App.__init__` 213, `_lookup_plan` 180, `_filter_plan` 178, `gui()` 697.
- Trước khi tách: viết bộ test golden (ghim hành vi) cho `prepare_sheets` với ít nhất 20 kịch bản lookup, kiểm chứng hai chiều (bản mới và bản trước khi tách cho kết quả giống hệt, lấy từ git).
- Tách thành hàm nhỏ, từng lát một, chạy test sau MỖI lát. Chỉ tách nếu các nhánh không dùng chung quá nhiều biến trạng thái; nếu rủi ro lớn hơn lợi ích thì ghi lý do vào log và bỏ qua mục này (như đã làm với `prepare_sheets` trước đây).
- Nghiệm thu: AST cho thấy hàm giảm dòng đúng mục tiêu, mutation 5 chỗ trong phần vừa tách đều bị test bắt, test Excel thật vẫn đạt, benchmark không tệ hơn mốc cũ.

## Q2 - Chu kỳ 5: PrivacyGuard trên câu người dùng thật
Hiện bộ đo 80 PII + 37 non-PII còn nhỏ, còn 2 rò đã biết (tên 2 từ, tên dân tộc).
- Dựng kho ≥200 câu yêu cầu thực tế bằng tiếng Việt về bảng tính (lọc, đối chiếu, tổng hợp, làm sạch danh sách sinh viên, nhân sự, kho, kế toán, điểm danh, lương); chia PII thật và câu bình thường. Đặc biệt nhiều câu bình thường có từ giống họ hoặc tên riêng (ví dụ địa danh, tên môn học, tên hàng hóa, đơn vị hành chính) để đo che nhầm của luật nhận tên.
- Đo tỷ lệ lọt và che nhầm. Mục tiêu: lọt PII thật ≤ mức hiện tại, che nhầm trên câu bình thường ≤ 2%. Nếu luật nhận tên che nhầm nhiều, thu hẹp luật thay vì cố bắt thêm.
- Thêm cho người dùng khả năng XEM CHÍNH XÁC đoạn chữ sẽ được gửi lên cloud (đã che) trước khi gửi, và có nút hủy. Chỉ áp dụng khi người dùng tự bật cloud. Có test.
- README ghi rõ giới hạn còn lại; KHÔNG dùng từ "tuyệt đối", "không bao giờ lọt".

## Q3 - Chu kỳ 6: hiệu năng đo thật trên Excel COM
- Chạy lại `benchmark.py` và `benchmark_python_stages.py` sau thay đổi PrivacyGuard, so với mốc trong `results/` (đối chiếu 5.000 dòng ~16 s, tổng hợp ~14 s, `inspect_source` ~5 s).
- Phân rã thời gian từng giai đoạn (khởi động Excel, mở file, đọc, ghi, tính, đối soát). Chỉ tối ưu giai đoạn chiếm nhiều nhất. Thử ghi theo khối, giảm số lần mở/đóng Excel, tái dùng phiên Excel trong một lần xử lý.
- Mỗi thay đổi phải có số trước/sau (ít nhất 5 lần đo, lấy trung vị) và chứng minh đối soát số học vẫn đúng. Không giữ thay đổi nếu cải thiện dưới 10% mà phức tạp hơn.
- Đo thêm ở 20.000 dòng (sát giới hạn) và ghi rõ giới hạn.

## Q4 - Chu kỳ 7: dữ liệu giống thật và độ chính xác lập kế hoạch
- Viết bộ tạo dữ liệu mẫu (`examples/real_like/`, không chứa dữ liệu thật) cho ≥6 loại bảng tính kiểu Việt Nam: danh sách sinh viên có banner hành chính, bảng lương, kho hàng, điểm danh, sổ quỹ, đơn hàng; có các lỗi thường gặp (mã giữ số 0, tiền kiểu 1.234.567, ngày đủ kiểu, dấu NFC/NFD, ô gộp, dòng tổng, ô lỗi, khóa trùng).
- Dựng ≥50 câu yêu cầu thực tế cho các bảng đó, chạy toàn bộ luồng (đọc, đề xuất, xem trước, thực thi, đối soát), đo tỷ lệ kế hoạch đúng ngay lần đầu. Mọi trường hợp sai thêm vào bộ golden `test_infer_golden.py` kèm cách sửa.
- Mục tiêu: nêu tỷ lệ đúng bằng số thật; không hứa hơn số đo.

## Q5 - Chu kỳ 8: giao diện (chỉnh nhẹ, có ảnh chụp)
- Chụp ảnh từng hộp thoại và màn hình chính bằng cách ép cửa sổ Tk `-topmost` + `lift()` (đã biết cách), lưu vào `docs/screens/`. Ghi rõ: ảnh chụp để CHỦ SẢN PHẨM xem; bạn không tự đánh giá "đẹp".
- Chỉ sửa những gì đo được: chữ bị cắt, nút chồng nhau, tab order, độ tương phản, kích thước cửa sổ nhỏ, thông báo lỗi tiếng Việt dễ hiểu, trạng thái đang xử lý và khóa nút khi chạy. Không đổi bố cục tổng thể hoặc luồng thao tác.
- Mọi thay đổi giao diện: chạy `test_gui.py` và `test_ui_acceptance.py`, và ghi vào `NEXT.md` mục "CẦN NGƯỜI XÁC NHẬN BẰNG MẮT".

## Q6 - Chu kỳ 9: tách tiếp `App.__init__`, `_filter_plan`, `gui()` (nếu Q1 thành công)
Cùng quy trình golden, từng lát nhỏ, mutation, đo AST trước/sau. Dừng ở chỗ rủi ro lớn hơn lợi ích và ghi lý do.

## Q7 - Chu kỳ 10: nền tảng phát hành
- Thêm `.github/workflows/tests.yml` chạy nhóm test không cần Excel trên Windows và Ubuntu (CHỈ tạo file, không push). Kiểm tra cú pháp YAML và rằng các test chọn chạy được không cần Excel.
- KHÔNG tự thêm `LICENSE`: đây là quyết định của chủ sản phẩm. Ghi câu hỏi vào `NEXT.md` kèm 2-3 lựa chọn phổ biến (MIT, Apache-2.0, giữ riêng tư), ưu/nhược ngắn gọn; làm việc khác.
- Cập nhật `README.md` (cài đặt, yêu cầu Windows + Excel + Python, cách chạy, giới hạn đã công bố, chính sách cloud/local, cách chạy test) và `RELEASES.md`; tăng phiên bản; tạo tag cục bộ.

## Sau Q7
Quay vòng chu kỳ 5 lại với góc nhìn mới, theo "chế độ không cạn việc" trong `PROMPT_COMPANY.md` (fuzz lớn hơn, mutation đợt mới, đọc lại mã như người mới, rà soát nợ kỹ thuật). Không bao giờ tự kết luận sản phẩm đã hoàn thiện.

# CỔNG NGHIỆM THU (áp dụng cho MỌI chu kỳ)
- `run_all_tests.ps1` đạt toàn bộ nhóm, không test nào bị tắt/làm yếu; số check ghi đúng với kết quả thật.
- Tiêu chí riêng của chu kỳ (nêu trong `cycles/CYCLE-NN.md`) đạt bằng số đo thật.
- Nguyên tắc sản phẩm không bị phá; `git status` sạch (không còn file mẫu/kết quả bị ghi đè); quét bí mật sạch.
- Không đạt thì quay lại sửa tối đa 3 lần, quá thì hoàn tác phần đó và ghi lý do.
Mục nào cần mắt người (giao diện, PDF, hộp thoại) thì ghi "CẦN NGƯỜI XÁC NHẬN", KHÔNG báo đạt.

# BÀN GIAO CHO CHỦ SẢN PHẨM (cuối MỖI lượt chạy)
Chủ sản phẩm là sinh viên, không phải lập trình viên chuyên nghiệp. Ở ĐẦU file `NEXT.md` viết mục "TÓM TẮT CHO CHỦ SẢN PHẨM" (tiếng Việt đơn giản, tối đa 12 dòng):
1. Lượt này đã làm xong gì (3-5 gạch, nói điều người dùng sẽ thấy, không dùng thuật ngữ khó).
2. Test: đạt/hỏng/bỏ qua và con số thật.
3. Việc CẦN BẠN LÀM BẰNG TAY (kèm từng bước, từng lệnh).
4. Câu hỏi đang chờ bạn quyết định (kèm đề xuất).
5. Rủi ro/điều chưa kiểm chứng.
Cứ sau 3 chu kỳ: cập nhật `WINDOWS_ACCEPTANCE.md`, xuất gói bàn giao vào thư mục `handoff/` (git bundle + zip + patch) và ghi đường dẫn vào `NEXT.md`.

# BẮT ĐẦU NGAY
Đọc các file ở mục "Bắt đầu mỗi lượt", làm Bước 0 (B0.1 đến B0.5) rồi Q1, và không dừng cho tới khi có `STOP.txt`.
