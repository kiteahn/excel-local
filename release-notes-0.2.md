# Bản 0.2 · chấp nhận thao tác, vá lỗi và hiệu năng

Đã triển khai ngày 06/10/2026 tại D:/Excel/excel-local. File dự án gốc không bị sửa.

## Luồng làm việc

Yêu cầu → AI trả đề xuất → xem trước số liệu và mẫu kết quả → người dùng chấp nhận → Excel thực hiện, đối soát và tạo file mới. Mỗi tác vụ mới đi qua bước xem trước/chấp nhận; ứng dụng không tự chuyển sang công đoạn khác. AI hỏi làm rõ nếu yêu cầu thiếu thông tin.

API nằm riêng trong `ai_client.py`, chỉ trả đề xuất. Cửa sổ xem trước có nút quay lại sửa yêu cầu và nút **Chấp nhận và thực hiện**. Kiểm tra hash, dữ liệu đã đọc, cột, thư mục đầu ra và PDF bảo đảm bản chấp nhận khớp thao tác sắp chạy. CLI mặc định cũng chờ chấp nhận, không tự ghi kết quả.

## Lỗi đã vá

- Danh mục cũ đã đổi không còn chặn làm sạch/tổng hợp vốn không dùng danh mục.
- Các nhóm `001`, `1`, `=a`, `>a`, `a*b` được tách chính xác; SUM/COUNT chạy trên vùng riêng từng nhóm, tránh ép tiêu chí của COUNTIF và quét lại toàn bộ dữ liệu cho mỗi nhóm.
- CSV giữ nguyên ô nhiều dòng có dấu nháy; chặn dấu nháy chưa đóng và kiểm tra giới hạn trong lúc đọc hàng. Không tạo thêm danh sách toàn bộ các dòng text.
- Xác minh mọi ô của sheet mới sau SaveAs, gồm kiểu bool/số/text, công thức và giá trị đã lưu. Lỗi đối soát ngăn xuất bản.
- Giải mã đúng chuỗi OOXML được escape, cặp surrogate Unicode và bỏ phần phiên âm; kiểm tra chuỗi công thức vẫn là dữ liệu.
- Chặn DTD/entity cả khi XML dùng UTF-16. Kiểm tra hash và nội dung của đúng bản sao trước khi Excel mở để chặn tráo file giữa các bước.
- Hết thời gian có nhánh dọn đúng tiến trình của tác vụ. Từ chối sử dụng phiên Excel có sẵn; không Quit/kill cửa sổ của người dùng. Ownership, PID, thời điểm và đường dẫn phải khớp khi dọn.
- Tiêu đề cột `#0`, hàng tiêu đề nhập sai và lỗi callback không làm hỏng vòng cập nhật giao diện. Cửa sổ chờ worker hoàn thành trước khi đóng.

## Số đo

| Thao tác · 1.000 dòng | Trước | Bản 0.2 | Tăng tốc |
|---|---:|---:|---:|
| Đối chiếu | 20.20s | 7.77s | 2.60× |
| Tổng hợp | 23.98s | 8.70s | 2.76× |
| Làm sạch | 11.38s | 7.54s | 1.51× |

Đo toàn bộ `execute` sau khi xem trước, gồm kiểm tra chấp nhận, ghi Excel, đối soát sau lưu, bảo toàn nội dung sheet nguồn và xuất bản XLSX/JSON; không yêu cầu PDF trong benchmark. 1.000 dòng có 1.000 khóa và 1.000 nhóm, 1.000 công thức đối chiếu hoặc 2.000 công thức tổng hợp.

Bộ 5.000 dòng: đối chiếu **15.83s**, tổng hợp **13.60s**, làm sạch **11.11s**. Đối soát 5.000 và 10.000 công thức tương ứng. Mọi dữ liệu benchmark là dữ liệu giả; các file tác vụ được dọn sau phép đo.

Bridge ghi công thức/định dạng theo vùng liên tiếp, đọc giá trị đối soát theo vùng và tạm tắt tính toán tự động trong lúc ghi. Chế độ tính toán ban đầu được khôi phục trước khi lưu.

AI: một yêu cầu benchmark từ **19.46s** xuống **7.74s**, lặp lại **0.173s**. Prompt bỏ schema bị lặp; cache tối đa 16 đề xuất trong RAM, vẫn kiểm tra runtime/model mỗi lượt. Lượt AI mới trong suite đầy đủ mất 17,62s; lượt lặp 0,23s. Máy đang chạy CPU, nên độ trễ phụ thuộc tải và trạng thái model. Các số đo là từng lượt, không phải trung bình hay cam kết độ trễ.

## Bằng chứng kiểm tra

**122 kiểm tra đạt**: suite chính 101 (Excel desktop thật và AI local), giao diện 10, CLI 4, dọn tiến trình Excel thật 4, kiểm tra bản sao đầu vào 3. Các test giao diện/CLI mock bước thực hiện để chứng minh chưa chấp nhận thì không chạy; integration riêng thực hiện trên Excel thật.

Suite chính được chạy trước bản gia cố cuối ở bước sao chép đầu vào; ba test tráo file/đọc Excel thật và hai benchmark cuối đã chạy qua đường đọc được gia cố. Mã nguồn cuối đã compile; GUI 1120×768 không có section bị cắt. `.work` sạch sau nghiệm thu.

`acceptance-v0.2.json` lưu checksum mã nguồn, checksum các báo cáo kiểm tra và số đo. `benchmark.py`/`benchmark_ai.py` cho phép đo lại. Giới hạn dữ liệu và các bước nghiệm thu bảo mật toàn máy/PDF còn thiếu được giữ rõ trong README.

Dùng cửa sổ **0.2**. Nếu cửa sổ 0.1 còn mở, đóng sau khi xong tác vụ cũ; tác vụ này giữ lại cửa sổ đó để không mất dữ liệu đang nhập.
