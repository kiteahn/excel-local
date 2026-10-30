# NEXT — việc tiếp theo

## Trạng thái (2026-10-30)
- Chu kỳ 0 trên Windows đã chạy trong checkout cô lập nhánh muse-upgrade, commit a9a1c5a.
- Các bài test tự động nhóm 1–3 đều đạt khi chạy Excel COM ngoài sandbox. Lần chạy trong sandbox bị E_ACCESSDENIED; ngoài sandbox chạy thành công.
- Kết quả đầy đủ và output nguyên văn: cycles/CYCLE-00.md.
- Không chuyển sang chu kỳ 1 và không tách gui() cho tới khi Windows acceptance hoàn tất.

## Việc tiếp theo
1. Hoàn tất các xác nhận UI còn lại trước khi kết luận Chu kỳ 0 sạch.
2. Nếu có thay đổi trong lần kiểm tra thủ công, thêm regression test, chạy lại suite liên quan và ghi output; sau mỗi test tích hợp kiểm tra không còn tiến trình Excel do test tạo.
3. Hoàn tất các kiểm tra thủ công bên dưới; không báo đạt khi chưa quan sát.

## Câu hỏi cho chủ sản phẩm — CẦN NGƯỜI XÁC NHẬN

- UC-32: chạy trọn luồng trong app và xem sheet kết quả/Kiem_tra/biên bản. W-1 riêng đã kiểm bằng Excel trên output của integration test: không có Sheet1 trống.
- UC-33: CHƯA CHẠY. Cần thử bằng workbook tự sinh trong phiên test riêng. Bài dừng trước khi tạo file vì đã có Excel người dùng mở; không tương tác với tiến trình đó.
- UC-35: xác nhận file có ô #DIV/0! bị từ chối rõ ràng.
- UC-36: PDF integration test đã được render và xem; bảng không cắt/chồng. Xác nhận thêm PDF xuất từ luồng app nếu cần.
- Tooltip/phím tắt: xác nhận tooltip hai nút và Ctrl+Enter chạy Đề xuất.
- Hộp lỗi Batch: thử thư mục có file bảng tính hỏng và xác nhận hộp báo lỗi tử tế. UI chưa quan sát được do computer-use node_repl lỗi khi khởi tạo.


## Đã kiểm chứng bằng script

- DPAPI thật: test_router.py 7/7 và test_security_privacy.py 43/43 đạt.
- Ollama mặc định và giá trị 500.000 bật/tắt: test_policy_defaults.py P-01..P-05, P-11..P-12 đạt.
- GUI: test_gui.py 14/14; test_app integration 97/97, test_read_boundary 3/3. Benchmark COM đã chạy, số liệu trong cycles/CYCLE-00.md và results/benchmark-current-1000.json.


## Giới hạn thao tác UI hiện tại

- Computer Use không khởi tạo được: 	rusted Node process exited unexpectedly, kể cả sau reset. Không tiếp tục tự động hóa UI khi chưa có runtime hoạt động.
- Hiện có một phiên Excel mở từ bên ngoài bài test; không tương tác hoặc đóng phiên đó.

## Điều kiện dừng và nhánh

- Nếu STOP.txt xuất hiện ở thư mục gốc, dừng ngay và kiểm tra lại sau mỗi bước lớn.
- Không push; việc push nhánh muse-upgrade thuộc về chủ sản phẩm.
