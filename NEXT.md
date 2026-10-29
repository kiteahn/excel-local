# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 16 — Phân tích và xử lý VĐ-15 (lookup value='bảng')

Hiện tượng: "Điền tên sinh viên sang bảng này" → lookup value='bảng' vì
regex cột đích bắt từ đứng sau "sang". value trong lookup đáng lẽ là cột
đích ở bảng nguồn hoặc để trống.

Làm:
1. Đọc đoạn trích cột đích (dest/target) trong _lookup_plan: liệt kê các
   mẫu câu nó phục vụ ("điền X vào cột Y", "lấy X sang Y", "áp X lên Y"...).
2. Thêm prompt vàng cho các mẫu ĐANG ĐÚNG trước khi sửa (đặc biệt các mẫu
   có cột đích thật: "điền Họ tên vào cột Ghi chú", "lấy Quê quán sang cột
   Địa chỉ" — kiểm tra hành vi thật hiện tại trước, ghim đúng kết quả hợp lý).
3. Sửa tối thiểu: nếu từ bắt được là từ chung chỉ bảng/vị trí (bảng, sheet,
   file, đây, đó...) và không khớp tên cột nào trong source_columns sau
   fuzzy match thì bỏ qua (value=''). Kiểm chứng: toàn bộ prompt vàng cũ
   chỉ thay đổi đúng case VĐ-15, còn lại giữ nguyên.
4. Nếu phân tích cho thấy sửa rủi ro hơn lợi: giữ nguyên, ghi lý do vào log
   và đóng VĐ-15 dạng "chấp nhận hiện trạng".

## Hàng đợi sau đó
- Vòng 17+: tách prepare_sheets (~261 dòng) theo mẫu test vàng; gui sau cùng.
- Tổng kết vòng 20 theo PROMPT (mỗi 10 vòng).
- Quyết định chờ người dùng: '500.000' (UC-20b); cloud/local mặc định.
- Hàng đợi Windows: UC-32/33/35/36; benchmark lại; test_gui + trực quan
  Vòng 10; nghiệm thu DPAPI thật.

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Nhánh `muse-upgrade`: 16 commit (Bước 0 + Vòng 1–15), CHƯA push — cần
  người dùng xác nhận mới push lên GitHub (repo công khai).
