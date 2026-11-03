# Chu kỳ 7 — Q4: dữ liệu giống thật, 55 câu và luồng Excel đầy đủ

Ngày: 2026-11-03. Nhánh `muse-upgrade`; không push; không dùng file thật.

## Dữ liệu tự sinh

`create_real_like_examples.py` tạo 9 workbook trong `examples/real_like/` với dữ
liệu hoàn toàn bịa, chạy lại được:

- `sinh_vien.xlsx`: 120 sinh viên, banner 2 dòng, mã 0 đầu, ngày sinh trộn 3
  định dạng, tên NFC/NFD.
- `bang_luong.xlsx`: 80 nhân viên, lương dạng chuỗi có dấu chấm nghìn.
- `kho_hang.xlsx`: 100 dòng + dòng `TỔNG CỘNG`, đơn giá/thành tiền dạng chuỗi.
- `diem_danh.xlsx`: 100 lượt điểm danh, trạng thái Có mặt/Vắng, ghi chú có
  khoảng trắng thừa.
- `so_quy.xlsx`: 60 phiếu thu/chi, một ô `#N/A`, một dòng tổng.
- `don_hang.xlsx`: 120 dòng, mã đơn lặp theo từng dòng hàng.
- `danh_muc_hang.xlsx`, `danh_muc_nhan_vien.xlsx`, `danh_muc_lop.xlsx`: danh mục
  tra cứu, trong đó lớp chỉ có 3 khóa duy nhất.

## Đo tỷ lệ kế hoạch đúng ngay lần đầu

`test_real_like_plans.py` đọc cả 9 file bằng `inspect_file` (chỉ `so_quy` đọc
`lenient`), gọi `infer_smart_plan` cho **55 câu tiếng Việt**:

- **50 câu curated** trải trên 6 loại bảng nguồn và 4 thao tác
  lookup/summary/clean/filter.
- **5 câu tự nhiên** phát hiện trong lúc thử; không được chọn theo code.

Kết quả lưu tại `results/plan-accuracy-q4.json`:

| Bước | Kế hoạch đúng | Tỷ lệ |
|---|---:|---:|
| Trước khi sửa Q4 | 46/55 | 83,6% |
| Sau khi sửa 3 lỗi hiểu câu | 53/55 | **96,4%** |

Hai câu còn sai là giới hạn có chủ đích: yêu cầu **trung bình** (`AVG`) chưa
được hỗ trợ; planner trả `None` thay vì `clarify` nói rõ chỉ hỗ trợ tổng. Đã
ghim cả 5 câu vào `test_infer_golden.py` (33 prompt vàng) kèm ghi chú cách sửa.

## Luồng đầy đủ

Với 50 kế hoạch curated đúng, test chạy tuần tự:

`preview_plan` → lấy `approval_token` → `execute` → Excel native ghi file →
`verify_saved_result` → `preserve_check`.

- **49/50 chạy thành công**, 0 lỗi.
- **1/50 bị từ chối có kiểm soát**: `so_quy` có ô `#N/A` trong cột `Số dư`;
  app dừng với thông báo chuyển đổi `#N/A`, không âm thầm bỏ qua. Test coi đây
  là hành vi đúng và ghi rõ `refused_as_expected`.
- Mọi workbook đều đi qua ít nhất một thao tác; 9/9 file đọc được.

## Lỗi thật phát hiện và sửa

1. **`app.xml_value` crash với ô rỗng `t="inlineStr"`**: openpyxl ghi ô trống
   thành `<c t="inlineStr"/>` không có `<is>`, làm `preserve_check` ném
   `AttributeError: 'NoneType' object has no attribute 'findall'`. Đã trả
   `('empty', None)` khi thiếu `<is>`. Đây là lỗi thật trên `kho_hang.xlsx` và
   `so_quy.xlsx`, không phải lỗi của test.
2. **Tổng hợp gặp dòng `TỔNG CỘNG`**: dòng tổng có khóa nhóm trống làm app từ
   chối cả bảng. Nay app bỏ qua dòng tổng đã in, ghi số dòng bỏ qua vào
   `stats['total_rows_skipped']`; dòng khóa trống không có nhãn tổng vẫn bị từ
   chối như cũ.
3. **Hiểu câu lọc/ làm sạch/ đối chiếu**:
   - chọn cột khớp dài nhất thay vì cột đầu tiên cho clean/lookup;
   - nhận diện giá trị categorical khi người dùng không nói tên cột
     (`Kinh doanh` → `Phòng ban`, `vắng` → `Trạng thái`, `Khách A` →
     `Khách hàng`, `lớp CNTT01` → `Lớp`);
   - `khác rỗng` được hiểu là `*` (có giá trị), không phải chuỗi literal;
   - `Thống kê ... Có mặt` chọn được cột `Có mặt` cho tổng hợp.

## Kiểm thử

- `test_infer_golden.py`: **33/33** prompt vàng.
- `test_real_like_plans.py`: **50/50** kế hoạch curated đúng; **49** chạy
  end-to-end, **1** từ chối đúng, **0** lỗi.
- `test_usecases.py` 37/37, `test_usecases2.py` 10/10,
  `test_prepare_sheets_golden.py` 27/27, `test_headers.py` 11/11,
  `test_diverse_scenarios_e2e.py` 6/6.

## Giới hạn / chưa kiểm chứng

- 9 workbook là dữ liệu tổng hợp; **CHƯA KIỂM CHỨNG** trên file thật của người
  dùng (quy định không dùng file thật).
- Chưa hỗ trợ trung bình/AVG; 2 câu L04/L05 ghi nhận là giới hạn.
- Tổng hợp bỏ qua dòng tổng theo nhãn văn bản; bảng có dòng tổng không ghi nhãn
  vẫn bị từ chối nếu khóa nhóm trống.
- Ô lỗi `#N/A` trong cột cần tính vẫn bị từ chối; đây là lựa chọn an toàn, chưa
  phải chức năng sửa lỗi dữ liệu.
