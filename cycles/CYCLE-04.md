# CYCLE-04 — Tách nhánh `lookup` khỏi `prepare_sheets` (M-4, lát 1)

## Pha 1 — Kế hoạch

### Vấn đề người dùng thật
`prepare_sheets()` dài 262 dòng, chứa 4 nhánh hành động (`lookup`, `summary`,
`clean`, `filter`) trong một thân hàm. Mỗi lần sửa cách đối chiếu phải đọc lẫn
logic tổng hợp/làm sạch/lọc; rủi ro sửa nhầm nhánh khác. Đây là M-4 trong ROADMAP
(số đo đã sửa lại: `infer_smart_plan` chỉ còn 25 dòng, không phải 478).

### Mục tiêu
1. Tách **riêng nhánh `if action == 'lookup':`** (~97 dòng) thành hàm cấp module
   `_prepare_lookup(source, reference, plan, rows, keys, formulas, stats, sheets,
   helper, value_result_col...)` — chữ ký cuối cùng chốt khi code, miễn không đổi
   hành vi.
2. `prepare_sheets()` gọi hàm mới và giữ nguyên dữ liệu trả về.
3. Không đổi công thức, nhãn cột, thứ tự cột, `stats`, hay token.

### Không làm trong chu kỳ này
- Không tách nhánh `summary`/`clean`/`filter` (để chu kỳ 5–6).
- Không đổi schema plan, template công thức, hay đối soát native Excel.
- Không đổi giao diện.

### Cách làm (một commit)
- Cắt nguyên khối `if action == 'lookup': ...` (đến trước `elif action == 'summary':`).
- Dedent, thay biến tự do bằng tham số hoặc trả về tuple
  `(headers, formulas, stats, sheets)`; ưu tiên trả về dữ liệu thay vì sửa biến
  ngoài để dễ đọc.
- So AST hằng chuỗi với bản cũ; chỉ được khác docstring/tên hàm.

### Tiêu chí nghiệm thu (đo được)
- AST: hằng chuỗi của khối cũ và hàm mới trùng khớp theo thứ tự.
- `test_app.py --integration` 97/97 (có lookup Excel thật).
- `test_usecases.py` 37/37 (UC-11..UC-16 khóa lookup).
- `test_gui.py` 14/14, `test_ui_acceptance.py` 48/48.
- Mutation: đổi một nhãn cột trong hàm mới (ví dụ `'Gia_tri'` → `'GiaTri'`) phải
  làm ít nhất một test fail; sau đó phục hồi.

### Rủi ro
- Nhánh lookup sửa `rows` tại chỗ và thêm cột; nếu trả về `rows` mới phải gán lại.
- Có hai nhánh con `clean_output` bật/tắt với công thức khác nhau; test phải chạy
  cả hai (đã có trong `test_app`/`test_usecases`).
- Cột `value_result_col` được tính trong nhánh; không được tính lại ở ngoài.

### Lệnh kiểm tra
```powershell
python -X utf8 -m py_compile app.py
python -X utf8 test_app.py --integration
python -X utf8 test_usecases.py
python -X utf8 test_gui.py
python -X utf8 test_ui_acceptance.py
```
