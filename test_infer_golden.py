"""Golden tests for ai_client.infer_smart_plan (rounds 13-14).

Pins the current, verified behaviour of the heuristic plan inference for 20
representative Vietnamese prompts across all branches (Q&A clarify, lookup,
summary, clean, filter, no-match). Created when infer_smart_plan was split
into _qa_clarify_plan/_summary_plan/_filter_plan/_clean_plan; any future
refactor of those helpers must keep these outputs identical.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import ai_client

# (prompt, source_columns, reference_columns, expected (action, key, value, ref_key, ref_value) or None)
CASES = [
    ("Tại sao bị lỗi #N/A vậy?", ["MSSV", "Họ tên"], ["MSSV", "Khoa"], ("clarify", "", "", "", "")),
    ("Bảng này có bao nhiêu cột?", ["MSSV", "Họ tên", "Lớp"], (), ("clarify", "", "", "", "")),
    ("Hướng dẫn sử dụng phần mềm", ["A"], (), ("clarify", "", "", "", "")),
    ("Vẽ biểu đồ điểm cho tôi", ["A"], (), ("clarify", "", "", "", "")),
    ("Hệ thống giúp gì được cho tôi?", ["A"], (), ("clarify", "", "", "", "")),
    ("Xóa dòng trùng trong bảng", ["A"], (), ("clarify", "", "", "", "")),
    ("Đối chiếu Mã sinh viên để lấy Họ tên từ danh mục", ["Mã sinh viên", "Lớp"], ["Mã sinh viên", "Họ tên"],
     ("lookup", "Mã sinh viên", "", "Mã sinh viên", "Họ tên")),
    ("Tra cứu theo MSSV lấy Quê quán", ["MSSV"], ["MSSV", "Quê quán"],
     ("lookup", "MSSV", "", "MSSV", "Quê quán")),
    # VD-15 fixed in round 16: generic words after 'sang'/'vào' ('bảng',
    # 'file') are not destination columns; value must be empty.
    ("Điền tên sinh viên sang bảng này", ["MSSV"], ["MSSV", "Họ tên"],
     ("lookup", "MSSV", "", "MSSV", "Họ tên")),
    ("Ghép dữ liệu vào file này", ["MSSV"], ["MSSV", "Họ tên"],
     ("lookup", "MSSV", "", "MSSV", "Họ tên")),
    # Real destination columns mentioned in the prompt must keep working.
    ("Điền Họ tên vào cột Ghi chú", ["MSSV", "Ghi chú"], ["MSSV", "Họ tên"],
     ("lookup", "MSSV", "Ghi chú", "MSSV", "Họ tên")),
    ("Lấy Quê quán sang cột Địa chỉ", ["MSSV", "Địa chỉ"], ["MSSV", "Quê quán"],
     ("lookup", "MSSV", "Địa chỉ", "MSSV", "Quê quán")),
    # VD-16 fixed in round 17: (a) an explicitly mentioned key that matches no
    # column must not pass through as a bogus key (falls back to ID matching);
    # (b) the natural split phrasing 'áp <col> lên' triggers lookup.
    ("Đối chiếu theo mã nhân viên để lấy Phòng ban", ["Mã NV", "Họ tên"], ["Mã NV", "Phòng ban"],
     ("lookup", "Mã NV", "", "Mã NV", "Phòng ban")),
    ("Đối chiếu theo xyz123 để lấy Phòng ban", ["Mã NV", "Họ tên"], ["Mã NV", "Phòng ban"],
     ("lookup", "Mã NV", "", "Mã NV", "Phòng ban")),
    ("Áp TTVR lên cột VR", ["MSSV", "VR"], ["MSSV", "TTVR"],
     ("lookup", "MSSV", "VR", "MSSV", "TTVR")),
    ("Áp dụng công thức tính điểm", ["MSSV", "Điểm"], (), None),
    ("Tính tổng học phí của từng lớp", ["Lớp", "Học phí"], (), ("summary", "Lớp", "Học phí", "", "")),
    ("Cộng tổng tiền theo khoa", ["Khoa", "Tiền"], (), ("summary", "Khoa", "Tiền", "", "")),
    ("Thống kê số lượng theo nhóm", ["Nhóm", "Số lượng"], (), ("summary", "Nhóm", "Số lượng", "", "")),
    ("Làm sạch cột Họ tên, bỏ khoảng trắng thừa", ["Họ tên"], (), ("clean", "Họ tên", "", "", "")),
    ("Chuẩn hóa unicode cột Ghi chú", ["Ghi chú"], (), ("clean", "Ghi chú", "", "", "")),
    ("Lọc danh sách cán sự, lớp trưởng, tổ trưởng", ["Họ tên", "Chức vụ"], (),
     ("filter", "Chức vụ", "lớp trưởng, tổ trưởng, bí thư, phó bí thư, lớp phó, thủ quỹ, cán sự", "", None)),
    ("Lọc sinh viên có Điểm >= 8.0", ["Họ tên", "Điểm"], (), ("filter", "Điểm", ">= 8.0", "", None)),
    ("Lọc các dòng lớp KTPM1", ["Họ tên", "Lớp"], (), ("filter", "Lớp", "ktpm1", "", None)),
    ("Đăng ký đơn cho sinh viên", ["Họ tên", "Hình thức"], (), ("filter", "Hình thức", "Đơn", "", None)),
    ("Cho tôi danh sách các sinh viên", ["Họ tên", "Lớp"], (), ("filter", "Lớp", "*", "", None)),
    # M2-4 regression (round 15): a filter-intent prompt with NO reference
    # columns must stay a filter; the lookup branch requires reference columns.
    ("Lấy riêng danh sách lớp trưởng", ["Họ tên", "Chức vụ"], (), ("filter", "Chức vụ", "lớp trưởng", "", None)),
    ("Hôm nay thời tiết đẹp", ["A", "B"], (), None),
]

def run():
    for prompt, src, ref, expected in CASES:
        result = ai_client.infer_smart_plan(prompt, list(src), list(ref))
        if expected is None:
            assert result is None, f'{prompt!r}: expected None, got {result}'
            continue
        assert result is not None, f'{prompt!r}: expected {expected}, got None'
        got = tuple(result.get(k) for k in ('action', 'key', 'value', 'ref_key', 'ref_value'))
        assert got == expected, f'{prompt!r}: expected {expected}, got {got}'
    print(f'All {len(CASES)} infer_smart_plan golden cases passed.')

if __name__ == '__main__':
    run()
