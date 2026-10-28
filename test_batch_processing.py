"""Comprehensive verification for Zero-Token Batch Processing across multiple spreadsheets."""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try: sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

import json
from pathlib import Path
import tempfile
import app
from test_app import fixture

CHECKS = []

def check(name, condition):
    assert condition, name
    CHECKS.append(name)
    print(f"  [PASS] {name}")

def run_batch_tests():
    print("=== KIỂM THỬ XỬ LÝ HÀNG LOẠT (BATCH PROCESSING) - 0 AI TOKEN ===")

    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        work_dir = Path(td)
        batch_src = work_dir / 'classes'
        batch_src.mkdir()
        out_dir = work_dir / 'output'
        out_dir.mkdir()

        # 1. Tạo 3 file lớp học mẫu
        fixture(batch_src / 'lop_ktpm1.xlsx', {'Sheet1': [
            ['MSSV', 'Họ tên', 'Lớp', 'Học phí'],
            ['001', 'Nguyễn An', 'KTPM1', '1.500.000 đ'],
            ['002', 'Trần Bình', 'KTPM1', '2.000.000 đ']
        ]})
        fixture(batch_src / 'lop_ktpm2.xlsx', {'Sheet1': [
            ['MSSV', 'Họ tên', 'Lớp', 'Học phí'],
            ['003', 'Lê Cường', 'KTPM2', '2.500.000 đ'],
            ['004', 'Phạm Dũng', 'KTPM2', '3.000.000 đ']
        ]})
        fixture(batch_src / 'lop_ktpm3.xlsx', {'Sheet1': [
            ['MSSV', 'Họ tên', 'Lớp', 'Học phí'],
            ['005', 'Vũ Em', 'KTPM3', '1.800.000 đ'],
            ['006', 'Hoàng Giang', 'KTPM3', '2.200.000 đ']
        ]})

        # Test 1: Chạy Batch Summary tính tổng học phí theo Lớp trên cả 3 file
        print("\n1. Kiểm tra Batch Summary (Tổng hợp theo nhóm):")
        plan_sum = {
            'action': 'summary',
            'key': 'Lớp',
            'value': 'Học phí',
            'ref_key': '',
            'ref_value': '',
            'lenient_numbers': True
        }
        progress_records = []
        def on_prog(cur, total, item):
            progress_records.append((cur, total, item['file'], item['status']))

        rep1 = app.execute_batch(batch_src, plan=plan_sum, output_dir=out_dir, on_progress=on_prog)
        check("Batch report ghi nhận đúng 3 file", rep1['total_files'] == 3)
        check("Tất cả 3 file đều thành công 100%", rep1['succeeded'] == 3 and rep1['failed'] == 0)
        check("Tiêu tốn đúng 0 AI token", rep1['zero_tokens_used'] is True)
        check("Callback tiến độ được gọi đủ 3 lần", len(progress_records) == 3)
        check("Tệp báo cáo tổng kết JSON đã được lưu", Path(rep1['report_file']).exists())

        # Test 2: Chạy Batch Lookup có tham chiếu danh mục
        print("\n2. Kiểm tra Batch Lookup (Đối chiếu danh mục):")
        ref_file = work_dir / 'danh_muc_khoa.xlsx'
        fixture(ref_file, {'DanhMuc': [
            ['MSSV', 'Khoa chuyên ngành'],
            ['001', 'CNTT'],
            ['002', 'CNTT'],
            ['003', 'KHMT'],
            ['004', 'KHMT'],
            ['005', 'HTTT'],
            ['006', 'HTTT']
        ]})
        plan_lookup = {
            'action': 'lookup',
            'key': 'MSSV',
            'value': '',
            'ref_key': 'MSSV',
            'ref_value': 'Khoa chuyên ngành'
        }
        rep2 = app.execute_batch(batch_src, reference_path=ref_file, plan=plan_lookup, output_dir=out_dir)
        check("Batch Lookup thành công 3/3 file", rep2['succeeded'] == 3 and rep2['failed'] == 0)

        # Test 3: Khả năng chống chịu lỗi (Fault Tolerance) khi có file hỏng hoặc thiếu cột
        print("\n3. Kiểm tra khả năng cách ly lỗi (File hỏng không làm dừng toàn bộ batch):")
        bad_file = batch_src / 'lop_loi_cot.xlsx'
        fixture(bad_file, {'Sheet1': [
            ['ID_Khac', 'Ten_Khac', 'Ghi_Chu'],
            [1, 'Test', 'Thieu cot MSSV']
        ]})
        rep3 = app.execute_batch(batch_src, reference_path=ref_file, plan=plan_lookup, output_dir=out_dir)
        check("Nhận diện 4 file trong thư mục", rep3['total_files'] == 4)
        check("3 file chuẩn thành công, 1 file lỗi được cô lập an toàn", rep3['succeeded'] == 3 and rep3['failed'] == 1)
        bad_res = next(r for r in rep3['results'] if r['file'] == 'lop_loi_cot.xlsx')
        check("File lỗi ghi rõ nguyên nhân thiếu cột", bad_res['status'] == 'failed' and 'thiếu cột' in bad_res['error'].lower())

        # Test 4: Lọc theo pattern tên file
        print("\n4. Kiểm tra lọc file theo pattern:")
        rep4 = app.execute_batch(batch_src, reference_path=ref_file, plan=plan_lookup, output_dir=out_dir, file_pattern='lop_ktpm*.xlsx')
        check("Lọc pattern lop_ktpm*.xlsx bỏ qua file lỗi", rep4['total_files'] == 3 and rep4['succeeded'] == 3)

    print("\n=======================================================")
    print(f"TẤT CẢ {len(CHECKS)} KIỂM TRA BATCH ĐỀU PASS 100%!")
    print("=======================================================")

if __name__ == '__main__':
    run_batch_tests()
