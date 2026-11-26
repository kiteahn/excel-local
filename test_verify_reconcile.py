"""Verify/reconcile operation (action='verify'), added 2026-11-10.

Read-only check: every source row is matched to the reference file by key
and the chosen columns are compared cell by cell. Covers the pure core
(run_verify), plan validation, the preview/token flow and a full CSV
end-to-end run (CSV in/out needs no Excel, so it runs on Linux too).
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import csv
import tempfile
from pathlib import Path

import app
import ai_client

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def snapshot(headers, rows, sheet='S1'):
    return {'headers': list(headers), 'data': [list(r) for r in rows],
            'metadata': {'sheets': {sheet: {'formats': {}}}}, 'sheet': sheet,
            'header_row': 1, 'column_offset': 0, 'hash': 'h-' + sheet, 'path': 'x.xlsx'}

def base_plan(**over):
    plan = {'action': 'verify', 'key': 'Mã SV', 'value': '', 'ref_key': 'Mã SV',
            'ref_value': '', 'message': '', 'selected_columns': 'Họ tên, Điểm'}
    plan.update(over)
    return plan

def run():
    src = snapshot(['Mã SV', 'Họ tên', 'Điểm'], [
        ['SV01', 'Nguyễn Văn An', 8],
        ['SV02', 'Trần Thị Bình', '500.000'],
        ['SV03', 'Lê Văn Cường', 7],
        ['SV04', 'Phạm Thị Dung', 9],
        ['', 'Không Khóa', 5],
        ['SV06', 'Đỗ Văn Em', 6],
    ], sheet='Goc')
    ref = snapshot(['Mã SV', 'Họ tên', 'Điểm'], [
        ['SV01', 'nguyễn văn an', 8.0],
        ['SV02', 'Trần Thị Bình', 500000],
        ['SV03', 'Lê Văn Cường', 7.5],
        ['SV06', 'Đỗ Văn Em', 6],
        ['SV06', 'Đỗ Văn Em (bản 2)', 6],
        ['SV99', 'Chỉ Có Ở Đối Chiếu', 1],
    ], sheet='DoiChieu')

    res = app.run_verify(src, ref, base_plan())
    s = res['summary']
    ok('V-01 dong khop: SV01 (hoa/thuong + 8 vs 8.0) va SV02 (500.000 vs 500000)',
       s['matched_rows'] == 2)
    ok('V-02 dong lech: SV03 diem 7 vs 7.5', s['mismatch_rows'] == 1 and s['cells_mismatch'] == 1)
    ok('V-03 thieu trong doi chieu: SV04', s['missing_rows'] == 1)
    ok('V-04 khoa trung phia doi chieu: SV06', s['duplicate_ref_rows'] == 1 and s['reference_duplicate_keys'] == 1)
    ok('V-05 khoa trong o nguon', s['blank_key_rows'] == 1)
    ok('V-06 khoa chi co o doi chieu: SV99', s['extra_ref_rows'] == 1 and res['extra_ref_keys'] == ['sv99'])
    statuses = {(i['key'], i['status']) for i in res['issues']}
    ok('V-07 issues co du cac loai trang thai',
       ('SV03', 'Lệch') in statuses and ('SV04', 'Không có trong file đối chiếu') in statuses
       and ('SV06', 'Khóa trùng trong file đối chiếu') in statuses)
    lech = next(i for i in res['issues'] if i['status'] == 'Lệch')
    ok('V-08 issue lech ghi ro hai gia tri va so dong Excel',
       lech['column'] == 'Điểm' and lech['source_value'] == '7' and lech['reference_value'] == '7.5' and lech['row'] == 4)

    # validate_plan: happy path + hard errors
    plan = app.validate_plan(base_plan(), src, ref)
    ok('V-09 validate giu nguyen cot so', plan['selected_columns'] == 'Họ tên, Điểm')
    try:
        app.validate_plan(base_plan(selected_columns='Họ tên, Cột Không Có'), src, ref)
        ok('V-10 cot khong ton tai bi chan', False)
    except ValueError:
        ok('V-10 cot khong ton tai bi chan', True)
    try:
        app.validate_plan(base_plan(), src, None)
        ok('V-11 thieu file doi chieu bi chan', False)
    except ValueError:
        ok('V-11 thieu file doi chieu bi chan', True)
    clar = app.validate_plan(base_plan(selected_columns=''), src, ref, allow_clarify=True)
    ok('V-12 chua chon cot -> clarify', clar['action'] == 'clarify')

    # Full flow on real CSV files: inspect -> preview -> execute -> CSV report
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        src_csv = td / 'goc.csv'
        ref_csv = td / 'doichieu.csv'
        with open(src_csv, 'w', encoding='utf-8', newline='') as f:
            w = csv.writer(f); w.writerow(['Mã SV', 'Họ tên', 'Điểm'])
            w.writerows([['SV01', 'Nguyễn Văn An', 8], ['SV02', 'Trần Bình', 6]])
        with open(ref_csv, 'w', encoding='utf-8', newline='') as f:
            w = csv.writer(f); w.writerow(['Mã SV', 'Họ tên', 'Điểm'])
            w.writerows([['SV01', 'Nguyễn Văn An', 8], ['SV02', 'Trần Bình', 6.5]])
        s_snap = app.inspect_file(src_csv)
        r_snap = app.inspect_file(ref_csv)
        eplan = app.validate_plan(base_plan(), s_snap, r_snap)
        preview = app.preview_plan(s_snap, r_snap, eplan, output_dir=td / 'out')
        ok('V-13 preview verify: 1 dong khop, 1 dong lech, khong cong thuc',
           preview['stats']['matched_rows'] == 1 and preview['stats']['mismatch_rows'] == 1
           and preview['formula_checks'] == 0 and len(preview['sample_rows']) == 1)
        report = app.execute(s_snap, r_snap, eplan, approval_token=preview['approval_token'], output_dir=td / 'out')
        out_csv = Path(report['output'])
        ok('V-14 execute xuat bao cao CSV + giu nguyen file goc',
           out_csv.exists() and out_csv.suffix == '.csv' and report['source_unchanged'] is True)
        text = out_csv.read_text(encoding='utf-8-sig')
        ok('V-15 bao cao ghi dong lech SV02', 'SV02' in text and 'Lệch' in text)
        try:
            app.execute(s_snap, r_snap, eplan, approval_token='sai-token', output_dir=td / 'out')
            ok('V-16 token sai bi chan', False)
        except ValueError:
            ok('V-16 token sai bi chan', True)

    # AI layer: heuristic inference + sanitize
    inferred = ai_client.infer_smart_plan(
        'Kiểm tra dữ liệu file gốc có đúng với file đối chiếu theo Mã SV không',
        ['Mã SV', 'Họ tên', 'Điểm'], ['Mã SV', 'Họ tên', 'Điểm'])
    ok('V-17 infer nhan dien verify va chon khoa + cot so',
       inferred is not None and inferred['action'] == 'verify' and inferred['key'] == 'Mã SV'
       and 'Họ tên' in inferred['selected_columns'] and 'Điểm' in inferred['selected_columns'])
    no_ref = ai_client.infer_smart_plan(
        'Kiểm chứng dữ liệu có khớp không', ['Mã SV', 'Họ tên'], [])
    ok('V-18 infer verify khong co file doi chieu -> clarify',
       no_ref is not None and no_ref['action'] == 'clarify')
    clean = ai_client.sanitize_plan(
        {'action': 'verify', 'key': 'Mã SV', 'ref_key': 'Mã SV', 'selected_columns': 'Họ tên, Điểm'},
        ['Mã SV', 'Họ tên', 'Điểm'], ['Mã SV', 'Họ tên', 'Điểm'])
    ok('V-19 sanitize giu verify + selected_columns',
       clean['action'] == 'verify' and clean['selected_columns'] == 'Họ tên, Điểm')
    lookup_still = ai_client.infer_smart_plan(
        'Đối chiếu theo Mã SV để lấy Họ tên từ danh mục',
        ['Mã SV', 'Họ tên'], ['Mã SV', 'Họ tên'])
    ok('V-20 prompt doi chieu lay gia tri van la lookup, khong bi verify nuot',
       lookup_still is not None and lookup_still['action'] == 'lookup')

    print(f'\nAll {len(CHECKS)} verify checks passed.')

if __name__ == '__main__':
    run()
