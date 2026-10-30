"""Round 28: unusual use cases, round 2 (UC2-xx).

Boundary/odd inputs through validate_plan + prepare_sheets + inspect_file.
Each case either produces the documented result or a clear rejection.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import tempfile
from pathlib import Path

import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def rejects(name, fn):
    try:
        fn()
    except (ValueError, RuntimeError) as exc:
        ok(name, True)
        return
    raise AssertionError(f'{name}: expected rejection, got success')

def snap(headers, rows, sheet='Data'):
    return {'headers': headers, 'data': rows,
            'metadata': {'sheets': {sheet: {'formats': {}}}},
            'sheet': sheet, 'header_row': 1, 'hash': 'test'}

def run():
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        d = Path(td)

        p = d / 'header_only.csv'
        p.write_text('ID,Note\n', encoding='utf-8')
        src = app.inspect_file(p)
        rejects('UC2-01 CSV chi co dong header: prepare tu choi vi khong co dong du lieu',
                lambda: app.prepare_sheets(src, None, {'action': 'clean', 'key': 'Note'}, 't'))

        src = snap(['MSSV', 'Ho ten'], [['001', 'An'], ['002', 'Binh']])
        sheets, name, st = app.prepare_sheets(src, src, {'action': 'lookup', 'key': 'MSSV', 'ref_key': 'MSSV', 'ref_value': 'Ho ten'}, 't')
        ok('UC2-02 tu doi chieu voi chinh no: moi khoa deu khop, missing=0',
           st['matched'] == 2 and st['missing'] == 0)

        src = snap(['Ho ten', 'Chuc vu'], [['An', 'Lop truong'], ['Binh', 'Sinh vien']])
        rejects('UC2-03 filter selected_columns chua cot khong ton tai bi validate chan',
                lambda: app.validate_plan({'action': 'filter', 'key': 'Chuc vu', 'value': 'Lop truong',
                                            'selected_columns': 'Ho ten, Cot Khong Co'}, src, None))

        src = snap(['MSSV'], [['001'], ['002'], ['003']])
        ref = snap(['MSSV', 'Ten'], [['001', 'An'], ['001', 'An Trung'], ['002', 'Binh']])
        _, _, st = app.prepare_sheets(src, ref, {'action': 'lookup', 'key': 'MSSV', 'ref_key': 'MSSV', 'ref_value': 'Ten'}, 't')
        ok('UC2-04 lookup vua trung khoa tham chieu vua thieu khoa: duplicate=1, missing=1',
           st['duplicate'] == 1 and st['missing'] == 1 and st['reference_duplicate_keys'] == 1)

        src = snap(['ID', 'So'], [['A', 5], ['B', 10]])
        sheets, name, st = app.prepare_sheets(src, None, {'action': 'clean', 'key': 'So'}, 't')
        result = next(s for s in sheets if s['name'] == name)
        ok('UC2-05 clean tren cot so: changed=0 va gia tri giu nguyen',
           st['changed_cells'] == 0 and result['rows'][1][1] == 5 and result['rows'][2][1] == 10)

        src = snap(['G', 'V'], [['', 1], [None, 2]])
        rejects('UC2-06 summary moi khoa nhom deu trong: bi tu choi',
                lambda: app.prepare_sheets(src, None, {'action': 'summary', 'key': 'G', 'value': 'V'}, 't'))

        # UC2-07 corrected by evidence: a numeric filter condition on text
        # cells treats them as non-matching (like Excel autofilter), it does
        # not error. filtered_rows in stats makes the outcome visible.
        src = snap(['Ten', 'Diem'], [['An', 'tam'], ['Binh', 'chin']])
        _, _, st = app.prepare_sheets(src, None, {'action': 'filter', 'key': 'Diem', 'value': '>= 8'}, 't')
        ok('UC2-07 filter >= 8 tren cot toan chu: 0 dong khop, khong loi', st['filtered_rows'] == 0)
        src = snap(['Ten', 'Diem'], [['An', '7'], ['Binh', '9'], ['Chi', 'muoi']])
        _, _, st = app.prepare_sheets(src, None, {'action': 'filter', 'key': 'Diem', 'value': '>= 8'}, 't')
        ok('UC2-07b filter >= 8 tren cot hon hop: chi dong so thoa man duoc giu', st['filtered_rows'] == 1)

        weird = 'Sheet ' + 'X' * 100 + ' [đặc biệt]'
        s = snap(['A'], [['1']], sheet=weird)
        sheets, name, st = app.prepare_sheets(s, None, {'action': 'clean', 'key': 'A'}, 't')
        audit = next(x for x in sheets if x['name'].startswith('Kiem_tra'))
        ok('UC2-08 ten sheet rat dai + ky tu la: van xu ly va ghi dung vao sheet kiem tra',
           any(row[0] == 'Sheet_nguon' and row[1] == weird for row in audit['rows']))

        src = snap(['MSSV', 'Ho ten'], [['001', 'An']])
        rejects('UC2-09 lookup thieu ref_value trong ke hoach bi validate chan',
                lambda: app.validate_plan({'action': 'lookup', 'key': 'MSSV', 'ref_key': 'MSSV', 'ref_value': ''}, src, src))

    print(f'\nAll {len(CHECKS)} unusual-use-case checks passed.')

if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    run()
