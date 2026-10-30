"""Round 22: fuzz round 2 - detect_header_row / select_header on dict snapshots.

Contract: every input either resolves to a sane header or raises
HeaderValidationError/ValueError. Any other exception type is a failure.
Behaviours that are surprising but by-design are pinned and labelled.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def outcome(fn):
    try:
        return 'value', fn()
    except (ValueError, app.HeaderValidationError) as exc:
        return 'valueerror', str(exc)

def raw(rows, formats=None):
    return {'sheet': 'Data', 'sheets': ['Data'], 'rows': rows,
            'metadata': {'sheets': {'Data': {'formats': formats or {}}}}}

def run():
    kind, _ = outcome(lambda: app.select_header(raw([[None, 'B'], ['1', '2']]), 1))
    ok('H-01 o header None bi tu choi nhu o trong', kind == 'valueerror')

    kind, res = outcome(lambda: app.detect_header_row([[1, 2], ['MSSV', 'Ho ten'], ['001', 'An']]))
    ok('H-02 dong dau toan so: detect nhay toi dong chu', kind == 'value' and res == 2)

    kind, res = outcome(lambda: app.detect_header_row([[1, 2], [3, 4]]))
    ok('H-03 moi dong deu la so: khong sap, tra ve int', kind == 'value' and isinstance(res, int))

    kind, res = outcome(lambda: app.detect_header_row([['MSSV', 'Ho ten']]))
    ok('H-04 chi mot dong duy nhat: khong sap', kind == 'value' and isinstance(res, int))

    kind, res = outcome(lambda: app.select_header(raw([['A', 'B', 'C'], ['1', '2'], ['3', '4', '5', '6']]), 1))
    ok('H-05 dong du lieu lech do dai: doc duoc hoac ValueError ro rang', kind in ('value', 'valueerror'))

    kind, res = outcome(lambda: app.select_header(raw([['\u00a0MSSV\u00a0', 'B'], ['1', '2']]), 1))
    ok('H-06 NBSP quanh ten cot duoc cat bo', kind == 'value' and res['headers'][0] == 'MSSV')

    kind, res = outcome(lambda: app.select_header(raw([['ID', 'id'], ['1', '2']]), 1))
    # Pinned current behaviour: duplicate check is case-SENSITIVE for headers
    # (only NFC + exact string), so 'ID'/'id' are two distinct columns.
    ok('H-07 cap ID/id khac hoa thuong: hanh vi hien tai la chap nhan 2 cot rieng', kind == 'value')

    kind, res = outcome(lambda: app.select_header(raw([['MS\u200bSV', 'B'], ['1', '2']]), 1))
    ok('H-08 ky tu zero-width trong ten cot: doc duoc hoac tu choi, khong sap', kind in ('value', 'valueerror'))

    kind, res = outcome(lambda: app.select_header(raw([[f'C{i}' for i in range(61)], [str(i) for i in range(61)]]), 1))
    ok('H-09 snapshot 61 cot qua select_header: chap nhan (gioi han 60 nam o inspect_file, khong o select_header)', kind == 'value')

    kind, res = outcome(lambda: app.select_header(raw([['A', 'A'], ['1', '2']]), 1, auto_heal=True))
    ok('H-10 auto_heal voi cot trung: ten thanh duy nhat', kind == 'value' and len(set(res['headers'])) == 2)

    kind, res = outcome(lambda: app.detect_header_row([
        ['TRUONG DAI HOC TON DUC THANG', None],
        ['DANH SACH SINH VIEN', None],
        ['STT', 'MSSV'],
        [1, '001'],
    ]))
    ok('H-11 banner hanh chinh + dong STT/MSSV: detect dung dong 3', kind == 'value' and res == 3)

    kind, _ = outcome(lambda: app.select_header(raw([['', ''], ['1', '2']]), 1))
    ok('H-12 header rong hoan toan bi tu choi', kind == 'valueerror')

    kind, res = outcome(lambda: app.select_header(raw([['A', '', 'B'], ['1', '2', '3']]), 1, auto_heal=True))
    ok('H-13 auto_heal dien o trong o giua bang ten _phu', kind == 'value' and res['headers'] == ['A', 'A_phụ', 'B'])

    kind, _ = outcome(lambda: app.select_header(raw([['', 'B', ''], ['1', '2', '3']]), 1, auto_heal=True))
    ok('H-13b auto_heal khong cuu dong chi co 1 ten cot (populated_count <= 1): tu choi nhu che do chat',
       kind == 'valueerror')

    kind, res = outcome(lambda: app.detect_header_row([[None, None], ['A', 'B'], ['1', '2']]))
    ok('H-14 dong dau toan None: detect bo qua toi dong chu', kind == 'value' and res == 2)

    # M5-4 regression (round 25): the +25 "has data below" bonus must be what
    # makes row 1 win here; without it row 3 (a footer-like row) would win.
    kind, res = outcome(lambda: app.detect_header_row([['MSSV', 'Họ tên'], ['001', 'An'], ['Lớp', 'Khoa', 'X1']]))
    ok('H-15 diem thuong du lieu ben duoi quyet dinh dong header that (M5-4)', kind == 'value' and res == 1)

    print(f'\nAll {len(CHECKS)} header-fuzz checks passed (no unexpected exception types).')

if __name__ == '__main__':
    run()
