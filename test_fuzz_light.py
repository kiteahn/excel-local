"""Round 12: light fuzzing around load_csv / inspect CSV / header handling.

Contract under test: every input either parses correctly or raises a clear
ValueError. Any other exception type (or a crash) is a failure here.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import tempfile
import unicodedata
from pathlib import Path

import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def outcome(fn):
    """Return ('value', result) or ('valueerror', msg); anything else propagates and fails the run."""
    try:
        return 'value', fn()
    except ValueError as exc:
        return 'valueerror', str(exc)

def run():
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        d = Path(td)
        def w(name, data, binary=False):
            p = d / name
            if binary: p.write_bytes(data)
            else: p.write_text(data, encoding='utf-8', newline='')
            return p

        kind, rows = outcome(lambda: app.load_csv(w('crlf.csv', 'A,B\r\n1,2\r\n')))
        ok('F-01 CRLF doc duoc', kind == 'value' and rows == [['A', 'B'], ['1', '2']])

        kind, rows = outcome(lambda: app.load_csv(w('noeol.csv', 'A,B\n1,2')))
        ok('F-02 thieu newline cuoi van doc duoc', kind == 'value' and rows[-1] == ['1', '2'])

        kind, _ = outcome(lambda: app.load_csv(w('blank.csv', 'A,B\n\n1,2\n')))
        ok('F-03 dong trong xen ke: strict tu choi bang ValueError', kind == 'valueerror')
        kind, rows = outcome(lambda: app.load_csv(w('blank2.csv', 'A,B\n\n1,2\n'), strict_ragged=False))
        ok('F-04 dong trong xen ke: lenient bo qua dong trong', kind == 'value' and rows == [['A', 'B'], ['1', '2']])

        kind, rows = outcome(lambda: app.load_csv(w('maxcell.csv', 'A\n' + 'x' * 32767 + '\n')))
        ok('F-05 o dung 32767 ky tu duoc chap nhan', kind == 'value' and len(rows[1][0]) == 32767)
        kind, _ = outcome(lambda: app.load_csv(w('bigcell.csv', 'A\n' + 'x' * 32768 + '\n')))
        ok('F-06 o 32768 ky tu bi tu choi', kind == 'valueerror')

        for pos, frag in (('dau', '\x01abc'), ('giua', 'ab\x01c'), ('cuoi', 'abc\x01')):
            kind, _ = outcome(lambda frag=frag: app.load_csv(w(f'ctl_{pos}.csv', 'A\n' + frag + '\n')))
            ok(f'F-07 ky tu dieu khien o vi tri {pos} bi tu choi', kind == 'valueerror')

        kind, res = outcome(lambda: app.load_csv(w('mixed.csv', 'A,B;C\n1,2;3\n')))
        ok('F-08 phan cach hon hop: doc duoc hoac ValueError, khong sap', kind in ('value', 'valueerror'))

        kind, _ = outcome(lambda: app.load_csv(w('empty.csv', ''), ))
        ok('F-09 file rong bi tu choi bang ValueError', kind == 'valueerror')

        kind, rows = outcome(lambda: app.load_csv(w('bom.csv', b'\xef\xbb\xbfA;B\n1;2\n', binary=True)))
        ok('F-10 UTF-8 BOM + cham phay', kind == 'value' and rows == [['A', 'B'], ['1', '2']])

        # Legacy cp1258: code tries cp1258/cp1252/latin1 fallbacks (docs say it
        # does not guess legacy encodings -> mismatch logged as VD-13).
        kind, rows = outcome(lambda: app.load_csv(w('cp1258.csv', 'Tên\nĐoàn\n'.encode('cp1258'), binary=True)))
        ok('F-11 CSV cp1258 doc duoc theo hanh vi code hien tai', kind == 'value' and rows[1][0] == 'Đoàn')

        kind, res = outcome(lambda: app.load_csv(w('quote.csv', 'A\nab"cd\n')))
        ok('F-12 nhay kep giua o khong ngoac: doc duoc hoac ValueError', kind in ('value', 'valueerror'))

        print('--- header fuzz ---')
        def raw(rows):
            return {'sheet': 'Data', 'sheets': ['Data'], 'rows': rows,
                    'metadata': {'sheets': {'Data': {'formats': {}}}}}
        kind, _ = outcome(lambda: app.select_header(raw([['  ', ' '], ['1', '2']]), 1))
        ok('F-13 header toan khoang trang bi tu choi', kind == 'valueerror')
        kind, _ = outcome(lambda: app.select_header(raw([['X' * 1000, 'B'], ['1', '2']]), 1))
        ok('F-14 ten cot 1000 ky tu bi tu choi (gioi han 200)', kind == 'valueerror')
        kind, res = outcome(lambda: app.select_header(raw([['X' * 200, 'B'], ['1', '2']]), 1))
        ok('F-14b ten cot dung 200 ky tu duoc chap nhan', kind == 'value' and res['headers'][0] == 'X' * 200)
        nfc = unicodedata.normalize('NFC', 'Đoàn')
        nfd = unicodedata.normalize('NFD', 'Đoàn')
        kind, res = outcome(lambda: app.select_header(raw([[nfc, nfd], ['1', '2']]), 1))
        # VD-14 fixed in round 12: duplicate detection compares NFC-normalised
        # names, so a visually identical NFC/NFD pair is rejected as duplicate.
        ok('F-15 cap header NFC/NFD giong het nhau bi tu choi la trung ten (VD-14)', kind == 'valueerror')
        kind, res = outcome(lambda: app.detect_header_row([['A', 'B'], ['A', 'B'], ['A', 'B']]))
        ok('F-16 detect_header_row voi moi dong giong nhau khong sap', kind == 'value' and isinstance(res, int))
        kind, res = outcome(lambda: app.detect_header_row([]))
        ok('F-17 detect_header_row voi danh sach rong: tra ve int hoac ValueError', kind in ('value', 'valueerror'))

    print(f'\nAll {len(CHECKS)} fuzz checks passed (no unexpected exception types).')

if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    run()
