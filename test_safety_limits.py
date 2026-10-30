"""Safety-limit tests (round 24). Mutation round 4 showed four guards with
no test coverage; these pin them:
- preflight 80 MB file-size limit (M4-2)
- preflight 10000-entry ZIP limit (M4-4)
- validate_plan 2000-char string limit (M4-3)
- inspect_file source-hash recheck after reading (M4-1)
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import tempfile
import zipfile
from pathlib import Path
from unittest.mock import patch

import app
from test_app import fixture

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

def run():
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        d = Path(td)

        big = d / 'big.csv'
        with open(big, 'wb') as fh:
            fh.write(b'A,B\n')
            fh.truncate(81 * 1024 * 1024)
        rejects('L-01 preflight tu choi file 81MB', lambda: app.preflight(big))

        exact = d / 'exact.csv'
        with open(exact, 'wb') as fh:
            fh.write(b'A,B\n')
            fh.truncate(80 * 1024 * 1024)
        app.preflight(exact)
        ok('L-02 preflight chap nhan file dung 80MB', True)

        many = d / 'many.xlsx'
        fixture(many, {'Data': [['ID'], ['001']]})
        with zipfile.ZipFile(many, 'a') as z:
            for i in range(10001):
                z.writestr(f'junk/{i}.bin', b'')
        rejects('L-03 preflight tu choi XLSX hon 10000 thanh phan', lambda: app.preflight(many))

        snap = {'sheet': 'Data', 'headers': ['A'], 'data': [['1']]}
        plan = {'action': 'clarify', 'message': 'x' * 2001}
        rejects('L-04 validate_plan tu choi chuoi dai hon 2000 ky tu',
                lambda: app.validate_plan(plan, snap, None, allow_clarify=True))
        plan_ok = {'action': 'clarify', 'message': 'x' * 2000}
        app.validate_plan(plan_ok, snap, None, allow_clarify=True)
        ok('L-05 validate_plan chap nhan chuoi dung 2000 ky tu', True)

        csv_path = d / 'swap.csv'
        csv_path.write_text('ID,Note\n001,x\n', encoding='utf-8')
        original_load = app.load_csv
        def mutating_load(path, strict_ragged=True):
            rows = original_load(path, strict_ragged=strict_ragged)
            with open(path, 'a', encoding='utf-8') as fh:
                fh.write('002,y\n')
            return rows
        with patch.object(app, 'load_csv', mutating_load):
            rejects('L-06 inspect_file phat hien file thay doi giua luc doc (hash)',
                    lambda: app.inspect_file(csv_path))

    print(f'\nAll {len(CHECKS)} safety-limit checks passed.')

if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    run()
