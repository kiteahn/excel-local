"""Round 29: execute_batch summary accounting (M6-5 regression).

test_batch_processing cannot see a successful execution on Linux (the real
execute needs Excel), so a broken succeeded-counter survived mutation.
Here app.execute is faked while inspect/preview run for real on CSVs,
which exercises the batch loop, per-file error isolation and the report.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import json
import tempfile
from pathlib import Path
from unittest.mock import patch

import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def rejects(name, fn):
    try:
        fn()
    except (ValueError, RuntimeError, NotADirectoryError) as exc:
        ok(name, True)
        return
    raise AssertionError(f'{name}: expected rejection, got success')

def run():
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        d = Path(td)
        src_dir = d / 'in'
        src_dir.mkdir()
        (src_dir / 'a.csv').write_text('ID,Note\n001, x \n', encoding='utf-8')
        (src_dir / 'b.csv').write_text('ID,Note\n002, y \n', encoding='utf-8')
        (src_dir / 'c.csv').write_text('ID,Khac\n003, z \n', encoding='utf-8')
        (src_dir / ' bo qua.txt').write_text('x', encoding='utf-8')
        out_dir = d / 'out'

        def fake_execute(source, reference, plan, export_pdf=False, output_dir=None, *, approval_token, _usage_batch=False):
            return {'output': str(Path(output_dir) / 'fake.xlsx'), 'pdf': None, 'stats': {'rows': 1}}

        with patch.object(app, 'execute', fake_execute):
            summary = app.execute_batch(src_dir, plan={'action': 'clean', 'key': 'Note'}, output_dir=out_dir)
        ok('B-01 batch dem dung: 3 file, 2 thanh cong, 1 loi (thieu cot)',
           summary['total_files'] == 3 and summary['succeeded'] == 2 and summary['failed'] == 1)
        statuses = {r['file']: r['status'] for r in summary['results']}
        ok('B-02 trang thai tung file dung, loi duoc ghi ly do',
           statuses == {'a.csv': 'success', 'b.csv': 'success', 'c.csv': 'failed'}
           and 'Note' in (next(r for r in summary['results'] if r['file'] == 'c.csv')['error'] or ''))
        report = Path(summary['report_file'])
        saved = json.loads(report.read_text(encoding='utf-8'))
        ok('B-03 file bao cao batch ton tai va khop so dem trong bo nho',
           report.exists() and saved['succeeded'] == 2 and saved['failed'] == 1 and saved['zero_tokens_used'] is True)

        rejects('B-04 plan action clarify bi tu choi o batch',
                lambda: app.execute_batch(src_dir, plan={'action': 'clarify'}, output_dir=out_dir))
        empty = d / 'empty'
        empty.mkdir()
        rejects('B-05 thu muc khong co file bang tinh bi tu choi',
                lambda: app.execute_batch(empty, plan={'action': 'clean', 'key': 'Note'}, output_dir=out_dir))
        rejects('B-06 lookup thieu file tham chieu bi tu choi',
                lambda: app.execute_batch(src_dir, plan={'action': 'lookup', 'key': 'ID', 'ref_key': 'ID', 'ref_value': 'Note'}, output_dir=out_dir))

    print(f'\nAll {len(CHECKS)} batch-summary checks passed.')

if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    run()
