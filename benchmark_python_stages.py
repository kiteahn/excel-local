"""Staged benchmark for the pure-Python parts of Excel Local Lazy.

Measures what can be measured without Excel Desktop: xlsx preflight, the
openpyxl fast reader (lenient inspect path), CSV inspect, and preview_plan
for lookup/summary/clean at 100 / 1,000 / 5,000 / 20,000 rows.

The native Excel bridge stages (inspect/execute through excel_bridge.ps1)
are NOT covered here; use benchmark.py on Windows for those.

Usage: python -X utf8 benchmark_python_stages.py [--json results.json]
"""
import argparse
import json
import sys
import tempfile
import time
from pathlib import Path

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import app
from test_app import fixture

SCALES = (100, 1000, 5000, 20000)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--json', default='')
    args = parser.parse_args()
    if app.load_xlsx_fast.__module__ and not _openpyxl_available():
        print('SKIP: openpyxl is not installed, the fast reader path cannot be measured.')
        return
    results = {}
    for rows in SCALES:
        print(f'===== rows={rows} =====', flush=True)
        stage = {}
        with tempfile.TemporaryDirectory(dir=app.WORK) as private:
            d = Path(private)
            src, ref = d / 'source.xlsx', d / 'reference.xlsx'
            fixture(src, {'Data': [['ID', 'Group', 'Amount', 'Note']] + [
                [f'{i:05}', f'G{i % 50}', i % 20 + 1, '  Doan  '] for i in range(rows)]})
            fixture(ref, {'Catalog': [['ID', 'Name']] + [
                [f'{i:05}', f'Name {i}'] for i in range(rows)]})

            def timed(name, fn):
                t0 = time.perf_counter()
                value = fn()
                stage[name] = round(time.perf_counter() - t0, 3)
                print(f'{name}: {stage[name]}s', flush=True)
                return value

            timed('preflight_xlsx', lambda: app.preflight(src))
            source = timed('inspect_source_lenient_fast', lambda: app.inspect_file(src, lenient=True))
            reference = timed('inspect_reference_lenient_fast', lambda: app.inspect_file(ref, lenient=True))
            for plan in [dict(action='lookup', key='ID', ref_key='ID', ref_value='Name'),
                         dict(action='summary', key='Group', value='Amount'),
                         dict(action='clean', key='Note')]:
                timed(f"preview_{plan['action']}",
                      lambda p=plan: app.preview_plan(source, reference, p, output_dir=d / 'output'))
            csv_path = d / 'data.csv'
            csv_path.write_text('ID,Group,Amount,Note\n' + '\n'.join(
                f'{i:05},G{i % 50},{i % 20 + 1},Doan' for i in range(rows)), encoding='utf-8')
            timed('inspect_csv', lambda: app.inspect_file(csv_path))
        results[str(rows)] = stage
    if args.json:
        Path(args.json).write_text(json.dumps(results, indent=2), encoding='utf-8')
        print(f'Wrote {args.json}')


def _openpyxl_available():
    try:
        import openpyxl  # noqa: F401
        return True
    except ImportError:
        return False


if __name__ == '__main__':
    main()
