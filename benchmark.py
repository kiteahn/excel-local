"""Repeatable native Excel benchmark on synthetic data; no user files."""
import argparse
import json
from pathlib import Path
import tempfile
import time
import app
from test_app import fixture


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--rows', type=int, default=1000)
    parser.add_argument('--label', default='current')
    args = parser.parse_args()
    if not 1 <= args.rows <= 20000:
        parser.error('rows must be between 1 and 20000')
    app.WORK.mkdir(exist_ok=True)
    measurements = {}
    def timed(name, function):
        started = time.perf_counter()
        value = function()
        measurements[name] = round(time.perf_counter()-started, 3)
        print(name, measurements[name], flush=True)
        return value
    with tempfile.TemporaryDirectory(prefix='benchmark-', dir=app.WORK) as private:
        directory = Path(private)
        source_path, reference_path = directory/'source.xlsx', directory/'reference.xlsx'
        fixture(source_path, {'Data':[['ID','Group','Amount','Note']]+[
            [f'{i:05}',f'G{i}',{'value':i%20+1,'format':'00000'},'  Đoàn  ']
            for i in range(args.rows)]})
        fixture(reference_path, {'Catalog':[['ID','Name']]+[
            [f'{i:05}',f'Name {i}'] for i in range(args.rows)]})
        timed('preflight', lambda:app.preflight(source_path))
        source = timed('inspect_source', lambda:app.inspect_file(source_path))
        reference = timed('inspect_reference', lambda:app.inspect_file(reference_path))
        checks = {}
        phases = {}
        for plan in [dict(action='lookup',key='ID',ref_key='ID',ref_value='Name'),
                     dict(action='summary',key='Group',value='Amount'),
                     dict(action='clean',key='Note')]:
            preview = app.preview_plan(source,reference,plan,output_dir=directory/'output')
            report = timed(plan['action'], lambda:app.execute(source,reference,plan,output_dir=directory/'output',approval_token=preview['approval_token']))
            checks[plan['action']] = report['formula_checks']
            phases[plan['action']] = report.get('timings',{})
    result = dict(label=args.label,rows=args.rows,seconds=measurements,formula_checks=checks,bridge_phases=phases,
                  app_sha256=app.file_hash(app.ROOT/'app.py'),bridge_sha256=app.file_hash(app.ROOT/'excel_bridge.ps1'))
    path = app.ROOT/f'benchmark-{args.label}-{args.rows}.json'
    path.write_text(json.dumps(result,indent=2),encoding='utf-8')
    print(path,flush=True)


if __name__ == '__main__':
    main()
