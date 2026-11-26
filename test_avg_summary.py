"""AVG summary support (agg='avg'), added 2026-11-04 after Q4 real-like
cases L04/L05 showed average requests were unsupported."""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def snapshot(headers, rows, sheet='Nguon'):
    return {'headers': list(headers), 'data': [list(r) for r in rows],
            'metadata': {'sheets': {sheet: {'formats': {}}}}, 'sheet': sheet,
            'header_row': 1, 'hash': 'h-' + sheet, 'path': 'x.xlsx'}

def run():
    src = snapshot(['Lớp', 'Điểm'], [['A', 8], ['A', 6], ['B', 10], ['B', 4]])
    plan = {'action': 'summary', 'key': 'Lớp', 'value': 'Điểm', 'agg': 'avg'}
    sheets, name, stats = app.prepare_sheets(src, None, plan, 'tok')
    result = next(s for s in sheets if s['name'] == name)
    ok('A-01 agg duoc ghi vao stats', stats.get('agg') == 'avg')
    ok('A-02 tieu de cot 3 la Trung_binh', result['rows'][0][2] == 'Trung_binh')
    avg_formulas = [f for f in result['formulas'] if f['col'] == 3]
    ok('A-03 dung cong thuc AVERAGE cho ca 2 nhom',
       len(avg_formulas) == 2 and all('AVERAGE(' in f['formula'] for f in avg_formulas))
    expected = sorted(f['expected'] for f in avg_formulas)
    ok('A-04 expected doc lap dung: 7.0 va 7.0', expected == [7.0, 7.0])

    plan_sum = {'action': 'summary', 'key': 'Lớp', 'value': 'Điểm'}
    sheets2, name2, stats2 = app.prepare_sheets(src, None, plan_sum, 'tok')
    result2 = next(s for s in sheets2 if s['name'] == name2)
    sum_formulas = [f for f in result2['formulas'] if f['col'] == 3]
    ok('A-05 mac dinh van la SUM voi expected 14/14',
       result2['rows'][0][2] == 'Tong' and all('SUM(' in f['formula'] for f in sum_formulas)
       and sorted(f['expected'] for f in sum_formulas) == [14.0, 14.0] and stats2.get('agg') == 'sum')

    try:
        app.prepare_sheets(src, None, {'action': 'summary', 'key': 'Lớp', 'value': 'Điểm', 'agg': 'median'}, 'tok')
        ok('A-06 agg khong ho tro bi tu choi cung', False)
    except ValueError:
        ok('A-06 agg khong ho tro bi tu choi cung', True)

    print(f'\nAll {len(CHECKS)} avg-summary checks passed.')

if __name__ == '__main__':
    run()
