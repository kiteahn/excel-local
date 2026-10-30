"""Unit tests for ai_client.sanitize_plan (added in round 15).

Mutation round 2 showed three surviving mutations around sanitize_plan and
the lookup gate; these tests close them:
- M2-3: unknown actions must be forced to 'clarify'.
- M2-5: key/ref columns must be fuzzy-matched to real column names.
(M2-4, the lookup and/or gate, is covered by a golden case in
test_infer_golden.py: a filter prompt without reference columns must stay a
filter, not become a lookup clarify.)
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import ai_client

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def run():
    out = ai_client.sanitize_plan({'action': 'delete_everything', 'key': 'X'}, ['X'], [])
    ok('S-01 action la bi ep ve clarify (M2-3)', out['action'] == 'clarify')

    out = ai_client.sanitize_plan({'action': 'lookup', 'key': 'X'}, ['X'], [])
    ok('S-02 action hop le duoc giu nguyen', out['action'] == 'lookup')

    out = ai_client.sanitize_plan({'action': 'lookup', 'key': 'ho ten'}, ['Họ tên'], [])
    ok('S-03 key khong dau duoc fuzzy match ve ten cot that (M2-5)', out['key'] == 'Họ tên')

    out = ai_client.sanitize_plan({'action': 'summary', 'key': 'Lớp', 'value': 'hoc phi'}, ['Lớp', 'Học phí'], [])
    ok('S-04 value cua summary duoc fuzzy match', out['value'] == 'Học phí')

    out = ai_client.sanitize_plan({'action': 'lookup', 'ref_key': 'ma sv', 'ref_value': 'ho ten'},
                                  ['MSSV'], ['Mã SV', 'Họ tên'])
    ok('S-05 ref_key/ref_value fuzzy match theo cot tham chieu', out['ref_key'] == 'Mã SV' and out['ref_value'] == 'Họ tên')

    out = ai_client.sanitize_plan({'action': 'filter', 'key': 'Lớp', 'selected_columns': 'ho ten, Lớp'},
                                  ['Họ tên', 'Lớp'], [])
    ok('S-06 filter selected_columns duoc chuan hoa ve cot that', out['selected_columns'] == 'Họ tên, Lớp')

    try:
        ai_client.sanitize_plan(['not-a-dict'], [], [])
        ok('S-07 plan khong phai dict bi tu choi', False)
    except ValueError:
        ok('S-07 plan khong phai dict bi tu choi', True)

    out = ai_client.sanitize_plan({'action': 'clean', 'key': 'Họ tên'}, ['Họ tên'], [])
    ok('S-08 cac truong bat buoc duoc dien mac dinh', all(k in out for k in ('action', 'key', 'value', 'ref_key', 'ref_value', 'message')))

    print('--- Plan cache key (round 25, M5-1) ---')
    k_a = ai_client.get_plan_cache_key('tinh tong', ('Lớp', 'Học phí'), ())
    k_b = ai_client.get_plan_cache_key('tinh tong', ('Khoa', 'Tiền'), ())
    ok('C-01 cache key khac nhau khi cot nguon khac nhau (M5-1)', k_a != k_b)
    k_c = ai_client.get_plan_cache_key('tinh tong', ('Học phí', 'Lớp'), ())
    ok('C-02 cache key khong phu thuoc thu tu cot', k_a == k_c)
    k_d = ai_client.get_plan_cache_key('tinh tong', ('Lớp', 'Học phí'), ('MSSV',))
    ok('C-03 cache key khac nhau khi cot tham chieu khac nhau', k_a != k_d)

    p1 = ai_client.request_plan('Đối chiếu MSSV để lấy Họ tên (cachetest25)', ['MSSV'], ['MSSV', 'Họ tên'], provider='heuristic')
    p2 = ai_client.request_plan('Đối chiếu MSSV để lấy Họ tên (cachetest25)', ['MaSV'], ['MaSV', 'Que quan'], provider='heuristic')
    ok('C-04 request_plan khong dung lai plan cache cua bo cot khac',
       p1.get('action') == 'lookup' and p2.get('key') == 'MaSV' and p2.get('ref_value') == 'Que quan')

    print(f'\nAll {len(CHECKS)} sanitize_plan checks passed.')

if __name__ == '__main__':
    run()
