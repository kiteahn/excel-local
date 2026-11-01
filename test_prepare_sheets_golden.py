"""Golden cho nhánh lookup của prepare_sheets (chạy thuần Python, không cần Excel).

Mục đích: ghim hành vi trước khi tách hàm để bản tách không đổi công thức,
nhãn cột, thứ tự cột, stats hay token. 26 kịch bản, mỗi kịch bản có kiểm tra
ngữ nghĩa riêng; toàn bộ đầu ra còn được so với ảnh chụp vàng
`results/prepare-sheets-lookup-golden.json`.

python -X utf8 test_prepare_sheets_golden.py                 # so với ảnh chụp vàng
python -X utf8 test_prepare_sheets_golden.py --update        # ghi lại ảnh chụp (chỉ khi cố ý đổi hành vi)
python -X utf8 test_prepare_sheets_golden.py --app old.py    # chạy cùng bộ vàng trên bản app khác
"""
from __future__ import annotations

import argparse
import copy
import importlib.util
import json
import sys
from pathlib import Path

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

ROOT = Path(__file__).resolve().parent
GOLDEN = ROOT / 'results' / 'prepare-sheets-lookup-golden.json'
TOKEN = 'gold04'


def load_app(path=None):
    if not path:
        import app
        return app
    path = Path(path)
    spec = importlib.util.spec_from_file_location('app_under_test', path)
    module = importlib.util.module_from_spec(spec)
    sys.modules['app_under_test'] = module
    spec.loader.exec_module(module)
    return module


def snapshot(headers, rows, formats=None, sheet='Nguon'):
    return {
        'headers': list(headers),
        'data': [list(row) for row in rows],
        'metadata': {'sheets': {sheet: {'formats': dict(formats or {})}}},
        'sheet': sheet,
        'header_row': 1,
        'hash': 'hash-' + sheet,
    }


def reference(headers, rows, formats=None):
    return snapshot(headers, rows, formats, sheet='Danh_muc')


def plan(**overrides):
    base = {'action': 'lookup', 'key': 'Mã', 'ref_key': 'Mã', 'ref_value': 'Tên'}
    base.update(overrides)
    return base


def expect(condition, message):
    if not condition:
        raise AssertionError(message)


def check_basic(sheets, name, stats):
    result = sheets[-2]
    expect(result['rows'][0][-3:] == ['Giá trị đối chiếu', 'Trạng thái đối chiếu', 'Khóa chuẩn hóa'], 'append labels')
    expect(result['rows'][1][3:] == ['KHOP', '00123'], 'matched status')
    expect(stats['matched'] == 2 and stats['missing'] == 0 and stats['duplicate'] == 0, 'basic stats')
    formula = sheets[-2]['formulas'][0]['formula']
    expect(formula == '=IF(INDEX(\'_EL_gold04\'!$B$2:$B$3,MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(E2,"~","~~"),"*","~*"),"?","~?"),\'_EL_gold04\'!$A$2:$A$3,0))="","",INDEX(\'_EL_gold04\'!$B$2:$B$3,MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(E2,"~","~~"),"*","~*"),"?","~?"),\'_EL_gold04\'!$A$2:$A$3,0)))', 'basic formula')
    expect(sheets[0]['rows'][0] == ['Khoa_chuan_hoa', 'Gia_tri'], 'helper header')


def check_clean_append(sheets, name, stats):
    result = sheets[-2]
    expect(result['rows'][0] == ['Mã', 'Nhóm', 'Họ tên'], 'append header')
    expect(result['rows'][1][2] is None, 'appended cell blank')
    expect(result['formulas'][0]['col'] == 3, 'formula at appended column')
    expect(sheets[0]['rows'][0] == ['Khoa_chuan_hoa', 'Gia_tri', '', 'Khoa_nguon'], 'helper 4 cols')
    expect(sheets[0]['rows'][1][3] == '00123', 'helper source key')


def check_clean_inplace(sheets, name, stats):
    result = sheets[-2]
    expect(result['rows'][0] == ['Mã', 'Nhóm', 'Ghi chú'], 'inplace headers unchanged')
    expect(result['rows'][1][2] is None, 'inplace target cleared')
    expect(result['formulas'][0]['col'] == 3, 'inplace formula column')


def check_wildcard(sheets, name, stats):
    expect(stats['matched'] == 1 and stats['missing'] == 1, 'wildcard stats')
    formula = sheets[-2]['formulas'][0]['formula']
    expect('SUBSTITUTE' in formula and '"*"' in formula and '"?"' in formula, 'wildcard escaping kept')
    expect(sheets[-2]['rows'][1][2] == 'KHOP' and sheets[-2]['rows'][2][2] == 'KHONG_TIM_THAY', 'literal wildcard only')


SCENARIOS = [
    {
        'name': 'basic_match_and_formula',
        'source': snapshot(['Mã', 'Nhóm'], [['00123', 'A'], ['00124', 'B']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An'], ['00124', 'Bình']]),
        'plan': plan(),
        'check': check_basic,
    },
    {
        'name': 'duplicate_reference_key',
        'source': snapshot(['Mã'], [['00124'], ['00123']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An'], ['00124', 'Bình'], ['00124', 'Cúc']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['duplicate'] == 1 and st['matched'] == 1 and st['reference_duplicate_keys'] == 1, 'duplicate stats'),
    },
    {
        'name': 'missing_and_blank_keys',
        'source': snapshot(['Mã'], [['999'], [None], ['']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['missing'] == 3 and st['blank_keys'] == 2, 'missing/blank stats'),
    },
    {
        'name': 'all_statuses_mixed',
        'source': snapshot(['Mã'], [['00123'], [None], ['999'], ['00124'], ['00124']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An'], ['00124', 'Bình'], ['00124', 'Cúc']]),
        'plan': plan(),
        'check': lambda s, n, st: expect((st['matched'], st['missing'], st['duplicate'], st['blank_keys']) == (1, 2, 2, 1), 'mixed stats'),
    },
    {
        'name': 'clean_output_append_new_column',
        'source': snapshot(['Mã', 'Nhóm'], [['00123', 'A'], ['00124', 'B']]),
        'reference': reference(['Mã', 'Họ tên'], [['00123', 'An'], ['00124', 'Bình']]),
        'plan': plan(ref_value='Họ tên', clean_output=True),
        'check': check_clean_append,
    },
    {
        'name': 'clean_output_inplace_target',
        'source': snapshot(['Mã', 'Nhóm', 'Ghi chú'], [['00123', 'A', 'giữ'], ['00124', 'B', 'x']]),
        'reference': reference(['Mã', 'Họ tên'], [['00123', 'An'], ['00124', 'Bình']]),
        'plan': plan(ref_value='Họ tên', value='Ghi chú', clean_output=True),
        'check': check_clean_inplace,
    },
    {
        'name': 'clean_output_label_collision',
        'source': snapshot(['Mã', 'Họ tên', 'Họ tên (đối chiếu)'], [['00123', 'cũ', 'x']]),
        'reference': reference(['Mã', 'Họ tên'], [['00123', 'An']]),
        'plan': plan(ref_value='Họ tên', value='Tên chuẩn', clean_output=True),
        'check': lambda s, n, st: expect(s[-2]['rows'][0][-1] == 'Họ tên (đối chiếu) (mới)', 'collision suffix'),
    },
    {
        'name': 'duplicate_source_keys_both_matched',
        'source': snapshot(['Mã'], [['00123'], ['00123']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 2, 'duplicate source keys'),
    },
    {
        'name': 'literal_wildcard_keys',
        'source': snapshot(['Mã'], [['a*b'], ['axb']]),
        'reference': reference(['Mã', 'Tên'], [['a*b', 'Sao']]),
        'plan': plan(),
        'check': check_wildcard,
    },
    {
        'name': 'numeric_text_leading_zero',
        'source': snapshot(['Mã'], [[123], [5]], formats={'A2': '00000', 'A3': '00000'}),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An'], ['00005', 'Dũng']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 2 and s[-2]['rows'][1][-1] == '00123' and s[-2]['rows'][2][-1] == '00005', 'leading zero'),
    },
    {
        'name': 'numeric_general_key',
        'source': snapshot(['Mã'], [[123], [45.5]]),
        'reference': reference(['Mã', 'Tên'], [[123, 'An'], ['45.5', 'Bình']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 2, 'numeric general'),
    },
    {
        'name': 'negative_float_key',
        'source': snapshot(['Mã'], [[-2.5]]),
        'reference': reference(['Mã', 'Tên'], [['-2.5', 'Âm']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 1, 'negative float key'),
    },
    {
        'name': 'unicode_case_and_space',
        'source': snapshot(['Mã'], [['  đà nẵng '], ['ĐÀ NẴNG']]),
        'reference': reference(['Mã', 'Tên'], [['Đà Nẵng', 'Trùng']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 2, 'unicode normalization'),
    },
    {
        'name': 'zero_width_and_nbsp',
        'source': snapshot(['Mã'], [['001\u200b23'], [' 00123 ']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'Sạch']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 2, 'invisible chars'),
    },
    {
        'name': 'allow_dates_string_dmy',
        'source': snapshot(['Mã'], [['01/02/2024'], ['2024-03-04']]),
        'reference': reference(['Mã', 'Tên'], [['2024-02-01', 'Tháng 2'], ['2024-03-04', 'Tháng 3']]),
        'plan': plan(allow_dates=True),
        'check': lambda s, n, st: expect(st['matched'] == 2, 'string dates'),
    },
    {
        'name': 'allow_dates_numeric_serial',
        'source': snapshot(['Mã'], [[45323]], formats={'A2': 'dd/mm/yyyy'}),
        'reference': reference(['Mã', 'Tên'], [['2024-02-01', 'Tháng 2']]),
        'plan': plan(allow_dates=True),
        'check': lambda s, n, st: expect(st['matched'] == 1, 'numeric date serial'),
    },
    {
        'name': 'date_format_without_permission',
        'source': snapshot(['Mã'], [[45323]], formats={'A2': 'dd/mm/yyyy'}),
        'reference': reference(['Mã', 'Tên'], [['2024-02-01', 'Tháng 2']]),
        'plan': plan(),
        'expect_error': 'ngày/giờ',
    },
    {
        'name': 'blank_reference_value',
        'source': snapshot(['Mã'], [['00123']]),
        'reference': reference(['Mã', 'Tên'], [['00123', None]]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 1 and s[-2]['formulas'][0]['expected'] == '' and s[0]['rows'][1][1] is None, 'blank ref value'),
    },
    {
        'name': 'false_and_zero_reference_values',
        'source': snapshot(['Mã'], [['001'], ['1']]),
        'reference': reference(['Mã', 'Tên'], [['001', False], ['1', 0]]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 2 and s[-2]['formulas'][0]['expected'] is False and s[-2]['formulas'][1]['expected'] == 0, 'false/zero kept'),
    },
    {
        'name': 'empty_reference_table',
        'source': snapshot(['Mã'], [['00123'], ['00124']]),
        'reference': reference(['Mã', 'Tên'], []),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['matched'] == 0 and st['missing'] == 2 and s[0]['rows'] == [['Khoa_chuan_hoa', 'Gia_tri'], ['', None]], 'empty ref placeholder'),
    },
    {
        'name': 'reference_blank_key_ignored',
        'source': snapshot(['Mã'], [['00123']]),
        'reference': reference(['Mã', 'Tên'], [[None, 'Bỏ'], ['00123', 'An'], ['00123', 'Bình']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(st['duplicate'] == 1 and st['reference_duplicate_keys'] == 1, 'blank ref ignored'),
    },
    {
        'name': 'header_label_collision_append_mode',
        'source': snapshot(['Mã', 'Giá trị đối chiếu', 'Trạng thái đối chiếu', 'Khóa chuẩn hóa'], [['00123', 'cũ', 'x', 'y']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(s[-2]['rows'][0][-3:] == ['Giá trị đối chiếu (mới)', 'Trạng thái đối chiếu (mới)', 'Khóa chuẩn hóa (mới)'], 'append collision'),
    },
    {
        'name': 'clean_output_header_collision_no_target',
        'source': snapshot(['Mã', 'Tên', 'Tên (đối chiếu)', 'Tên (đối chiếu) (mới)'], [['00123', 'a', 'b', 'c']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An']]),
        'plan': plan(ref_value='Tên', clean_output=True),
        'check': lambda s, n, st: expect(s[-2]['rows'][0][-1] == 'Tên (đối chiếu) (mới) (mới)', 'double collision suffix'),
    },
    {
        'name': 'audit_rows_and_token_names',
        'source': snapshot(['Mã'], [['00123']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An']]),
        'plan': plan(),
        'check': lambda s, n, st: (expect(n == 'Ket_qua_gold04', 'result name'),
                                   expect(s[0]['name'] == '_EL_gold04' and s[-1]['name'] == 'Kiem_tra_gold04', 'helper/audit names'),
                                   expect(['Thao_tac', 'lookup'] in s[-1]['rows'] and ['Cong_thuc_can_doi_soat', 1] in s[-1]['rows'], 'audit rows')),
    },
    {
        'name': 'stats_cover_all_counters',
        'source': snapshot(['Mã'], [['00123'], [None], ['999'], ['00124']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An'], ['00124', 'Bình'], ['00124', 'Cúc']]),
        'plan': plan(),
        'check': lambda s, n, st: expect(set(st) == {'action', 'source_rows', 'blank_keys', 'matched', 'missing', 'duplicate', 'reference_duplicate_keys'}, 'stats keys'),
    },
    {
        'name': 'source_input_not_mutated',
        'source': snapshot(['Mã', 'Ghi chú'], [['00123', 'giữ nguyên']]),
        'reference': reference(['Mã', 'Tên'], [['00123', 'An']]),
        'plan': plan(),
        'check': None,
    },
]


def normalize(value):
    return json.loads(json.dumps(value, ensure_ascii=False, default=str))


def run_scenario(mod, scenario):
    source = copy.deepcopy(scenario['source'])
    ref = copy.deepcopy(scenario['reference'])
    before = copy.deepcopy(source)
    if scenario.get('expect_error'):
        try:
            mod.prepare_sheets(source, ref, scenario['plan'], TOKEN)
        except ValueError as exc:
            if scenario['expect_error'] not in str(exc):
                raise AssertionError(f"{scenario['name']}: wrong error {exc!r}")
            return {'error': 'ValueError', 'input_unchanged': source == before}
        raise AssertionError(f"{scenario['name']}: expected ValueError")
    sheets, name, stats = mod.prepare_sheets(source, ref, scenario['plan'], TOKEN)
    if scenario.get('check'):
        scenario['check'](sheets, name, stats)
    if scenario['name'] == 'source_input_not_mutated':
        expect(source == before, 'source snapshot mutated in place')
    return normalize({'sheets': sheets, 'name': name, 'stats': stats})


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--update', action='store_true')
    parser.add_argument('--app', default=None, help='đường dẫn app.py khác (bản trước khi tách)')
    parser.add_argument('--golden', default=str(GOLDEN))
    args = parser.parse_args()
    mod = load_app(args.app)
    try:
        mod.prepare_sheets
    except AttributeError:
        raise SystemExit(f'{args.app} không có prepare_sheets')
    outputs = [{'name': scenario['name'], 'output': run_scenario(mod, scenario)} for scenario in SCENARIOS]
    golden_path = Path(args.golden)
    if args.update:
        golden_path.parent.mkdir(exist_ok=True)
        golden_path.write_text(json.dumps(outputs, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f'Đã ghi {len(outputs)} kịch bản vào {golden_path}')
        return 0
    expected = json.loads(golden_path.read_text(encoding='utf-8'))
    if len(expected) != len(outputs):
        raise SystemExit(f'Số kịch bản vàng {len(expected)} khác {len(outputs)}')
    for want, got in zip(expected, outputs):
        if want['name'] != got['name']:
            raise SystemExit(f"Lệch tên kịch bản: {want['name']} vs {got['name']}")
        if want['output'] != got['output']:
            with open(golden_path.with_suffix('.diff.json'), 'w', encoding='utf-8') as stream:
                json.dump({'want': want, 'got': got}, stream, ensure_ascii=False, indent=2)
            raise SystemExit(f"FAIL: {got['name']} khác ảnh chụp vàng (xem {golden_path.with_suffix('.diff.json')})")
    print(f'OK: {len(outputs)} kịch bản lookup khớp ảnh chụp vàng ({golden_path.name}).')
    return 0


if __name__ == '__main__':
    sys.exit(main())
