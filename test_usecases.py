"""Round 11 (layer-2): use-case checks that run without Excel Desktop.

Each case maps to USECASES.md. Cases that need native Excel are NOT here;
they are listed as CHUA CHAY DUOC in USECASES.md instead of being skipped
silently. Snapshot dicts follow the same shape as test_app.snapshot().
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import tempfile
import unicodedata
import zipfile
from decimal import Decimal
from pathlib import Path

import app
from test_app import fixture

CHECKS = []

def check(name, condition):
    assert condition, name
    CHECKS.append(name)
    print(f'  [PASS] {name}')

def rejects(name, call):
    try:
        call()
    except (ValueError, RuntimeError):
        CHECKS.append(name)
        print(f'  [PASS] {name}')
        return
    raise AssertionError('Expected rejection: ' + name)

def snap(headers, rows, formats=None):
    return {'headers': headers, 'data': rows,
            'metadata': {'sheets': {'Data': {'formats': formats or {}}}},
            'sheet': 'Data', 'header_row': 1, 'hash': 'test'}

def run():
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        d = Path(td)

        print('--- CSV inspect variants (UC-01..UC-06) ---')
        p = d / 'u16.csv'
        p.write_bytes('ID,Name\n001,An\n'.encode('utf-16'))
        s = app.inspect_file(p)
        check('UC-01 CSV UTF-16 BOM doc duoc', s['headers'] == ['ID', 'Name'] and s['data'] == [['001', 'An']])

        p = d / 'semi.csv'
        p.write_text('ID;Name\n001;An\n', encoding='utf-8')
        s = app.inspect_file(p)
        check('UC-02 CSV dau cham phay', s['headers'] == ['ID', 'Name'] and s['data'][0][1] == 'An')

        p = d / 'tab.csv'
        p.write_text('ID\tName\n001\tAn\n', encoding='utf-8')
        s = app.inspect_file(p)
        check('UC-03 CSV dau tab', s['headers'] == ['ID', 'Name'])

        p = d / 'banner.csv'
        p.write_text('Cộng hòa xã hội chủ nghĩa Việt Nam\nĐộc lập - Tự do - Hạnh phúc\nID,Name\n001,An\n', encoding='utf-8')
        s = app.inspect_file(p, header_row='auto')
        check('UC-04 CSV banner hanh chinh + tu tim dong tieu de', s['headers'] == ['ID', 'Name'] and s['data'] == [['001', 'An']])

        p = d / 'wide.csv'
        p.write_text(','.join(f'C{i}' for i in range(61)) + '\n', encoding='utf-8')
        rejects('UC-05 CSV 61 cot bi tu choi', lambda: app.inspect_file(p))

        p = d / 'max.csv'
        p.write_text('ID\n' + '\n'.join(str(i) for i in range(20000)) + '\n', encoding='utf-8')
        s = app.inspect_file(p)
        check('UC-06 CSV dung 20000 dong duoc chap nhan', len(s['data']) == 20000)

        print('--- Header handling (UC-07..UC-10) ---')
        rows = [['Cộng hòa xã hội chủ nghĩa Việt Nam'], ['Độc lập - Tự do - Hạnh phúc'],
                ['DANH SÁCH SINH VIÊN'], ['ID', 'Name'], ['001', 'An']]
        check('UC-07 detect_header_row bo qua banner', app.detect_header_row(rows) == 4)
        raw = {'sheet': 'Data', 'sheets': ['Data'], 'rows': [['ID', ' ID '], ['1', '2']],
               'metadata': {'sheets': {'Data': {'formats': {}}}}}
        rejects('UC-08 cot trung ten sau khi cat khoang trang bi tu choi',
                lambda: app.select_header(raw, 1))
        rejects('UC-09 dong tieu de vuot qua du lieu bi tu choi',
                lambda: app.select_header(raw, 99))
        healed = app.select_header(
            {'sheet': 'Data', 'sheets': ['Data'],
             'rows': [['ID', 'ID', ''], ['1', '2', '3']],
             'metadata': {'sheets': {'Data': {'formats': {}}}}}, 1, auto_heal=True)
        check('UC-10 auto_heal xu ly cot trung/khuyet', len(healed['headers']) == 3 and len(set(healed['headers'])) == 3)

        print('--- Lookup key semantics via prepare_sheets (UC-11..UC-16) ---')
        src = snap(['ID', 'Note'], [[' abc ', 'x']])
        ref = snap(['ID', 'Name'], [['ABC', 'An']])
        _, _, st = app.prepare_sheets(src, ref, {'action': 'lookup', 'key': 'ID', 'ref_key': 'ID', 'ref_value': 'Name'}, 't')
        check('UC-11 khoa thua khoang trang + khac hoa thuong van khop', st['matched'] == 1)

        nfc = unicodedata.normalize('NFC', 'Đoàn')
        nfd = unicodedata.normalize('NFD', 'Đoàn')
        src = snap(['ID'], [[nfd]])
        ref = snap(['ID', 'Name'], [[nfc, 'An']])
        _, _, st = app.prepare_sheets(src, ref, {'action': 'lookup', 'key': 'ID', 'ref_key': 'ID', 'ref_value': 'Name'}, 't')
        check('UC-12 khoa NFD khop khoa NFC', st['matched'] == 1)

        src = snap(['ID'], [[123]])
        ref = snap(['ID', 'Name'], [['00123', 'An']])
        _, _, st = app.prepare_sheets(src, ref, {'action': 'lookup', 'key': 'ID', 'ref_key': 'ID', 'ref_value': 'Name'}, 't')
        check('UC-13 so 123 (General) khong khop chuoi 00123', st['matched'] == 0 and st['missing'] == 1)

        src = snap(['ID'], [[123]], formats={'A2': '00000'})
        ref = snap(['ID', 'Name'], [['00123', 'An']])
        _, _, st = app.prepare_sheets(src, ref, {'action': 'lookup', 'key': 'ID', 'ref_key': 'ID', 'ref_value': 'Name'}, 't')
        check('UC-14 so 123 dinh dang 00000 khop chuoi 00123', st['matched'] == 1)

        src = snap(['ID'], [[''], ['001']])
        ref = snap(['ID', 'Name'], [['001', 'An']])
        _, _, st = app.prepare_sheets(src, ref, {'action': 'lookup', 'key': 'ID', 'ref_key': 'ID', 'ref_value': 'Name'}, 't')
        check('UC-15 khoa trong duoc dem rieng', st['blank_keys'] == 1 and st['matched'] == 1)

        src = snap(['ID'], [['A*B']])
        ref = snap(['ID', 'Name'], [['A*B', 'An'], ['AXB', 'Binh']])
        _, _, st = app.prepare_sheets(src, ref, {'action': 'lookup', 'key': 'ID', 'ref_key': 'ID', 'ref_value': 'Name'}, 't')
        check('UC-16 ky tu * trong khoa duoc hieu theo nghia den', st['matched'] == 1)

        print('--- Filter (UC-17..UC-19) ---')
        src = snap(['Name', 'Chức vụ', 'Điểm'], [['An', 'lớp trưởng', 9], ['Bình', 'thành viên', 7], ['Cường', 'tổ trưởng', 8]])
        _, _, st = app.prepare_sheets(src, None, {'action': 'filter', 'key': 'Chức vụ', 'value': 'lớp trưởng, tổ trưởng'}, 't')
        check('UC-17 loc theo tu khoa chuc vu', st['filtered_rows'] == 2)
        _, _, st = app.prepare_sheets(src, None, {'action': 'filter', 'key': 'Điểm', 'value': '>=8'}, 't')
        check('UC-18 loc dieu kien so >=8', st['filtered_rows'] == 2)
        _, _, st = app.prepare_sheets(src, None, {'action': 'filter', 'key': 'Chức vụ', 'value': 'khác rỗng'}, 't')
        check('UC-19 loc khac rong', st['filtered_rows'] == 3)

        print('--- Summary numbers (UC-20..UC-24) ---')
        src = snap(['Group', 'Amount'], [['A', '1.234.567'], ['A', '2.500.000'], ['B', '100']])
        _, _, st = app.prepare_sheets(src, None, {'action': 'summary', 'key': 'Group', 'value': 'Amount', 'lenient_numbers': True}, 't')
        check('UC-20 tong tien VN dang 1.234.567 o che do lenient', Decimal(st['independent_total']) == Decimal('3734667'))
        check('UC-20b mot dau cham + 3 so duoc hieu la thap phan (500.000 -> 500)',
              app.parse_lenient_number('500.000') == Decimal('500'))
        rejects('UC-21 so dang 1,234 bi tu choi o che do chat',
                lambda: app.prepare_sheets(snap(['G', 'V'], [['A', '1,234']]), None, {'action': 'summary', 'key': 'G', 'value': 'V'}, 't'))
        rejects('UC-22 tong vuot 1e14 bi tu choi',
                lambda: app.prepare_sheets(snap(['G', 'V'], [['A', 99999999999999], ['A', 99999999999999]]), None, {'action': 'summary', 'key': 'G', 'value': 'V'}, 't'))
        src = snap(['G', 'V'], [['A', -50], ['A', 100]])
        _, _, st = app.prepare_sheets(src, None, {'action': 'summary', 'key': 'G', 'value': 'V'}, 't')
        check('UC-23 tong co so am', Decimal(st['independent_total']) == Decimal('50'))
        rejects('UC-24 o so trong khong tu doi thanh 0',
                lambda: app.prepare_sheets(snap(['G', 'V'], [['A', '']]), None, {'action': 'summary', 'key': 'G', 'value': 'V'}, 't'))

        print('--- Dates, clean, token gate (UC-25..UC-28) ---')
        check('UC-25 key_value ngay serial khi allow_dates', app.key_value(45000, 'dd/mm/yyyy', allow_dates=True) == '2023-03-15')
        rejects('UC-26 khoa so dinh dang ngay bi tu choi khi chua cho phep ngay',
                lambda: app.key_value(45000, 'dd/mm/yyyy', allow_dates=False))
        src = snap(['Note'], [['  Đoàn  Văn  ']])
        sheets, name, st = app.prepare_sheets(src, None, {'action': 'clean', 'key': 'Note'}, 't')
        result = next(s for s in sheets if s['name'] == name)
        check('UC-27 clean cat khoang trang dau/cuoi + NFC', result['rows'][1][0] == unicodedata.normalize('NFC', 'Đoàn  Văn') and st['changed_cells'] == 1)

        p = d / 'gate.csv'
        p.write_text('ID,Note\n001,x\n', encoding='utf-8')
        local = app.inspect_file(p)
        plan = {'action': 'clean', 'key': 'Note'}
        rejects('UC-28 execute voi token sai bi chan truoc khi cham Excel',
                lambda: app.execute(local, None, plan, output_dir=d / 'out', approval_token='sai-token'))

        print('--- Preflight OOXML (UC-29..UC-31) ---')
        good = d / 'good.xlsx'
        fixture(good, {'Data': [['ID'], ['001']]})
        app.preflight(good)
        check('UC-29 preflight file sach duoc chap nhan', True)
        macro = d / 'macro.xlsx'
        fixture(macro, {'Data': [['ID'], ['001']]})
        with zipfile.ZipFile(macro, 'a') as z:
            z.writestr('xl/vbaProject.bin', b'fake')
        rejects('UC-30 workbook chua macro bi preflight tu choi', lambda: app.preflight(macro))
        ext = d / 'ext.xlsx'
        fixture(ext, {'Data': [['ID'], ['001']]})
        with zipfile.ZipFile(ext, 'a') as z:
            z.writestr('xl/externalLinks/externalLink1.xml', '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>')
        rejects('UC-31 workbook chua lien ket ngoai bi preflight tu choi', lambda: app.preflight(ext))

    print(f'\nAll {len(CHECKS)} use-case checks passed.')

if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    run()
