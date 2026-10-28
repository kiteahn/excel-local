"""Comprehensive test suite for the v0.3 upgrades:
1. Auto-detect header row
2. Auto-healing merged and missing header columns
3. Safe external hyperlinks allowed in preflight
4. Lenient Vietnamese currency & number formatting
5. Date conversion in key
6. OpenAI API client & fuzzy column matching
"""
from decimal import Decimal
import json
from pathlib import Path
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import app
import ai_client
from test_app import fixture

RESULTS = []

def check(name, condition):
    assert condition, name
    RESULTS.append({'test': name, 'status': 'pass'})

def run_tests():
    # 1. Header auto-detection
    messy_rows = [
        ['BỘ GIÁO DỤC VÀ ĐÀO TẠO', None, None],
        ['TRƯỜNG ĐẠI HỌC BÁCH KHOA', None, None],
        [None, None, None],
        ['STT', 'Mã sinh viên', 'Họ và tên'],
        [1, '2023001', 'Nguyễn Văn A'],
        [2, '2023002', 'Trần Thị B']
    ]
    detected = app.detect_header_row(messy_rows)
    check('auto-detect table header row ignores title banner', detected == 4)

    # 2. Header auto-healing for merged/empty columns
    snapshot_raw = {
        'sheet': 'Data',
        'rows': [
            ['STT', 'Họ tên', None, 'Điểm thi', None],
            [1, 'Nguyễn', 'An', 8.5, 9.0]
        ]
    }
    healed = app.select_header(snapshot_raw, 1, auto_heal=True)
    check('merged/empty header columns healed automatically', healed['headers'] == ['STT', 'Họ tên', 'Họ tên_phụ', 'Điểm thi', 'Điểm thi_phụ'])
    check('healed snapshot marked with auto_healed flag', healed.get('auto_healed') is True)

    # 3. Duplicate headers disambiguated
    dup_raw = {
        'sheet': 'Data',
        'rows': [
            ['STT', 'Ghi chú', 'Ghi chú'],
            [1, 'Ghi chú 1', 'Ghi chú 2']
        ]
    }
    healed_dup = app.select_header(dup_raw, 1, auto_heal=True)
    check('duplicate headers disambiguated', healed_dup['headers'] == ['STT', 'Ghi chú', 'Ghi chú_2'])

    # 4. Vietnamese number parsing
    check('parse 1.500.000 đ', app.number('1.500.000 đ', lenient=True) == Decimal('1500000'))
    check('parse 1.234.567,89 VND', app.number('1.234.567,89 VND', lenient=True) == Decimal('1234567.89'))
    check('parse 8,5 (comma decimal)', app.number('8,5', lenient=True) == Decimal('8.5'))
    check('parse blank as 0 in lenient mode', app.number('', lenient=True) == Decimal('0'))
    check('parse None as 0 in lenient mode', app.number(None, lenient=True) == Decimal('0'))

    # 5. Date key conversion in lenient mode
    # Excel serial 45000 is 2023-03-15
    date_str = app.key_value(45000, 'dd/mm/yyyy', allow_dates=True)
    check('date serial converts to YYYY-MM-DD', date_str == '2023-03-15')

    # 6. Safe hyperlinks allowed in preflight
    with tempfile.TemporaryDirectory(dir=app.WORK) as private:
        hl_path = Path(private) / 'hyperlinks.xlsx'
        fixture(hl_path, {'Data': [['Email'], ['test@edu.vn']]})
        with zipfile.ZipFile(hl_path, 'a') as archive:
            archive.writestr('xl/worksheets/_rels/sheet1.xml.rels',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                'Target="mailto:test@edu.vn" TargetMode="External"/>'
                '</Relationships>')
        meta = app.preflight(hl_path)
        check('safe hyperlink allowed in preflight', 'Data' in meta['sheets'])

    # 7. Fuzzy column matching in AI client
    valid_cols = ['Mã sinh viên', 'Họ và tên', 'Điểm trung bình', 'Lớp học']
    check('fuzzy match MSSV to Mã sinh viên', ai_client.fuzzy_match_column('Mã sinh viên', valid_cols) == 'Mã sinh viên')
    check('fuzzy match case-insensitive', ai_client.fuzzy_match_column('họ và tên', valid_cols) == 'Họ và tên')
    check('fuzzy match close string', ai_client.fuzzy_match_column('Điểm trung binh', valid_cols) == 'Điểm trung bình')

    # 8. Sanitize plan
    raw_plan = {'action': 'lookup', 'key': 'họ và tên', 'ref_key': 'mssv', 'ref_value': 'điểm trung bình'}
    sanitized = ai_client.sanitize_plan(raw_plan, valid_cols, valid_cols)
    check('plan sanitized with fuzzy matched columns', sanitized['key'] == 'Họ và tên' and sanitized['ref_value'] == 'Điểm trung bình')

    # 9. XLS preservation check bypass
    with tempfile.TemporaryDirectory(dir=app.WORK) as private:
        dummy_xls = Path(private) / 'test.xls'
        dummy_xls.write_bytes(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00' * 504) # OLE CFB header
        pres = app.preserve_check(dummy_xls, dummy_xls)
        check('preserve_check handles binary XLS without crash', pres['unchanged'] is True and pres['original_sheets'] == 1)

    # 10. Config persistence
    cfg = ai_client.get_config()
    check('config has required provider keys', 'provider' in cfg and 'openai_model' in cfg)

    # 11. OpenAI response fence parsing & ai_client re module
    check('ai_client has re module imported', hasattr(ai_client, 're'))
    raw_fence = '```json\n{"action": "lookup", "key": "Mã sinh viên", "value": "", "ref_key": "Mã sinh viên", "ref_value": "Họ và tên", "message": "Đối chiếu MSSV"}\n```'
    cleaned = re.sub(r'^```(?:json)?\s*', '', raw_fence, flags=re.MULTILINE)
    cleaned = re.sub(r'\s*```$', '', cleaned, flags=re.MULTILINE)
    match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    parsed = json.loads(match.group(0))
    sanitized = ai_client.sanitize_plan(parsed, valid_cols, valid_cols)
    # 12. Clean Output mode for Lookup
    src_snap = {
        'path': 'source.xlsx', 'sheet': 'Data', 'hash': 'h1', 'header_row': 1,
        'headers': ['STT', 'Mã hồ sơ', 'Ghi chú'],
        'data': [[1, '00123', 'A'], [2, '00124', 'B'], [3, '00999', 'C']],
        'metadata': {'sheets': {'Data': {'formats': {}}}}
    }
    ref_snap = {
        'path': 'ref.xlsx', 'sheet': 'Danh_muc', 'hash': 'h2', 'header_row': 1,
        'headers': ['Mã hồ sơ', 'Họ tên'],
        'data': [['00123', 'Nguyễn An'], ['00124', 'Trần Bình'], ['00124', 'Trần Cúc']],
        'metadata': {'sheets': {'Danh_muc': {'formats': {}}}}
    }
    clean_plan = {'action': 'lookup', 'key': 'Mã hồ sơ', 'ref_key': 'Mã hồ sơ', 'ref_value': 'Họ tên', 'clean_output': True}
    clean_sheets, clean_name, clean_stats = app.prepare_sheets(src_snap, ref_snap, clean_plan, 'testclean')
    clean_result = next(s for s in clean_sheets if s['name'] == clean_name)
    clean_helper = next(s for s in clean_sheets if s['name'] == '_EL_testclean')

    check('clean output adds only 1 result column', len(clean_result['rows'][0]) == 4 and clean_result['rows'][0][-1] == 'Họ tên')
    check('clean output excludes status header', 'Trạng thái đối chiếu' not in clean_result['rows'][0])
    check('clean output excludes helper key header', 'Khóa chuẩn hóa' not in clean_result['rows'][0])
    # Verify no raw status strings in any row of result table
    all_cells = [cell for row in clean_result['rows'][1:] for cell in row]
    check('clean output has no technical status tags', not any(tag in all_cells for tag in ('KHOP', 'KHONG_TIM_THAY', 'KHOA_TRUNG', 'KHOA_TRONG')))
    check('clean output helper holds source keys in Col D', len(clean_helper['rows'][0]) == 4 and clean_helper['rows'][0][3] == 'Khoa_nguon')
    check('clean output formula MATCH points to helper Col D', any('_EL_testclean\'!$D$' in f['formula'] for f in clean_result['formulas']))

    # 13. Legacy full audit mode
    legacy_plan = {'action': 'lookup', 'key': 'Mã hồ sơ', 'ref_key': 'Mã hồ sơ', 'ref_value': 'Họ tên', 'clean_output': False}
    leg_sheets, leg_name, leg_stats = app.prepare_sheets(src_snap, ref_snap, legacy_plan, 'testleg')
    leg_result = next(s for s in leg_sheets if s['name'] == leg_name)
    check('legacy mode preserves status header', 'Trạng thái đối chiếu' in leg_result['rows'][0])
    # 14. Graceful AI clarification when columns or reference catalog are missing
    clarify_missing_ref = app.validate_plan({'action': 'lookup', 'key': 'Mã hồ sơ', 'ref_key': 'Mã hồ sơ', 'ref_value': 'Họ tên'}, src_snap, None, allow_clarify=True)
    check('missing reference returns clarify when allow_clarify is True', clarify_missing_ref['action'] == 'clarify' and 'danh mục' in clarify_missing_ref['message'])
    
    clarify_missing_col = app.validate_plan({'action': 'clean', 'key': 'Cột_Không_Tồn_Tại'}, src_snap, None, allow_clarify=True)
    check('missing column returns clarify when allow_clarify is True', clarify_missing_col['action'] == 'clarify' and 'Không tìm thấy' in clarify_missing_col['message'])

    try:
        app.validate_plan({'action': 'clean', 'key': 'Cột_Không_Tồn_Tại'}, src_snap, None, allow_clarify=False)
        check('missing column raises ValueError when allow_clarify is False', False)
    except ValueError:
        check('missing column raises ValueError when allow_clarify is False', True)

    print(f"All {len(RESULTS)} upgrade checks passed successfully!")

if __name__ == '__main__':
    run_tests()
