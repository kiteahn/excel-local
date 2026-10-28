"""Comprehensive test suite for Zero-Leakage Privacy Guard, Security Guardrails, and Robustness upgrades."""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try: sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

from decimal import Decimal
import json
from pathlib import Path
import re
from unittest.mock import patch

import ai_client
import app

CHECKS = []

def check(name, condition):
    assert condition, name
    CHECKS.append(name)

def run_tests():
    # 1. PII Redaction: CCCD 12-digit & CMND 9-digit
    prompt_cccd = 'Tìm sinh viên có số CCCD 079201004567 và CMND 241567891 để lấy điểm thi'
    clean_cccd = ai_client.PrivacyGuard.anonymize_prompt(prompt_cccd)
    check('CCCD 12-digit redacted', '079201004567' not in clean_cccd and '[CCCD_AN_DANH]' in clean_cccd)
    check('CMND 9-digit redacted', '241567891' not in clean_cccd and '[CMND_AN_DANH]' in clean_cccd)

    # 2. PII Redaction: Vietnamese Phone Numbers
    prompt_phone = 'Đối chiếu SĐT 0901234567 và +84987654321 với danh mục'
    clean_phone = ai_client.PrivacyGuard.anonymize_prompt(prompt_phone)
    check('Mobile 090 redacted', '0901234567' not in clean_phone and '[SDT_AN_DANH]' in clean_phone)
    check('Mobile +84987 redacted', '+84987654321' not in clean_phone)

    # 3. PII Redaction: Email Addresses
    prompt_email = 'Tìm học sinh có email nguyen.van.an@hcmut.edu.vn'
    clean_email = ai_client.PrivacyGuard.anonymize_prompt(prompt_email)
    check('Email address redacted', 'nguyen.van.an@hcmut.edu.vn' not in clean_email and '[EMAIL_AN_DANH]' in clean_email)

    # 4. PII Redaction: Bank Account & Card numbers
    prompt_bank = 'Lấy số tiền chuyển khoản vào STK 19034567890123 ngân hàng Techcombank'
    clean_bank = ai_client.PrivacyGuard.anonymize_prompt(prompt_bank)
    check('Bank account number redacted', '19034567890123' not in clean_bank and '[STK_AN_DANH]' in clean_bank)

    # 5. PII Redaction: Tax Code (MST)
    prompt_tax = 'Kiểm tra thuế công ty mã số thuế 0312345678-001'
    clean_tax = ai_client.PrivacyGuard.anonymize_prompt(prompt_tax)
    check('Tax code (MST) redacted', '0312345678-001' not in clean_tax and '[MST_AN_DANH]' in clean_tax)

    # 6. PII Redaction: Passwords & Secrets
    prompt_pass = 'Tìm mật khẩu: SecretPass123! và token Bearer eyJhbGciOiJIUzI1Ni...'
    clean_pass = ai_client.PrivacyGuard.anonymize_prompt(prompt_pass)
    check('Password redacted', 'SecretPass123!' not in clean_pass and '[MAT_KHAU_AN_DANH]' in clean_pass)
    check('Bearer token redacted', 'Bearer eyJhbGciOiJIUzI1Ni...' not in clean_pass and '[KEY_AN_DANH]' in clean_pass)

    # 7. PII Redaction: Personal names following Vietnamese titles
    prompt_name = 'Lấy học phí của sinh viên Nguyễn Văn An và thầy Trần Quốc Toản'
    clean_name = ai_client.PrivacyGuard.anonymize_prompt(prompt_name)
    check('Student name redacted', 'Nguyễn Văn An' not in clean_name and 'sinh viên [TEN_AN_DANH]' in clean_name)
    check('Teacher name redacted', 'Trần Quốc Toản' not in clean_name and 'thầy [TEN_AN_DANH]' in clean_name)

    # 8. Column Header PII Anonymization & Deanonymization
    raw_headers = ['STT', 'Họ và tên', '079201004567', 'an.nguyen@edu.vn', 'Số tiền']
    clean_cols, mapping = ai_client.PrivacyGuard.anonymize_columns(raw_headers)
    check('CCCD column header anonymized', '079201004567' not in clean_cols and '[COT_BAO_MAT_3]' in clean_cols)
    check('Email column header anonymized', 'an.nguyen@edu.vn' not in clean_cols and '[COT_BAO_MAT_4]' in clean_cols)
    check('Safe column headers preserved', clean_cols[0] == 'STT' and clean_cols[1] == 'Họ và tên' and clean_cols[4] == 'Số tiền')

    # Deanonymize plan round-trip
    mock_plan = {'action': 'lookup', 'key': '[COT_BAO_MAT_3]', 'ref_key': '[COT_BAO_MAT_3]', 'ref_value': '[COT_BAO_MAT_4]'}
    restored_plan = ai_client.PrivacyGuard.deanonymize_plan(mock_plan, mapping)
    check('Deanonymized key restored', restored_plan['key'] == '079201004567')
    check('Deanonymized ref_value restored', restored_plan['ref_value'] == 'an.nguyen@edu.vn')

    # 9. Secret masking for UI and logs
    check('Secret key masked', ai_client.PrivacyGuard.mask_secret('sk-1234567890abcdefghij') == 'sk-...ghij')
    check('Short secret key masked', ai_client.PrivacyGuard.mask_secret('short') == '***')
    check('Empty secret handled', ai_client.PrivacyGuard.mask_secret('') == '')

    # 10. Error message sanitization
    err_raw = 'Failed call: Authorization: Bearer sk-1234567890abcdefghij to https://api.openai.com'
    err_clean = ai_client.PrivacyGuard.sanitize_error(err_raw)
    check('Bearer token scrubbed from error', 'Bearer sk-1234567890abcdefghij' not in err_clean and '[MASKED]' in err_clean)

    # 11. Local Plan Cache: Instant proposal with Zero Egress
    source_cols = ['MSSV', 'Họ và tên', 'Số tiền']
    plan_first = {'action': 'summary', 'key': 'Họ và tên', 'value': 'Số tiền', 'ref_key': '', 'ref_value': '', 'message': ''}
    with patch.object(ai_client, 'call_openai_plan', return_value=plan_first) as mock_call:
        res1 = ai_client.request_plan('Cộng Số tiền theo Họ và tên', source_cols, provider='openai', api_key='sk-test')
        check('First call executed API', mock_call.call_count == 1)
        res2 = ai_client.request_plan('Cộng Số tiền theo Họ và tên', source_cols, provider='openai', api_key='sk-test')
        check('Second call served from cache (Zero Egress)', mock_call.call_count == 1 and res2 == res1)

    # 12. Accounting Parentheses & Trailing Minus Negative Numbers in parse_lenient_number
    check('Accounting parenthesis negative (1,234.56)', app.parse_lenient_number('(1,234.56)') == Decimal('-1234.56'))
    check('Accounting parenthesis with currency (1.500.000 đ)', app.parse_lenient_number('(1.500.000 đ)') == Decimal('-1500000'))
    check('Trailing minus sign 1,000-', app.parse_lenient_number('1,000-') == Decimal('-1000'))
    check('Leading minus with space - 25.5', app.parse_lenient_number('- 25.5') == Decimal('-25.5'))
    check('Leading plus sign +100', app.parse_lenient_number('+100') == Decimal('100'))

    # 13. Date String Normalization in key_value (allow_dates=True)
    check('VN date DD/MM/YYYY normalized', app.key_value('15/05/2024', allow_dates=True) == '2024-05-15')
    check('VN date DD-MM-YYYY normalized', app.key_value('15-05-2024', allow_dates=True) == '2024-05-15')
    check('ISO date YYYY-MM-DD normalized', app.key_value('2024-05-15', allow_dates=True) == '2024-05-15')
    check('Standard VN date 01/02/2023 normalized', app.key_value('01/02/2023', allow_dates=True) == '2023-02-01')

    # 14. Zero-Width Space & Invisible Unicode Cleanup
    dirty_text = '\u200b\u00a0 Nguyễn Văn A \ufeff\u00a0'
    dummy_source = {'headers': ['Tên'], 'data': [[dirty_text]], 'metadata': {'sheets': {'Data': {'formats': {}}}}, 'sheet': 'Data', 'header_row': 1, 'hash': 'test'}
    sheets, name, stats = app.prepare_sheets(dummy_source, None, {'action': 'clean', 'key': 'Tên'}, 'test')
    res_sheet = next(s for s in sheets if s['name'] == name)
    cleaned_cell = res_sheet['rows'][1][0]
    check('Zero-width and non-breaking spaces cleaned', cleaned_cell == 'Nguyễn Văn A' and '\u200b' not in cleaned_cell and '\ufeff' not in cleaned_cell)

    # 15. .gitignore safety verification
    gitignore = (app.ROOT / '.gitignore')
    check('.gitignore exists', gitignore.exists())
    gi_content = gitignore.read_text(encoding='utf-8')
    check('.gitignore protects config.json', 'config.json' in gi_content)
    check('.gitignore protects .work directory', '.work/' in gi_content)
    check('.gitignore protects output directory', 'output/' in gi_content)

    # 16. Windows DPAPI protection round-trip
    test_key = 'sk-proj-SuperSecretKey1234567890'
    enc_key = ai_client.dpapi_protect(test_key)
    check('DPAPI ciphertext formatted', enc_key.startswith('dpapi:') and len(enc_key) > 50)
    check('DPAPI plaintext recovered', ai_client.dpapi_unprotect(enc_key) == test_key)

    # 17. Safe Config Storage: No plaintext API keys on disk
    with patch.object(ai_client, 'CONFIG_FILE', app.WORK / 'test-config.json'):
        ai_client.save_config({'openai_api_key': 'sk-test-secret-1234'})
        disk_content = (app.WORK / 'test-config.json').read_text(encoding='utf-8')
        check('Plaintext API key NOT on disk', 'sk-test-secret-1234' not in disk_content)
        check('Encrypted key stored on disk', 'openai_api_key_enc' in disk_content and 'dpapi:' in disk_content)
        read_back = ai_client.get_config()
        check('API key recovered seamlessly in memory', read_back['openai_api_key'] == 'sk-test-secret-1234')

    # 18. Stale work directories garbage collection
    fake_stale = app.WORK / 'inspect-stale-test'
    fake_stale.mkdir(exist_ok=True)
    check('Fake stale dir exists before cleanup', fake_stale.exists())
    app.cleanup_stale_work_dirs(max_age_seconds=0)
    check('Fake stale dir cleaned up', not fake_stale.exists())

    print(f'All {len(CHECKS)} security, privacy, and robustness checks passed successfully!')
    return CHECKS

if __name__ == '__main__':
    run_tests()
