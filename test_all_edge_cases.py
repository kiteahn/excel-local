"""Edge-case stress test and comprehensive verification across all boundaries."""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try: sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

import copy
from decimal import Decimal
import io
import json
from pathlib import Path
import tempfile
import unicodedata
from unittest.mock import patch

import ai_client
import app
from test_app import fixture

PASSES = []
FAILS = []

def test(name, condition, error_detail=None):
    if condition:
        PASSES.append(name)
    else:
        FAILS.append((name, error_detail))
        print(f"FAILED: {name} - {error_detail}")

def run_all_edge_cases():
    print("Starting Comprehensive Edge-Case Stress Testing...")

    # --- DOMAIN 1: File I/O & Malformed Files ---
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        dir_path = Path(td)
        
        # 1.1 Empty 0-byte XLSX
        empty_xlsx = dir_path / "empty.xlsx"
        empty_xlsx.write_bytes(b"")
        try:
            app.inspect_file(empty_xlsx)
            test("0-byte XLSX rejected", False, "Accepted empty file")
        except Exception as e:
            test("0-byte XLSX rejected", True)

        # 1.2 Empty 0-byte CSV
        empty_csv = dir_path / "empty.csv"
        empty_csv.write_bytes(b"")
        try:
            app.inspect_file(empty_csv)
            test("0-byte CSV rejected", False, "Accepted empty CSV")
        except Exception as e:
            test("0-byte CSV rejected", "CSV rỗng" in str(e) or "rỗng" in str(e).lower())

        # 1.3 Corrupted random bytes with .xlsx extension
        corrupt_xlsx = dir_path / "corrupt.xlsx"
        corrupt_xlsx.write_bytes(b"PK\x03\x04randomcorruptedbytes1234567890")
        try:
            app.inspect_file(corrupt_xlsx)
            test("Corrupted XLSX rejected", False, "Accepted corrupted file")
        except Exception as e:
            test("Corrupted XLSX rejected", True)

        # 1.4 Non-existent file
        missing_file = dir_path / "does_not_exist.xlsx"
        try:
            app.inspect_file(missing_file)
            test("Missing file rejected", False, "Accepted missing file")
        except (FileNotFoundError, OSError, ValueError):
            test("Missing file rejected", True)

        # 1.5 Unsupported extension (.txt, .xlsm)
        txt_file = dir_path / "data.txt"
        txt_file.write_text("a,b,c", encoding="utf-8")
        try:
            app.inspect_file(txt_file)
            test("Unsupported .txt extension rejected", False, "Accepted .txt")
        except ValueError as e:
            test("Unsupported .txt extension rejected", "V1 chỉ nhận XLSX, CSV hoặc XLS" in str(e))

        # 1.6 CSV with escaped quotes, semicolons, and CRLF
        tricky_csv = dir_path / "tricky.csv"
        tricky_csv.write_bytes('STT;Họ "tên";Ghi chú\r\n1;"Nguyễn ""An""";"Dòng 1"\r\n2;"Trần Bình";"Dòng 2"\r\n'.encode('utf-8'))
        try:
            snap = app.inspect_file(tricky_csv)
            test("Semicolon CSV with escaped quotes parsed", snap['headers'] == ['STT', 'Họ "tên"', 'Ghi chú'] and len(snap['data']) == 2)
        except Exception as e:
            test("Semicolon CSV with escaped quotes parsed", False, str(e))

        # 1.7 CSV with UTF-8 BOM
        bom_csv = dir_path / "bom.csv"
        bom_csv.write_bytes('\ufeffMã,Tên\r\n001,An\r\n'.encode('utf-8'))
        try:
            snap_bom = app.inspect_file(bom_csv)
            test("UTF-8 BOM CSV parsed cleanly", snap_bom['headers'] == ['Mã', 'Tên'] and snap_bom['data'][0] == ['001', 'An'])
        except Exception as e:
            test("UTF-8 BOM CSV parsed cleanly", False, str(e))

    # --- DOMAIN 2: Header Auto-Healing & Margins ---
    # 2.1 Table with all numeric row 1 (auto-heal)
    raw_numeric = {
        'sheet': 'Data',
        'rows': [
            [2023, 2024, 2025],
            [10, 20, 30]
        ]
    }
    healed_num = app.select_header(raw_numeric, 1, auto_heal=True)
    test("All-numeric header row converted to strings", healed_num['headers'] == ['2023', '2024', '2025'])

    # 2.2 Table with 3 identical column names
    raw_triplicate = {
        'sheet': 'Data',
        'rows': [
            ['Điểm', 'Điểm', 'Điểm'],
            [8, 9, 10]
        ]
    }
    healed_trip = app.select_header(raw_triplicate, 1, auto_heal=True)
    test("Triplicate headers disambiguated", healed_trip['headers'] == ['Điểm', 'Điểm_2', 'Điểm_3'])

    # 2.3 Table with consecutive empty headers
    raw_empty_gaps = {
        'sheet': 'Data',
        'rows': [
            ['Mã', None, None, 'Điểm'],
            ['001', 'A', 'B', 10]
        ]
    }
    healed_gaps = app.select_header(raw_empty_gaps, 1, auto_heal=True)
    test("Consecutive empty headers disambiguated", healed_gaps['headers'] == ['Mã', 'Mã_phụ', 'Mã_phụ_2', 'Điểm'])

    # 2.4 Headers with trailing whitespace
    raw_ws_headers = {
        'sheet': 'Data',
        'rows': [
            ['  Mã SV  ', '  Họ và tên  '],
            ['SV01', 'Nguyễn An']
        ]
    }
    healed_ws = app.select_header(raw_ws_headers, 1, auto_heal=True)
    test("Whitespace stripped from headers", healed_ws['headers'] == ['Mã SV', 'Họ và tên'])

    # 2.5 Out of bounds header row rejected
    try:
        app.select_header(raw_ws_headers, 999, auto_heal=True)
        test("Out-of-bounds header row rejected", False, "Accepted row 999")
    except app.HeaderValidationError:
        test("Out-of-bounds header row rejected", True)

    # --- DOMAIN 3: Number & Currency Parsing Stress Tests ---
    # 3.1 Various accounting & negative notations
    test("Accounting negative (1,234.56)", app.parse_lenient_number("(1,234.56)") == Decimal("-1234.56"))
    test("Accounting currency (1.500.000 đ)", app.parse_lenient_number("(1.500.000 đ)") == Decimal("-1500000"))
    test("Accounting with dollar ($ 2,500.00)", app.parse_lenient_number("($ 2,500.00)") == Decimal("-2500.00"))
    test("Trailing minus 999-", app.parse_lenient_number("999-") == Decimal("-999"))
    test("Leading minus with spaces - 123.4", app.parse_lenient_number("- 123.4") == Decimal("-123.4"))
    test("Leading plus +500", app.parse_lenient_number("+500") == Decimal("500"))

    # 3.2 Euro, Yen, VND variations
    test("Euro currency 1.234,56 €", app.parse_lenient_number("1.234,56 €") == Decimal("1234.56"))
    test("Yen currency ¥50,000", app.parse_lenient_number("¥50,000") == Decimal("50000"))
    test("VND lowercase vnđ 2.000.000 vnđ", app.parse_lenient_number("2.000.000 vnđ") == Decimal("2000000"))

    # 3.3 Zero variations
    test("Zero string '0'", app.parse_lenient_number("0") == Decimal("0"))
    test("Zero with decimals '0.00'", app.parse_lenient_number("0.00") == Decimal("0.00"))
    test("Zero in parens '(0)'", app.parse_lenient_number("(0)") == Decimal("0"))

    # 3.4 Invalid numbers reject gracefully with ValueError
    for bad in ["abc", "--5", "((", "1.2.3.4.5", "10^20", ""]:
        try:
            app.parse_lenient_number(bad, allow_blank_zero=False)
            test(f"Invalid number '{bad}' rejected", False, f"Accepted '{bad}'")
        except ValueError:
            test(f"Invalid number '{bad}' rejected", True)

    # 3.5 Extremely large numbers > 10^14 rejected
    try:
        app.parse_lenient_number("999999999999999999999")
        test("Oversized number rejected", False, "Accepted > 10^14")
    except ValueError as e:
        test("Oversized number rejected", "10^14" in str(e))

    # --- DOMAIN 4: Date & Key Normalization Edge Cases ---
    # 4.1 Leap year date 29/02/2024
    test("Leap year 29/02/2024 normalized", app.key_value("29/02/2024", allow_dates=True) == "2024-02-29")
    test("Year-end 31/12/2024 normalized", app.key_value("31/12/2024", allow_dates=True) == "2024-12-31")

    # 4.2 Invalid date 31/02/2024 gracefully falls back to text without crashing
    test("Invalid date 31/02/2024 handled without crash", app.key_value("31/02/2024", allow_dates=True) == "31/02/2024")

    # 4.3 Serial date with Excel format
    test("Excel serial date 45351 (2024-02-29)", app.key_value(45351, "dd/mm/yyyy", allow_dates=True) == "2024-02-29")

    # 4.4 Zero-padded code preservation
    test("Zero-padded code preserved", app.key_value(123, "00000") == "00123")
    test("Raw string code with zero preserved", app.key_value("00123") == "00123")

    # 4.5 NFC vs NFD Unicode normalization equivalence
    nfc_name = unicodedata.normalize('NFC', 'Nguyễn Văn An')
    nfd_name = unicodedata.normalize('NFD', 'Nguyễn Văn An')
    test("NFC and NFD strings produce identical key", app.key_value(nfc_name) == app.key_value(nfd_name))

    # 4.6 Excel wildcard escaping in escape_criteria
    escaped_star = app.escape_criteria('A1')
    test("Criteria has SUBSTITUTE for ~, *, ?", '~~' in escaped_star and '~*' in escaped_star and '~?' in escaped_star)

    # --- DOMAIN 5: Clean Action & Zero-Width Characters ---
    # 5.1 Multiple zero-width characters stripped
    zw_text = "\u200b\u200c\u200d\ufeff\u200e\u200f\u00a0 Nguyễn Văn An \u200b\u00a0"
    zw_source = {
        'headers': ['Tên'],
        'data': [[zw_text]],
        'metadata': {'sheets': {'Data': {'formats': {}}}},
        'sheet': 'Data',
        'header_row': 1,
        'hash': 'test'
    }
    sheets, name, stats = app.prepare_sheets(zw_source, None, {'action': 'clean', 'key': 'Tên'}, 'zw_test')
    res_sheet = next(s for s in sheets if s['name'] == name)
    cleaned_val = res_sheet['rows'][1][0]
    test("Zero-width & non-breaking spaces cleaned cleanly", cleaned_val == "Nguyễn Văn An" and len(cleaned_val) == 13)
    test("Changed cells counter incremented", stats['changed_cells'] == 1)

    # 5.2 Numeric cell in clean action preserved untouched
    num_source = {
        'headers': ['Số'],
        'data': [[12345], [67.89]],
        'metadata': {'sheets': {'Data': {'formats': {}}}},
        'sheet': 'Data',
        'header_row': 1,
        'hash': 'test'
    }
    num_sheets, num_name, num_stats = app.prepare_sheets(num_source, None, {'action': 'clean', 'key': 'Số'}, 'num_test')
    num_res = next(s for s in num_sheets if s['name'] == num_name)
    test("Numeric cells in clean action preserved", num_res['rows'][1][0] == 12345 and num_res['rows'][2][0] == 67.89)

    # --- DOMAIN 6: Privacy Guard & AI Delimiter Security ---
    # 6.1 Multi-PII prompt
    mega_prompt = (
        "Sinh viên Nguyễn Văn An số CCCD 079201004567, CMND 241567891, "
        "SĐT 090.123.4567 và (+84) 987 654 321, email an.nguyen@edu.vn, "
        "STK 19034567890123 ngân hàng, MST 0312345678, mật khẩu: TestPass@123, "
        "thầy Trần Văn Bình hãy đối chiếu điểm thi."
    )
    clean_mega = ai_client.PrivacyGuard.anonymize_prompt(mega_prompt)
    test("Mega-prompt: CCCD removed", "079201004567" not in clean_mega)
    test("Mega-prompt: CMND removed", "241567891" not in clean_mega)
    test("Mega-prompt: SĐT 090 removed", "090" not in clean_mega or "090.123.4567" not in clean_mega)
    test("Mega-prompt: Email removed", "an.nguyen@edu.vn" not in clean_mega)
    test("Mega-prompt: STK removed", "19034567890123" not in clean_mega)
    test("Mega-prompt: MST removed", "0312345678" not in clean_mega)
    test("Mega-prompt: Password removed", "TestPass@123" not in clean_mega)
    test("Mega-prompt: Student name removed", "Nguyễn Văn An" not in clean_mega)
    test("Mega-prompt: Teacher name removed", "Trần Văn Bình" not in clean_mega)

    # 6.2 Normal query with no PII untouched
    safe_prompt = "Cộng tổng Số tiền theo từng Phòng ban trong quý 3."
    test("Safe prompt untouched", ai_client.PrivacyGuard.anonymize_prompt(safe_prompt) == safe_prompt)

    # 6.3 Prompt Injection rejection by validate_plan
    malicious_plan = {'action': 'eval', 'key': 'os.system("calc")', 'value': '', 'ref_key': '', 'ref_value': '', 'message': ''}
    try:
        app.validate_plan(malicious_plan, {'headers': ['A', 'B']})
        test("Malicious action='eval' rejected", False, "Accepted evil action")
    except ValueError as e:
        test("Malicious action='eval' rejected", "Thao tác ngoài danh sách được phép" in str(e))

    # 6.4 Missing column rejection by validate_plan
    non_existent_col_plan = {'action': 'summary', 'key': 'CotKhongTonTai', 'value': 'A', 'ref_key': '', 'ref_value': '', 'message': ''}
    try:
        app.validate_plan(non_existent_col_plan, {'headers': ['A', 'B']})
        test("Non-existent column in plan rejected", False, "Accepted missing column")
    except ValueError as e:
        test("Non-existent column in plan rejected", "không tồn tại trong nguồn" in str(e))

    # 6.5 Fuzzy match resolution
    fuzzy_plan = {'action': 'lookup', 'key': 'Mã sv', 'ref_key': 'Mã sv', 'ref_value': 'Họ tên', 'message': ''}
    src_cols = {'headers': ['Mã sinh viên', 'Điểm']}
    ref_cols = {'headers': ['Mã sinh viên', 'Họ và tên']}
    val_plan = app.validate_plan(fuzzy_plan, src_cols, ref_cols)
    test("Fuzzy match mapped 'Mã sv' to 'Mã sinh viên'", val_plan['key'] == 'Mã sinh viên')
    test("Fuzzy match mapped 'Họ tên' to 'Họ và tên'", val_plan['ref_value'] == 'Họ và tên')

    # --- DOMAIN 7: Windows DPAPI & Encryption Edge Cases ---
    # 7.1 Empty string DPAPI handling
    test("DPAPI empty string handles cleanly", ai_client.dpapi_protect("") == "")
    test("DPAPI unprotect empty string handles cleanly", ai_client.dpapi_unprotect("") == "")
    test("DPAPI unprotect unencrypted string returns as-is", ai_client.dpapi_unprotect("sk-plain") == "sk-plain")

    # 7.2 Unicode and long API key encryption
    long_secret = "sk-proj-SpecialKeyWithUnicode_TiếngViệt_1234567890_abcdefghij"
    enc_long = ai_client.dpapi_protect(long_secret)
    dec_long = ai_client.dpapi_unprotect(enc_long)
    test("DPAPI handles complex unicode key round-trip", dec_long == long_secret)

    # --- DOMAIN 8: Stale Work Dir Cleanup Boundary ---
    # 8.1 Active file vs Stale directory
    with tempfile.TemporaryDirectory(dir=app.WORK) as current_td:
        app.cleanup_stale_work_dirs(max_age_seconds=7200) # 2 hours
        test("Active current directory not deleted", Path(current_td).exists())

    print(f"\nCompleted! Total Checks: {len(PASSES) + len(FAILS)}")
    print(f"PASSED: {len(PASSES)}")
    print(f"FAILED: {len(FAILS)}")
    if FAILS:
        for f, err in FAILS:
            print(f"  - {f}: {err}")
    return len(FAILS) == 0

if __name__ == '__main__':
    ok = run_all_edge_cases()
    sys.exit(0 if ok else 1)
