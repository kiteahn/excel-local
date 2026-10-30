"""Policy decisions of 2026-10-30, pinned as tests.

(1) Default engine is the local Ollama only; cloud modes are opt-in.
    Existing stored configs keep whatever provider the user already chose.
(2) In lenient number parsing, '500.000' (single dot + exactly 3 digits)
    reads as 500000 when the 'vn_dot_thousands' setting is on (default),
    and as decimal 500 when off. '0.500' is always decimal.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import json
import tempfile
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import ai_client
import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def with_config(cfg_dict, fn):
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        cfg = Path(td) / 'config.json'
        if cfg_dict is not None:
            cfg.write_text(json.dumps(cfg_dict), encoding='utf-8')
        with patch.object(ai_client, 'CONFIG_FILE', cfg):
            return fn()

def run():
    ok('P-01 mac dinh khi chua co config: provider=ollama',
       with_config(None, lambda: ai_client.get_config()['provider']) == 'ollama')
    ok('P-02 mac dinh vn_dot_thousands=True',
       with_config(None, lambda: ai_client.get_config()['vn_dot_thousands']) is True)
    ok('P-03 config da luu cua nguoi dung duoc ton trong (smart giu nguyen)',
       with_config({'provider': 'smart'}, lambda: ai_client.get_config()['provider']) == 'smart')

    ok('P-04 500.000 voi dot_thousands=True -> 500000',
       app.parse_lenient_number('500.000', dot_thousands=True) == Decimal('500000'))
    ok('P-05 500.000 voi dot_thousands=False -> 500',
       app.parse_lenient_number('500.000', dot_thousands=False) == Decimal('500'))
    ok('P-06 12.345 voi True -> 12345', app.parse_lenient_number('12.345', dot_thousands=True) == Decimal('12345'))
    ok('P-07 0.500 luon la thap phan 0.5 ke ca khi True',
       app.parse_lenient_number('0.500', dot_thousands=True) == Decimal('0.5'))
    ok('P-08 1.2345 (4 chu so sau cham) van la thap phan',
       app.parse_lenient_number('1.2345', dot_thousands=True) == Decimal('1.2345'))
    ok('P-09 -500.000 voi True -> -500000',
       app.parse_lenient_number('-500.000', dot_thousands=True) == Decimal('-500000'))
    ok('P-10 da dau cham 1.234.567 khong doi: 1234567',
       app.parse_lenient_number('1.234.567', dot_thousands=False) == Decimal('1234567'))

    def prepare_total(dot_setting):
        def _run():
            snapshot = {
                'sheet': 'Data', 'sheets': ['Data'], 'path': 'mem', 'hash': 'x',
                'headers': ['Lop', 'Tien'],
                'data': [['A', '500.000'], ['A', '100']],
                'metadata': {'sheets': {'Data': {'formats': {}}}},
                'header_row': 1, 'column_offset': 0,
            }
            plan = {'action': 'summary', 'key': 'Lop', 'value': 'Tien', 'lenient_numbers': True}
            _, _, stats = app.prepare_sheets(snapshot, None, plan, 'pol')
            return stats['independent_total']
        return with_config({'vn_dot_thousands': dot_setting}, _run)
    ok('P-11 prepare summary theo config True: tong 500100', prepare_total(True) == '500100')
    ok('P-12 prepare summary theo config False: tong 600', Decimal(prepare_total(False)) == Decimal('600'))

    print(f'\nAll {len(CHECKS)} policy checks passed.')

if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    run()
