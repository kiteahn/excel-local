"""Mutation check for the Q6 refactor.

Each mutation must be caught by the existing golden suite. The file is restored
byte-for-byte in `finally`, so line endings cannot be damaged by this script.
"""
from pathlib import Path
import os
import shutil
import subprocess
import sys

path = Path('ai_client.py')
original = path.read_bytes()
text = original.decode('utf-8')
canonical = text.replace('\r\n', '\n')
uses_crlf = '\r\n' in text
mutations = [
    ('M1 registration team value', "                'value': 'Đội',",
     "                'value': 'ĐộiX',"),
    ('M2 numeric value spacing', 'filter_val = f"{op} {num_clean}"',
     'filter_val = f"{op}{num_clean}"'),
    ('M3 officer role list', "filter_val = 'lớp trưởng, tổ trưởng, bí thư, phó bí thư, lớp phó, thủ quỹ, cán sự'",
     "filter_val = 'lớp trưởng, tổ trưởng, bí thư, phó bí thư, lớp phó, thủ quỹ, cán sự, x'"),
    ('M4 registration value', "                'value': 'Đơn',",
     "                'value': 'ĐơnX',"),
    ('M5 registration key', "                'key': reg_col,\n                'value': 'Đơn',",
     "                'key': source_columns[0],\n                'value': 'Đơn',"),
    ('M6 department value', 'filter_col, filter_val = dept_col, dept_val',
     "filter_col, filter_val = dept_col, dept_val + 'X'"),
]
results = []
env = dict(os.environ, PYTHONDONTWRITEBYTECODE='1')
try:
    for label, old, new in mutations:
        count = canonical.count(old)
        if count != 1:
            results.append((label, 'SKIP', f'occurrences={count}'))
            continue
        mutated = canonical.replace(old, new)
        if uses_crlf:
            mutated = mutated.replace('\n', '\r\n')
        path.write_bytes(mutated.encode('utf-8'))
        shutil.rmtree(Path('__pycache__'), ignore_errors=True)
        proc = subprocess.run(
            [sys.executable, '-X', 'utf8', 'test_infer_golden.py'],
            capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)
        caught = proc.returncode != 0
        output = (proc.stderr or '').strip() or (proc.stdout or '').strip()
        detail = output.splitlines()[-1][:120] if output else f'exit={proc.returncode}'
        results.append((label, 'CAUGHT' if caught else 'MISSED', detail))
finally:
    path.write_bytes(original)
    shutil.rmtree(Path('__pycache__'), ignore_errors=True)

for label, state, detail in results:
    print(f'[{state}] {label} :: {detail}')
missed = [row for row in results if row[1] != 'CAUGHT']
print(f'{len(results) - len(missed)}/{len(results)} mutations caught')
sys.exit(1 if missed else 0)
