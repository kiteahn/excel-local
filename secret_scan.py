"""Quét bí mật trong cây làm việc và toàn bộ lịch sử git (không gọi mạng).

Chỉ đọc file đang được git theo dõi và `git log -p --all`; không đọc file
ngoài repo, không gửi dữ liệu đi đâu. Báo cáo ghi vào results/secret-scan.json.

python -X utf8 secret_scan.py
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / 'results' / 'secret-scan.json'
ALLOWLIST = ROOT / 'secret_scan_allowlist.json'

PATTERNS = [
    ('openai_key', re.compile(r'\bsk-[A-Za-z0-9_-]{20,}')),
    ('github_token', re.compile(r'\b(?:ghp|gho|ghu|ghs|ghr)_[A-Za-z0-9]{20,}|\bgithub_pat_[A-Za-z0-9_]{20,}')),
    ('aws_key', re.compile(r'\bAKIA[0-9A-Z]{16}\b')),
    ('google_key', re.compile(r'\bAIza[0-9A-Za-z_-]{35}\b')),
    ('slack_token', re.compile(r'\bxox[baprs]-[A-Za-z0-9-]{10,}')),
    ('groq_key', re.compile(r'\bgsk_[A-Za-z0-9]{20,}')),
    ('personal_path', re.compile(r"[A-Za-z]:\\Users\\[^\\\s\"'<>]+", re.IGNORECASE)),
    ('private_key', re.compile(r'-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----')),
    ('assigned_secret', re.compile(
        r'(?i)\b(?:api[_-]?key|secret|token|password|passwd|pwd)\b\s*[:=]\s*[\'"]([^\'"\s]{8,})[\'"]')),
]

PLACEHOLDER_HINTS = (
    'test', 'demo', 'example', 'sample', 'dummy', 'fake', 'your', 'xxx', 'redacted',
    'changeme', 'placeholder', 'replace', 'insert', 'local', 'none', 'null', 'viet',
)


def mask(value):
    value = str(value)
    if len(value) <= 4:
        return '***'
    return value[:3] + '***' + value[-1:]


def is_placeholder(value):
    lowered = str(value).lower()
    if '...' in lowered or '<' in lowered or '>' in lowered:
        return True
    if any(hint in lowered for hint in PLACEHOLDER_HINTS):
        return True
    return len(set(lowered)) <= 2


def scan_line(line):
    hits = []
    for name, pattern in PATTERNS:
        for match in pattern.finditer(line):
            candidate = match.group(1) if match.lastindex else match.group(0)
            if is_placeholder(candidate):
                continue
            hits.append({'pattern': name, 'sample': mask(candidate)})
    return hits


def tracked_files():
    completed = subprocess.run(['git', 'ls-files', '-z'], cwd=ROOT, capture_output=True, check=True)
    return [name for name in completed.stdout.decode('utf-8', 'replace').split('\0') if name]


def load_allowlist():
    if not ALLOWLIST.exists():
        return set()
    data = json.loads(ALLOWLIST.read_text(encoding='utf-8'))
    return {(entry['file'], entry['pattern'], entry['sample']) for entry in data.get('allowed', [])}


def scan_working_tree(paths):
    findings = []
    scanned = 0
    for name in paths:
        path = ROOT / name
        if not path.is_file():
            continue
        raw = path.read_bytes()
        if b'\0' in raw[:4096]:
            continue
        scanned += 1
        for number, line in enumerate(raw.decode('utf-8', 'replace').splitlines(), 1):
            for hit in scan_line(line):
                findings.append({'file': name, 'line': number, **hit})
    return scanned, findings


def scan_history():
    completed = subprocess.run(['git', 'log', '-p', '--all', '--no-color', '--pretty=format:commit %H %s'],
                               cwd=ROOT, capture_output=True, check=True)
    text = completed.stdout.decode('utf-8', 'replace')
    findings = []
    commit = None
    file_name = None
    for line in text.splitlines():
        if line.startswith('commit '):
            commit = line[7:].split(' ', 1)[0]
            continue
        if line.startswith('+++ b/'):
            file_name = line[6:]
            continue
        if not line.startswith('+') or line.startswith('+++'):
            continue
        for hit in scan_line(line[1:]):
            findings.append({'commit': commit, 'file': file_name or '?', **hit})
    return findings


def main():
    allowed = load_allowlist()
    paths = tracked_files()
    scanned, work_findings = scan_working_tree(paths)
    history_findings = scan_history()
    work_allowed = [f for f in work_findings if (f['file'], f['pattern'], f['sample']) in allowed]
    history_allowed = [f for f in history_findings if (f['file'], f['pattern'], f['sample']) in allowed]
    work_findings = [f for f in work_findings if (f['file'], f['pattern'], f['sample']) not in allowed]
    history_findings = [f for f in history_findings if (f['file'], f['pattern'], f['sample']) not in allowed]
    report = {
        'mode': 'tracked files + git log -p --all; no network',
        'tracked_files': len(paths),
        'text_files_scanned': scanned,
        'working_tree_findings': work_findings,
        'history_findings': history_findings,
        'allowlisted_synthetic_fixtures': {'working_tree': len(work_allowed), 'history': len(history_allowed),
                                           'source': ALLOWLIST.name},
        'note': 'Chỉ bắt mẫu bí mật phổ biến; không phải chứng nhận không còn bí mật nào.',
    }
    OUTPUT.parent.mkdir(exist_ok=True)
    OUTPUT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({key: value for key, value in report.items() if not key.endswith('findings')}, ensure_ascii=False, indent=2))
    print('working tree findings (not allowlisted):', len(work_findings))
    print('history findings (not allowlisted):', len(history_findings))
    for finding in work_findings + history_findings:
        print(' -', finding)
    if work_findings or history_findings:
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
