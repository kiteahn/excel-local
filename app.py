"""Excel Local Expert: fixed operations, independent checks, native Excel, dual-engine AI (OpenAI & Ollama).

Python standard library only. Run `python app.py` or use start.cmd.
"""
from __future__ import annotations

import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try: sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

import argparse
import collections
import csv
import datetime
import difflib
import hashlib
import io
import json
import math
import os
from pathlib import Path
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import time
import unicodedata
import uuid
import xml.etree.ElementTree as ET
import zipfile
from decimal import Decimal, InvalidOperation

import ai_client

ROOT = Path(__file__).resolve().parent
WORK = ROOT / '.work'
OUTPUT = ROOT / 'output'
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
SAFE_FUNCTIONS = set('SUM SUMIF SUMIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS AVERAGE AVERAGEIF AVERAGEIFS MIN MAX IF IFS IFNA IFERROR AND OR NOT TRUE FALSE VLOOKUP HLOOKUP LOOKUP XLOOKUP INDEX MATCH XMATCH FILTER SORT SORTBY UNIQUE LET ROW ROWS COLUMN COLUMNS ROUND ROUNDUP ROUNDDOWN ABS INT MOD SUMPRODUCT SUBTOTAL AGGREGATE TRIM CLEAN LOWER UPPER PROPER LEN LEFT RIGHT MID SUBSTITUTE REPLACE TEXT VALUE NUMBERVALUE EXACT CONCAT CONCATENATE TEXTJOIN FIND SEARCH DATE DAY MONTH YEAR TODAY NOW DATEDIF EOMONTH ISBLANK ISNUMBER ISTEXT ISERROR ISNA CHOOSE OFFSET INDIRECT TRANSPOSE SEQUENCE'.split())
SAFE_FUNCTIONS.discard('INDIRECT')
SAFE_FUNCTIONS.add('NA')
BUILTIN_FORMATS = {0:'General', 1:'0', 2:'0.00', 9:'0%', 10:'0.00%', 14:'mm-dd-yy', 15:'d-mmm-yy', 16:'d-mmm', 17:'mmm-yy', 20:'h:mm', 21:'h:mm:ss', 22:'m/d/yy h:mm', 49:'@'}


def cleanup_stale_work_dirs(max_age_seconds=3600):
    """Clean up stale temporary directories in .work older than max_age_seconds.
    Guarantees no lingering sensitive data in temporary directories.
    """
    if not WORK.exists():
        return
    now = time.time()
    for item in WORK.iterdir():
        if item.is_dir() and item.name.startswith(('inspect-', 'run-', 'cleanup-', 'detect')):
            try:
                mtime = item.stat().st_mtime
                if now - mtime > max_age_seconds:
                    shutil.rmtree(item, ignore_errors=True)
            except Exception:
                pass


def file_hash(path):
    with open(path, 'rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def publish_file(source, target):
    """Publish a complete file atomically on Windows; never replace an output."""
    source, target = Path(source), Path(target)
    pending = target.with_name(target.name + '.partial-' + uuid.uuid4().hex[:8])
    try:
        with pending.open('xb') as out, source.open('rb') as incoming:
            shutil.copyfileobj(incoming, out)
            out.flush()
            os.fsync(out.fileno())
        if target.exists():
            raise FileExistsError('Tệp đầu ra đã tồn tại.')
        os.rename(pending, target)
    finally:
        pending.unlink(missing_ok=True)


def column_name(index):
    name = ''
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def formula_preflight(formula, lenient=False):
    without_text = re.sub(r'"(?:[^"]|"")*"', '', formula)
    if not lenient and any(token in without_text for token in ('[', ']', '|', '\\')):
        raise ValueError('V1 không nhận tham chiếu workbook ngoài, DDE hoặc tham chiếu bảng có ngoặc vuông.')
    if lenient:
        if any(token in without_text for token in ('|', '\\')):
            raise ValueError('Công thức chứa ký tự không an toàn.')
        return
    masked = re.sub(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'', '', formula)
    names = re.findall(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(', masked)
    unsupported = {name.upper().removeprefix('_XLFN.').removeprefix('_XLWS.') for name in names} - SAFE_FUNCTIONS
    if unsupported:
        raise ValueError('Công thức chưa được hỗ trợ an toàn: ' + ', '.join(sorted(unsupported)))


def read_xml(archive, name):
    data = archive.read(name)
    markup = data.replace(b'\x00', b'').upper()
    if b'<!DOCTYPE' in markup or b'<!ENTITY' in markup:
        raise ValueError('XML có DTD/entity; không hỗ trợ.')
    return ET.fromstring(data)


def preflight(path, lenient=False):
    path = Path(path).resolve(strict=True)
    if path.suffix.lower() not in ('.xlsx', '.csv', '.xls'):
        raise ValueError('V1 chỉ nhận XLSX, CSV hoặc XLS; không nhận XLSM/XLSB.')
    if path.stat().st_size > 80 * 1024 * 1024:
        raise ValueError('File vượt giới hạn 80 MB của v1.')
    if path.suffix.lower() == '.csv':
        return {'sheets': {'CSV': {}}}
    if path.suffix.lower() == '.xls':
        return {'sheets': {}}
    with zipfile.ZipFile(path) as archive:
        entries = archive.infolist()
        if len(entries) > 10000 or sum(e.file_size for e in entries) > 256 * 1024 * 1024:
            raise ValueError('XLSX vượt giới hạn giải nén 256 MB / 10000 thành phần.')
        names = {e.filename for e in entries}
        if len(names) != len(entries):
            raise ValueError('XLSX có thành phần ZIP trùng tên.')
        for entry in entries:
            name = entry.filename.lower()
            if name.startswith('/') or '..' in name.split('/') or '\\' in name:
                raise ValueError('Đường dẫn ZIP không hợp lệ.')
            if entry.file_size > max(1024 * 1024, entry.compress_size * 1000):
                raise ValueError('Tỷ lệ nén XLSX vượt giới hạn.')
            if name.endswith('.xml'):
                raw_xml = archive.read(entry.filename)
                markup = raw_xml.replace(b'\x00', b'').upper()
                if b'<!DOCTYPE' in markup or b'<!ENTITY' in markup:
                    raise ValueError('Thành phần XML có DTD/entity; không hỗ trợ.')
            forbidden = ('vbaproject', 'embeddings/', 'activex/', 'connections.xml', 'querytables/', 'customui/') if lenient else ('vbaproject', 'externallinks/', 'embeddings/', 'activex/', 'connections.xml', 'querytables/', 'customui/')
            if any(token in name for token in forbidden):
                raise ValueError('File có macro, liên kết, kết nối hoặc đối tượng nhúng; v1 từ chối mở.')
            if name.endswith('.rels'):
                relationships = read_xml(archive, entry.filename)
                for relationship in relationships:
                    if relationship.get('TargetMode', '').lower() == 'external':
                        rel_type = relationship.get('Type', '')
                        if not rel_type.endswith('/hyperlink'):
                            if not (lenient and 'externallink' in rel_type.lower()):
                                raise ValueError('File có quan hệ bên ngoài (kể cả hyperlink); v1 từ chối mở.')
        types = ET.tostring(read_xml(archive, '[Content_Types].xml'), encoding='unicode').lower()
        if any(token in types for token in ('macroenabled', 'macrosheet', 'dialogsheet', 'vba', 'oleobject', 'activex')):
            raise ValueError('File giả đuôi XLSX chứa macro.')
        workbook = read_xml(archive, 'xl/workbook.xml')
        for defined in workbook.findall('s:definedNames/s:definedName', NS):
            if not defined.get('name', '').startswith('_xlnm.'):
                formula_preflight(defined.text or '', lenient=lenient)
        relationships = read_xml(archive, 'xl/_rels/workbook.xml.rels')
        for relationship in relationships:
            if relationship.get('Type', '').endswith(('/macrosheet', '/intlMacrosheet', '/dialogsheet')):
                raise ValueError('Workbook chứa sheet macro Excel 4.0 hoặc dialog sheet; không hỗ trợ.')
        targets = {r.get('Id'): r.get('Target') for r in relationships}
        style_codes = ['General']
        if 'xl/styles.xml' in names:
            styles = read_xml(archive, 'xl/styles.xml')
            custom = {int(e.get('numFmtId')): e.get('formatCode') for e in styles.findall('s:numFmts/s:numFmt', NS)}
            style_codes = [custom.get(int(e.get('numFmtId', 0)), BUILTIN_FORMATS.get(int(e.get('numFmtId', 0)), 'General')) for e in styles.findall('s:cellXfs/s:xf', NS)]
        sheets = {}
        for sheet in workbook.findall('s:sheets/s:sheet', NS):
            relation = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target = targets[relation]
            target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            xml = read_xml(archive, target)
            formats = {}
            for cell in xml.findall('.//s:sheetData/s:row/s:c', NS):
                formula = cell.find('s:f', NS)
                if formula is not None:
                    formula_preflight(formula.text or '', lenient=lenient)
                    if formula.get('t') in ('dataTable',):
                        raise ValueError('V1 chưa hỗ trợ công thức data table.')
                style = int(cell.get('s', 0))
                if style < len(style_codes) and style_codes[style] != 'General':
                    formats[cell.get('r')] = style_codes[style]
            sheets[sheet.get('name')] = {'formats': formats}
        return {'sheets': sheets}


def terminate_active_bridges():
    try:
        subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command',
            'Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like "*excel_bridge.ps1*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }'],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0), timeout=5)
    except Exception:
        pass
    try:
        for p in WORK.glob('*/excel-process.json'):
            try:
                rec = json.loads(p.read_text(encoding='utf-8-sig'))
                if rec.get('owned') and rec.get('pid'):
                    subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', f'Stop-Process -Id {int(rec["pid"])} -Force -ErrorAction SilentlyContinue'],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0), timeout=5)
            except Exception:
                pass
    except Exception:
        pass


def load_xlsx_fast(path, sheet_name=''):
    try:
        import openpyxl
    except ImportError:
        return None
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    except Exception:
        return None
    try:
        names = list(wb.sheetnames)
        if not names:
            return None
        target_name = sheet_name if sheet_name and sheet_name in names else names[0]
        sheet = wb[target_name]
        rows = []
        max_w = 0
        for r in sheet.iter_rows(values_only=True):
            row_list = list(r)
            for c_idx in range(len(row_list)):
                val = row_list[c_idx]
                if isinstance(val, (datetime.date, datetime.datetime)):
                    row_list[c_idx] = val.strftime('%Y-%m-%d %H:%M:%S') if isinstance(val, datetime.datetime) and (val.hour or val.minute or val.second) else val.strftime('%Y-%m-%d')
            if len(row_list) > max_w:
                max_w = len(row_list)
            rows.append(row_list)
        for i in range(len(rows)):
            if len(rows[i]) < max_w:
                rows[i] = rows[i] + [None] * (max_w - len(rows[i]))
        return {
            'ok': True,
            'sheets': names,
            'sheet': target_name,
            'rows': rows,
            'excel_version': '16.0'
        }
    except Exception:
        return None
    finally:
        try:
            wb.close()
        except Exception:
            pass


def bridge(job, directory, timeout=240):
    directory = Path(directory)
    request = directory / 'job.json'
    result = directory / 'result.json'
    job = dict(job, result_file=str(result))
    request.write_text(json.dumps(job, ensure_ascii=False, allow_nan=False), encoding='utf-8')
    env = dict(os.environ, TEMP=str(directory), TMP=str(directory))
    command = ['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', str(ROOT / 'excel_bridge.ps1'), '-JobFile', str(request)]
    options = dict(env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    try:
        completed = subprocess.run(command, timeout=timeout, **options)
    except subprocess.TimeoutExpired as exc:
        cleanup = subprocess.run(command + ['-CleanupOnly'], timeout=20, **options)
        raise RuntimeError('Excel quá thời gian; đã yêu cầu dọn đúng tiến trình của tác vụ.' if cleanup.returncode == 0 else 'Excel quá thời gian; dọn tiến trình chưa thành công. Kiểm tra .work.') from exc
    if not result.exists():
        subprocess.run(command + ['-CleanupOnly'], timeout=20, **options)
        raise RuntimeError('Không nhận được kết quả từ Excel. Kiểm tra Excel desktop và quyền chạy PowerShell.')
    response = json.loads(result.read_text(encoding='utf-8-sig'))
    if completed.returncode or not response.get('ok'):
        raise RuntimeError('Excel: ' + response.get('error', 'Thao tác thất bại'))
    return response


def load_csv(path, strict_ragged=True):
    raw = Path(path).read_bytes()
    text = None
    for enc in ('utf-8-sig', 'utf-16', 'utf-8', 'cp1258', 'cp1252', 'latin1'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError('CSV phải là UTF-8 hoặc UTF-16 có BOM; hãy chuyển encoding rõ ràng.')
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=',;\t')
    except csv.Error:
        dialect = csv.excel
    rows = []
    width = None
    try:
        for row in csv.reader(io.StringIO(text, newline=''), dialect, strict=True):
            if not strict_ragged and not any(c.strip() for c in row if isinstance(c, str)):
                continue
            if width is None:
                width = len(row)
            if len(row) != width:
                if strict_ragged:
                    raise ValueError('Các dòng CSV có số cột khác nhau; cần sửa cấu trúc trước.')
                else:
                    if len(row) < width:
                        row = row + [''] * (width - len(row))
                    else:
                        row = row[:width]
            if not width or width > 60 or (len(rows) + 1) * width > 400000 or len(rows) >= 20001:
                raise ValueError('CSV vượt giới hạn 20000 dòng dữ liệu / 60 cột / 400000 ô.')
            if any(len(value) > 32767 or re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', value) for value in row):
                raise ValueError('CSV có ô dài quá 32767 ký tự hoặc ký tự điều khiển không hợp lệ trong Excel.')
            rows.append(row)
    except csv.Error as exc:
        raise ValueError('CSV có ô quá lớn hoặc dấu nháy/cấu trúc không hợp lệ.') from exc
    if not rows:
        raise ValueError('CSV rỗng.')
    return rows


class HeaderValidationError(ValueError):
    def __init__(self, message, snapshot):
        super().__init__(message)
        self.snapshot = snapshot


def detect_header_row(rows, max_search=25):
    """Auto-detect the most likely header row among the top rows of a spreadsheet,
    specifically tailored for Vietnamese administrative & university documents with banners.
    """
    best_row = 1
    best_score = -999999
    limit = min(max_search, len(rows))

    BANNER_PHRASES = (
        'cộng hòa', 'độc lập', 'tự do', 'hạnh phúc', 'quốc hiệu', 'tiêu ngữ',
        'bộ giáo dục', 'sở giáo dục', 'ubnd', 'ủy ban nhân dân',
        'trường đại học', 'học viện', 'viện nghiên cứu', 'khoa công nghệ',
        'danh sách sinh viên', 'danh sách ban cán sự', 'danh sách lớp', 'danh sách học sinh',
        'bảng tổng hợp', 'bảng kê', 'báo cáo', 'công văn', 'thông báo',
        'ngày ... tháng', 'ngày   tháng'
    )

    HEADER_KEYWORDS = (
        'stt', 'số tt', 'tt', 'no', 'mã', 'mssv', 'mã sv', 'mã sinh viên', 'mã học sinh', 'mã hs',
        'họ và tên', 'họ tên', 'tên', 'họ', 'chức vụ', 'chức danh', 'vai trò', 'nhiệm vụ', 'vị trí',
        'lớp', 'lớp sinh hoạt', 'khoa', 'ngành', 'phòng', 'ban', 'đơn vị',
        'ngày sinh', 'năm sinh', 'nơi sinh', 'quê quán', 'địa chỉ', 'giới tính',
        'sđt', 'số điện thoại', 'điện thoại', 'email', 'cccd', 'cmnd',
        'điểm', 'điểm rèn luyện', 'học lực', 'kết quả', 'xếp loại',
        'số tiền', 'tiền', 'học phí', 'lương', 'thành tiền', 'số lượng', 'đơn giá',
        'trạng thái', 'tình trạng', 'ghi chú', 'kỳ học', 'học kỳ'
    )

    for r in range(limit):
        row = rows[r]
        if not row:
            continue
        cells = [str(c).strip() for c in row if c is not None and str(c).strip() != '']
        if len(cells) < 2:
            continue

        row_text_low = ' '.join(cells).lower()

        # Banner row penalty: skip banner rows
        banner_hits = sum(1 for b in BANNER_PHRASES if b in row_text_low)
        if banner_hits > 0:
            continue

        kw_hits = 0
        for cell_low in [c.lower() for c in cells]:
            if any(cell_low == kw or cell_low.startswith(kw + ' ') or cell_low.endswith(' ' + kw) or f' {kw} ' in f' {cell_low} ' for kw in HEADER_KEYWORDS):
                kw_hits += 1

        distinct_ratio = len(set(cells)) / len(cells)
        has_data_below = (r + 1 < len(rows)) and sum(1 for c in rows[r + 1] if c is not None and str(c).strip() != '') >= 2

        score = (len(cells) * 3) + (kw_hits * 15) + int(distinct_ratio * 10) + (25 if has_data_below else 0)

        numeric_count = sum(1 for c in cells if c.replace('.', '', 1).replace('-', '', 1).isdigit())
        if numeric_count > len(cells) // 2:
            score -= 50

        if score > best_score:
            best_score = score
            best_row = r + 1

    return best_row


def select_header(snapshot, header_row, auto_heal=False):
    """Select a table row explicitly; optionally auto-heal merged or missing columns."""
    rows = snapshot['rows']
    if not 1 <= header_row <= len(rows):
        raise HeaderValidationError(f'Dòng tiêu đề {header_row} không tồn tại. Chọn từ 1 đến {len(rows)} rồi bấm Đọc.', snapshot)
    table = rows[header_row - 1:]
    width = len(table[0])
    occupied = [any(row[c] is not None and row[c] != '' for row in table) for c in range(width)]
    if not any(occupied):
        raise HeaderValidationError('Không có dữ liệu từ dòng tiêu đề đã chọn. Chọn lại hàng tiêu đề rồi bấm Đọc.', snapshot)
    first = occupied.index(True)
    last = width - occupied[::-1].index(True)
    raw_headers = [str(value).strip() if value is not None else '' for value in table[0][first:last]]

    populated_count = sum(bool(h) for h in raw_headers)
    if auto_heal and populated_count > 1:
        headers = []
        counts = collections.Counter()
        last_non_empty = None
        for i, val in enumerate(raw_headers):
            col_letter = column_name(first + i + 1)
            name = val
            if not name:
                if last_non_empty:
                    name = f"{last_non_empty}_phụ"
                else:
                    name = f"Cột_{col_letter}"
            else:
                last_non_empty = val
            counts[name] += 1
            if counts[name] > 1:
                name = f"{name}_{counts[name]}"
            headers.append(name)
        snapshot = dict(snapshot, headers=headers, data=[row[first:last] for row in rows[header_row:]], header_row=header_row, column_offset=first, auto_healed=True)
        return snapshot

    empty = [f'{column_name(first+i+1)}{header_row}' for i, value in enumerate(raw_headers) if not value]
    repeated = collections.defaultdict(list)
    for i, value in enumerate(raw_headers):
        if value:
            repeated[value].append(f'{column_name(first+i+1)}{header_row}')
    duplicates = [f'“{value[:60]}” tại ' + ', '.join(addresses) for value, addresses in repeated.items() if len(addresses) > 1]
    problems = []
    if empty:
        problems.append('Ô tiêu đề trống: ' + ', '.join(empty[:10]))
    if duplicates:
        problems.append('Tên cột trùng: ' + '; '.join(duplicates[:5]))
    if problems:
        raise HeaderValidationError(f'Sheet {snapshot["sheet"]}, dòng tiêu đề {header_row}.\n' + '\n'.join(problems) + '\nChọn dòng chứa tên cột thực tế ở ô Tiêu đề rồi bấm Đọc. Nếu đúng dòng, cần tên riêng cho từng cột; ô gộp không dùng làm tiêu đề bảng.', snapshot)
    if any(len(header) > 200 for header in raw_headers):
        raise HeaderValidationError('Tên cột vượt 200 ký tự. Kiểm tra có chọn nhầm dòng hoặc dùng bản sao với tên cột ngắn hơn.', snapshot)
    snapshot = dict(snapshot, headers=raw_headers, data=[row[first:last] for row in rows[header_row:]], header_row=header_row, column_offset=first)
    return snapshot


def inspect_file(path, sheet='', header_row=1, lenient=False, auto_heal=False):
    path = Path(path).resolve(strict=True)
    before = file_hash(path)
    metadata = preflight(path, lenient=lenient) if lenient else preflight(path)
    WORK.mkdir(exist_ok=True)
    if path.suffix.lower() == '.csv':
        snapshot = {'sheets': ['CSV'], 'sheet': 'CSV', 'rows': load_csv(path, strict_ragged=not lenient), 'excel_version': None}
    elif lenient and path.suffix.lower() == '.xlsx':
        snapshot = load_xlsx_fast(path, sheet_name=sheet)
        if not snapshot:
            with tempfile.TemporaryDirectory(prefix='inspect-', dir=WORK) as directory:
                copy = Path(directory) / ('input' + path.suffix.lower())
                shutil.copyfile(path, copy)
                if file_hash(copy) != before:
                    raise ValueError('Bản sao nguồn thay đổi trong lúc đọc; nạp lại file.')
                metadata = preflight(copy, lenient=lenient) if lenient else preflight(copy)
                snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
    else:
        with tempfile.TemporaryDirectory(prefix='inspect-', dir=WORK) as directory:
            copy = Path(directory) / ('input' + path.suffix.lower())
            shutil.copyfile(path, copy)
            if file_hash(copy) != before:
                raise ValueError('Bản sao nguồn thay đổi trong lúc đọc; nạp lại file.')
            metadata = preflight(copy, lenient=lenient) if lenient else preflight(copy)
            snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
    if file_hash(path) != before:
        raise ValueError('File nguồn thay đổi trong lúc đọc; hãy thử lại.')
    snapshot.update(path=str(path), hash=before, metadata=metadata)
    if header_row in ('auto', 0) and snapshot.get('rows'):
        header_row = detect_header_row(snapshot['rows'])
    return select_header(snapshot, header_row, auto_heal=auto_heal)


def key_value(value, format_code='General', allow_dates=False):
    if value is None or value == '':
        return ''
    if isinstance(value, bool):
        raise ValueError('Không dùng giá trị TRUE/FALSE làm mã đối chiếu.')
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise ValueError('Khóa có số không hữu hạn.')
        if re.fullmatch(r'0+', format_code) and float(value).is_integer():
            return ('-' if value < 0 else '') + str(abs(int(value))).zfill(len(format_code))
        stripped = re.sub(r'"[^"]*"|\[[^]]*\]', '', format_code.lower())
        if re.search(r'[dyh]', stripped):
            if not allow_dates:
                raise ValueError('Khóa có định dạng ngày/giờ. V1 yêu cầu chuyển mã ngày thành chuỗi theo quy tắc rõ ràng.')
            base_date = datetime.date(1899, 12, 30)
            date_val = base_date + datetime.timedelta(days=int(value))
            return date_val.strftime('%Y-%m-%d')
        value = str(int(value)) if float(value).is_integer() else str(value)
    elif allow_dates and isinstance(value, str):
        v_str = value.strip()
        m_dmy = re.fullmatch(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', v_str)
        if m_dmy:
            d, m, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
            if 1 <= m <= 12 and 1 <= d <= 31:
                try:
                    return datetime.date(y, m, d).strftime('%Y-%m-%d')
                except ValueError:
                    pass
        m_ymd = re.fullmatch(r'(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})', v_str)
        if m_ymd:
            y, m, d = int(m_ymd.group(1)), int(m_ymd.group(2)), int(m_ymd.group(3))
            if 1 <= m <= 12 and 1 <= d <= 31:
                try:
                    return datetime.date(y, m, d).strftime('%Y-%m-%d')
                except ValueError:
                    pass
    clean_val = str(value).replace('\u00a0', ' ')
    clean_val = re.sub(r'[\u200b\u200c\u200d\ufeff\u200e\u200f]', '', clean_val)
    return unicodedata.normalize('NFC', clean_val).strip().casefold()


def format_at(snapshot, data_row, col):
    address = f'{column_name(col+1+snapshot.get("column_offset",0))}{snapshot["header_row"] + data_row + 1}'
    return snapshot['metadata']['sheets'].get(snapshot['sheet'], {}).get('formats', {}).get(address, 'General')


def normalized_keys(snapshot, col, allow_dates=False):
    return [key_value(row[col], format_at(snapshot, i, col), allow_dates=allow_dates) for i, row in enumerate(snapshot['data'])]


def parse_lenient_number(value, allow_blank_zero=True):
    if value is None or value == '':
        if allow_blank_zero:
            return Decimal(0)
        raise ValueError('Cột tổng hợp có ô trống.')
    if isinstance(value, bool):
        raise ValueError('Không nhận giá trị TRUE/FALSE.')
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise ValueError('Số không hữu hạn.')
        return Decimal(str(value))
    cleaned = str(value).strip()
    cleaned = re.sub(r'[\sđĐVNDvnđ\$€¥]', '', cleaned)
    if not cleaned:
        if allow_blank_zero:
            return Decimal(0)
        raise ValueError('Số rỗng sau khi làm sạch.')
    is_negative = False
    if cleaned.startswith('(') and cleaned.endswith(')'):
        is_negative = True
        cleaned = cleaned[1:-1].strip()
    elif cleaned.endswith('-'):
        is_negative = True
        cleaned = cleaned[:-1].strip()
    elif cleaned.startswith('-'):
        is_negative = True
        cleaned = cleaned[1:].strip()
    elif cleaned.startswith('+'):
        cleaned = cleaned[1:].strip()

    if any(c in cleaned for c in ('-', '+', '(', ')')):
        raise ValueError(f'Dấu số không hợp lệ trong "{value}".')

    if '.' in cleaned and ',' in cleaned:
        if cleaned.rfind(',') > cleaned.rfind('.'):
            int_part, dec_part = cleaned.rsplit(',', 1)
            groups = int_part.split('.')
            if len(groups) > 1 and any(len(g) != 3 for g in groups[1:]):
                raise ValueError(f'Định dạng số không hợp lệ: "{value}".')
            cleaned = int_part.replace('.', '') + '.' + dec_part
        else:
            int_part, dec_part = cleaned.rsplit('.', 1)
            groups = int_part.split(',')
            if len(groups) > 1 and any(len(g) != 3 for g in groups[1:]):
                raise ValueError(f'Định dạng số không hợp lệ: "{value}".')
            cleaned = int_part.replace(',', '') + '.' + dec_part
    elif ',' in cleaned:
        parts = cleaned.split(',')
        if len(parts) == 2 and len(parts[1]) <= 2:
            cleaned = cleaned.replace(',', '.')
        else:
            if any(len(p) != 3 for p in parts[1:]):
                raise ValueError(f'Định dạng dấu phẩy không hợp lệ: "{value}".')
            cleaned = cleaned.replace(',', '')
    elif '.' in cleaned:
        parts = cleaned.split('.')
        if len(parts) > 2:
            if any(len(p) != 3 for p in parts[1:]):
                raise ValueError(f'Định dạng dấu chấm không hợp lệ: "{value}".')
            cleaned = cleaned.replace('.', '')
    try:
        res = Decimal(cleaned)
        if is_negative:
            res = -res
    except InvalidOperation as exc:
        raise ValueError(f'Không thể chuyển đổi giá trị "{value}" sang số.') from exc
    if not res.is_finite() or abs(res) > Decimal('1e14'):
        raise ValueError('Số vượt giới hạn 10^14.')
    return res


def number(value, lenient=False):
    if lenient:
        return parse_lenient_number(value)
    if value is None or value == '' or isinstance(value, bool):
        raise ValueError('Cột tổng hợp có ô trống/TRUE/FALSE; v1 không tự đổi thành 0.')
    if isinstance(value, str) and not re.fullmatch(r'[+-]?\d+(?:\.\d+)?', value.strip()):
        raise ValueError('Số dạng chuỗi không rõ định dạng. Chỉ nhận số Excel hoặc chuỗi số dùng dấu chấm thập phân, không phân nhóm nghìn.')
    try:
        result = Decimal(str(value).strip())
    except InvalidOperation as exc:
        raise ValueError('Giá trị không phải số hợp lệ.') from exc
    if not result.is_finite() or abs(result) > Decimal('1e14'):
        raise ValueError('Giá trị không hữu hạn hoặc vượt 10^14 của v1.')
    if len(result.normalize().as_tuple().digits) > 15:
        raise ValueError('V1 không chuyển số có hơn 15 chữ số có nghĩa sang Excel.')
    return result


def validate_plan(plan, source, reference=None, allow_clarify=False):
    allowed = {'action', 'key', 'value', 'ref_key', 'ref_value', 'message', 'lenient_numbers', 'allow_dates', 'provider', 'clean_output', 'engine_used', 'condition', 'selected_columns'}
    if not isinstance(plan, dict) or set(plan) - allowed:
        raise ValueError('Kế hoạch chứa trường không được phép.')
    for k in ('action', 'key', 'value', 'ref_key', 'ref_value', 'message', 'engine_used', 'condition', 'selected_columns'):
        if k in plan:
            val = plan.get(k, '')
            if not isinstance(val, str) or len(val) > 2000:
                raise ValueError(f'Tham số kế hoạch {k} phải là chuỗi ngắn.')
    action = plan.get('action')
    if action not in ('lookup', 'summary', 'clean', 'filter', 'clarify'):
        raise ValueError('Thao tác ngoài danh sách được phép.')
    if action == 'clarify':
        if allow_clarify:
            return plan
        raise ValueError(plan.get('message') or 'Yêu cầu chưa đủ rõ.')
    if plan.get('key') not in source['headers']:
        match = ai_client.fuzzy_match_column(plan.get('key', ''), source['headers'])
        if match in source['headers']:
            plan['key'] = match
        elif allow_clarify:
            return {'action': 'clarify', 'key': '', 'value': '', 'ref_key': '', 'ref_value': '', 'message': f'Không tìm thấy cột "{plan.get("key")}" trong bảng nguồn. Vui lòng kiểm tra lại tên cột.'}
        else:
            raise ValueError(f'Cột khóa/nhóm/làm sạch/lọc "{plan.get("key")}" không tồn tại trong nguồn.')
    if action == 'filter':
        if not plan.get('value'):
            plan['value'] = '*'
        sel = plan.get('selected_columns') or plan.get('ref_key')
        if sel:
            raw_candidates = [c.strip() for c in re.split(r'[,;\n]', str(sel)) if c.strip()]
            resolved = []
            for c in raw_candidates:
                matched_c = c if c in source['headers'] else ai_client.fuzzy_match_column(c, source['headers'])
                if matched_c in source['headers'] and matched_c not in resolved:
                    resolved.append(matched_c)
            if resolved:
                plan['selected_columns'] = ', '.join(resolved)
                plan['ref_key'] = plan['selected_columns']
        return plan
    if action == 'summary':
        if plan.get('value') not in source['headers']:
            match = ai_client.fuzzy_match_column(plan.get('value', ''), source['headers'])
            if match in source['headers']:
                plan['value'] = match
            elif allow_clarify:
                return {'action': 'clarify', 'key': plan.get('key', ''), 'value': '', 'ref_key': '', 'ref_value': '', 'message': f'Không tìm thấy cột cần cộng "{plan.get("value")}" trong nguồn.'}
            else:
                raise ValueError(f'Cột cần cộng "{plan.get("value")}" không tồn tại trong nguồn.')
    if action == 'lookup':
        if reference is None:
            if allow_clarify:
                return {'action': 'clarify', 'key': plan.get('key', ''), 'value': '', 'ref_key': '', 'ref_value': '', 'message': 'Yêu cầu đối chiếu cần có bảng danh mục. Vui lòng chọn và nạp file Danh mục trước khi tiếp tục.'}
            raise ValueError('Đối chiếu cần một bảng danh mục.')
        if plan.get('ref_key') not in reference['headers']:
            match = ai_client.fuzzy_match_column(plan.get('ref_key', ''), reference['headers'])
            if match in reference['headers']:
                plan['ref_key'] = match
            elif allow_clarify:
                return {'action': 'clarify', 'key': plan.get('key', ''), 'value': '', 'ref_key': '', 'ref_value': '', 'message': f'Không tìm thấy cột đối chiếu "{plan.get("ref_key")}" trong danh mục.'}
            else:
                raise ValueError(f'Cột đối chiếu "{plan.get("ref_key")}" không tồn tại trong danh mục.')
        if plan.get('ref_value') not in reference['headers']:
            match = ai_client.fuzzy_match_column(plan.get('ref_value', ''), reference['headers'])
            if match in reference['headers']:
                plan['ref_value'] = match
            elif allow_clarify:
                return {'action': 'clarify', 'key': plan.get('key', ''), 'value': '', 'ref_key': '', 'ref_value': '', 'message': f'Không tìm thấy cột trả về "{plan.get("ref_value")}" trong danh mục.'}
            else:
                raise ValueError(f'Cột trả về "{plan.get("ref_value")}" không tồn tại trong danh mục.')
    return plan


def escape_criteria(cell):
    return f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({cell},"~","~~"),"*","~*"),"?","~?")'


def prepare_sheets(source, reference, plan, token):
    validate_plan(plan, source, reference)
    name = 'Ket_qua_' + token
    helper = '_EL_' + token
    action = plan['action']
    key_col = source['headers'].index(plan['key'])
    allow_dates = bool(plan.get('allow_dates', False))
    lenient_nums = bool(plan.get('lenient_numbers', False))
    keys = normalized_keys(source, key_col, allow_dates=allow_dates) if action not in ('clean', 'filter') else []
    rows = [list(row) for row in source['data']]
    if not rows:
        raise ValueError('Nguồn không có dòng dữ liệu.')
    for i, row in enumerate(rows):
        fmt = format_at(source, i, key_col)
        if action == 'lookup' and isinstance(row[key_col], (int, float)) and re.fullmatch(r'0+', fmt):
            row[key_col] = key_value(row[key_col], fmt, allow_dates=allow_dates)
    formulas = []
    sheets = []
    stats = {'action': action, 'source_rows': len(rows)}
    if action not in ('clean', 'filter'):
        stats['blank_keys'] = sum(not k for k in keys)
    if action == 'lookup':
        ref_col = reference['headers'].index(plan['ref_key'])
        val_col = reference['headers'].index(plan['ref_value'])
        ref_keys = normalized_keys(reference, ref_col, allow_dates=allow_dates)
        groups = collections.defaultdict(list)
        for i, key in enumerate(ref_keys):
            if key:
                groups[key].append(reference['data'][i][val_col])
        table = [['Khoa_chuan_hoa', 'Gia_tri']]
        for key, values in groups.items():
            if len(values) == 1:
                table.append([key, values[0]])
        if len(table) == 1:
            table.append(['', None])
        clean_output = bool(plan.get('clean_output', False))
        if clean_output:
            max_helper_len = max(len(table), len(keys) + 1)
            helper_rows = []
            for r in range(max_helper_len):
                row_ab = table[r] if r < len(table) else ['', None]
                key_d = 'Khoa_nguon' if r == 0 else (keys[r - 1] if r - 1 < len(keys) else '')
                helper_rows.append([row_ab[0], row_ab[1], '', key_d])
            sheets.append({'name': helper, 'rows': helper_rows, 'formulas': [], 'formats': []})
            target_col_name = plan.get('value')
            if target_col_name and target_col_name in source['headers']:
                headers = list(source['headers'])
                value_result_col = headers.index(target_col_name) + 1
                inplace_target = True
            else:
                val_label = plan.get('ref_value', 'Giá trị đối chiếu')
                label = f'{val_label} (đối chiếu)' if val_label in source['headers'] else val_label
                unique = label
                while unique in source['headers']:
                    unique += ' (mới)'
                headers = source['headers'] + [unique]
                value_result_col = len(headers)
                inplace_target = False
            stats.update(matched=0, missing=0, duplicate=0)
            for i, (row, key) in enumerate(zip(rows, keys), 2):
                status = 'KHOP'
                if not key:
                    status = 'KHOA_TRONG'
                elif key not in groups:
                    status = 'KHONG_TIM_THAY'
                elif len(groups[key]) > 1:
                    status = 'KHOA_TRUNG'
                if status == 'KHOP':
                    stats['matched'] += 1
                    last = len(table)
                    match = f'MATCH({escape_criteria(f"\'{helper}\'!$D${i}")},\'{helper}\'!$A$2:$A${last},0)'
                    result = f'INDEX(\'{helper}\'!$B$2:$B${last},{match})'
                    formula = f'=IF({result}="","",{result})'
                    expected = groups[key][0]
                    formulas.append({'row': i, 'col': value_result_col, 'formula': formula, 'expected': expected if expected is not None else ''})
                elif status == 'KHOA_TRUNG':
                    stats['duplicate'] += 1
                else:
                    stats['missing'] += 1
                if inplace_target:
                    row[value_result_col - 1] = None
                else:
                    row.append(None)
            stats['reference_duplicate_keys'] = sum(len(v) > 1 for v in groups.values())
        else:
            sheets.append({'name': helper, 'rows': table, 'formulas': [], 'formats': []})
            extra_headers = []
            for label in ('Giá trị đối chiếu', 'Trạng thái đối chiếu', 'Khóa chuẩn hóa'):
                unique = label
                while unique in source['headers'] or unique in extra_headers:
                    unique += ' (mới)'
                extra_headers.append(unique)
            headers = source['headers'] + extra_headers
            key_result_col = len(headers)
            value_result_col = len(headers) - 2
            stats.update(matched=0, missing=0, duplicate=0)
            for i, (row, key) in enumerate(zip(rows, keys), 2):
                status = 'KHOP'
                if not key:
                    status = 'KHOA_TRONG'
                elif key not in groups:
                    status = 'KHONG_TIM_THAY'
                elif len(groups[key]) > 1:
                    status = 'KHOA_TRUNG'
                if status == 'KHOP':
                    stats['matched'] += 1
                    last = len(table)
                    match = f'MATCH({escape_criteria(column_name(key_result_col)+str(i))},\'{helper}\'!$A$2:$A${last},0)'
                    result = f'INDEX(\'{helper}\'!$B$2:$B${last},{match})'
                    formula = f'=IF({result}="","",{result})'
                    expected = groups[key][0]
                    formulas.append({'row': i, 'col': value_result_col, 'formula': formula, 'expected': expected if expected is not None else ''})
                elif status == 'KHOA_TRUNG':
                    stats['duplicate'] += 1
                else:
                    stats['missing'] += 1
                row.extend([None, status, key])
            stats['reference_duplicate_keys'] = sum(len(v) > 1 for v in groups.values())
    elif action == 'summary':
        value_col = source['headers'].index(plan['value'])
        totals = collections.defaultdict(lambda: Decimal(0))
        counts = collections.Counter()
        grouped = collections.defaultdict(list)
        table = [['Nhom_chuan_hoa', 'Gia_tri']]
        for row, key in zip(rows, keys):
            if not key:
                raise ValueError('Cột nhóm có khóa trống; cần xử lý rõ trước khi tổng hợp.')
            val_num = number(row[value_col], lenient=lenient_nums)
            totals[key] += val_num
            counts[key] += 1
            grouped[key].append(float(val_num))
        headers = ['Nhom_chuan_hoa', 'So_dong', 'Tong']
        rows = []
        for i, key in enumerate(totals, 2):
            number(totals[key], lenient=lenient_nums)
            rows.append([key, None, None])
            first = len(table) + 1
            table.extend([key, value] for value in grouped[key])
            last = len(table)
            guard = f'A{i}=\'{helper}\'!$A${first}'
            values = f"'{helper}'!$B${first}:$B${last}"
            formulas.extend([
                {'row': i, 'col': 2, 'formula': f'=IF({guard},COUNT({values}),NA())', 'expected': counts[key]},
                {'row': i, 'col': 3, 'formula': f'=IF({guard},SUM({values}),NA())', 'expected': float(totals[key])},
            ])
        sheets.append({'name': helper, 'rows': table, 'formulas': [], 'formats': []})
        stats.update(groups=len(totals), independent_total=str(sum(totals.values())))
    elif action == 'filter':
        crit = str(plan.get('value', '')).strip()
        filtered_rows = []
        filtered_indices = []

        # 1. Numeric comparison: > 8, >= 8.0, < 5, <= 5, == 10, != 0
        num_match = re.match(r'^([><!=]=?)\s*([0-9]+(?:[\.,][0-9]+)?)$', crit)

        # 2. Non-empty check
        is_not_empty = crit in ('*', '[khác rỗng]', 'khác rỗng', 'không trống', 'khac rong', 'khong trong')

        # 3. Comma-separated keywords
        keywords = [unicodedata.normalize('NFC', k.strip().lower()) for k in re.split(r'[,;\n/]', crit) if k.strip()]

        for orig_i, row in enumerate(rows):
            # Skip completely blank rows
            if not any(c is not None and str(c).strip() != '' for c in row):
                continue

            # Skip trailing administrative/signature rows (e.g. "Hà Nội, ngày...", "Người lập biểu")
            non_empty_cells = [str(c).strip() for c in row if c is not None and str(c).strip() != '']
            if len(non_empty_cells) <= 2 and any(any(sig in c.lower() for sig in ['người lập biểu', 'trưởng khoa', 'hiệu trưởng', 'giám đốc', 'ngày ... tháng', 'ngày   tháng']) for c in non_empty_cells):
                continue

            raw_val = row[key_col]
            matched = False

            if is_not_empty:
                if raw_val is not None:
                    s_val = str(raw_val).strip()
                    if s_val and s_val.lower() not in ('-', 'none', 'null', 'không', 'k'):
                        matched = True
            elif num_match:
                op, target_str = num_match.groups()
                target_val = float(target_str.replace(',', '.'))
                try:
                    c_num = float(str(raw_val).replace(',', '.'))
                    if op == '>': matched = (c_num > target_val)
                    elif op == '>=': matched = (c_num >= target_val)
                    elif op == '<': matched = (c_num < target_val)
                    elif op == '<=': matched = (c_num <= target_val)
                    elif op in ('==', '='): matched = (c_num == target_val)
                    elif op == '!=': matched = (c_num != target_val)
                except (ValueError, TypeError):
                    matched = False
            else:
                if raw_val is not None:
                    cell_clean = unicodedata.normalize('NFC', str(raw_val).strip().lower())
                    if any(kw in cell_clean for kw in keywords):
                        matched = True

            if matched:
                filtered_rows.append(list(row[:len(source['headers'])]))
                filtered_indices.append(orig_i)

        sel = plan.get('selected_columns') or plan.get('ref_key')
        selected_cols = []
        if sel:
            raw_cands = [c.strip() for c in re.split(r'[,;\n]', str(sel)) if c.strip()]
            for c in raw_cands:
                mc = c if c in source['headers'] else ai_client.fuzzy_match_column(c, source['headers'])
                if mc in source['headers'] and mc not in selected_cols:
                    selected_cols.append(mc)

        if selected_cols:
            col_indices = [source['headers'].index(c) for c in selected_cols]
            headers = selected_cols
            rows = [[r[idx] for idx in col_indices] for r in filtered_rows]
            stats.update(filtered_rows=len(filtered_rows), criteria=crit, source_rows=len(source['data']), selected_columns=len(headers))
        else:
            headers = source['headers'][:]
            rows = filtered_rows
            stats.update(filtered_rows=len(filtered_rows), criteria=crit, source_rows=len(source['data']))
    else:
        changed = 0
        for row in rows:
            value = row[key_col]
            if isinstance(value, str):
                c_val = value.replace('\u00a0', ' ')
                c_val = re.sub(r'[\u200b\u200c\u200d\ufeff\u200e\u200f]', '', c_val)
                clean = unicodedata.normalize('NFC', c_val).strip()
                if clean != value:
                    changed += 1
                row[key_col] = clean
        headers = source['headers'][:]
        stats['changed_cells'] = changed
    formats = []
    if action == 'filter':
        if selected_cols:
            col_indices = [source['headers'].index(c) for c in selected_cols]
            for i, (orig_i, row) in enumerate(zip(filtered_indices, rows), 2):
                for c, orig_c in enumerate(col_indices):
                    value = row[c]
                    fmt = format_at(source, orig_i, orig_c)
                    if isinstance(value, (int, float)) and not isinstance(value, bool) and fmt != 'General':
                        formats.append({'row': i, 'col': c + 1, 'format': fmt})
        else:
            for i, (orig_i, row) in enumerate(zip(filtered_indices, rows), 2):
                for c, value in enumerate(row[:len(source['headers'])]):
                    fmt = format_at(source, orig_i, c)
                    if isinstance(value, (int, float)) and not isinstance(value, bool) and fmt != 'General':
                        formats.append({'row': i, 'col': c + 1, 'format': fmt})
    elif action != 'summary':
        for i, row in enumerate(rows, 2):
            for c, value in enumerate(row[:len(source['headers'])]):
                fmt = format_at(source, i - 2, c)
                if isinstance(value, (int, float)) and not isinstance(value, bool) and fmt != 'General':
                    formats.append({'row': i, 'col': c + 1, 'format': fmt})
    sheets.append({'name': name, 'rows': [headers] + rows, 'formulas': formulas, 'formats': formats})
    audit_name = 'Kiem_tra_' + token
    audit_rows = [['Hang_muc', 'Ket_qua'], ['Thao_tac', action], ['Sheet_nguon', source['sheet']], ['SHA256_nguon', source['hash']], ['Chinh_sach_khoa', 'NFC, bo khoang trang hai dau, khong phan biet hoa thuong; giu so 0 dau'], ['Cong_thuc_can_doi_soat', len(formulas)]]
    audit_rows += [[key, str(value)] for key, value in stats.items()]
    sheets.append({'name': audit_name, 'rows': audit_rows, 'formulas': [], 'formats': []})
    return sheets, name, stats


def preview_plan(source, reference, plan, export_pdf=False, output_dir=OUTPUT):
    validate_plan(plan, source, reference)
    reference = reference if plan['action'] == 'lookup' else None
    for snapshot in (source, reference):
        if snapshot and file_hash(snapshot['path']) != snapshot['hash']:
            raise ValueError('File đã thay đổi sau khi đọc; nạp lại để kiểm tra kết quả dự kiến.')
    sheets, name, stats = prepare_sheets(source, reference, plan, 'preview')
    result = next(sheet for sheet in sheets if sheet['name'] == name)
    samples = [row[:] for row in result['rows'][1:11]]
    for patch in result['formulas']:
        if patch['row'] <= len(samples) + 1:
            samples[patch['row'] - 2][patch['col'] - 1] = patch['expected']
    preview = dict(plan=dict(plan), source=dict(path=source['path'], sheet=source['sheet'], hash=source['hash'], header_row=source['header_row'], column_offset=source.get('column_offset', 0)),
                   reference=dict(path=reference['path'], sheet=reference['sheet'], hash=reference['hash'], header_row=reference['header_row'], column_offset=reference.get('column_offset', 0)) if reference else None,
                   export_pdf=bool(export_pdf), stats=stats, columns=result['rows'][0], sample_rows=samples,
                   formula_checks=len(result['formulas']), output_directory=str(Path(output_dir).resolve()),
                   scope='Tạo workbook mới; thêm sheet kết quả, dữ liệu đối soát và kiểm tra. Giữ file gốc.',
                   status='Chưa thực hiện; kết quả dự kiến tính bằng Python, sẽ đối soát bằng Excel sau khi chấp nhận.')
    preview['snapshot_digest'] = hashlib.sha256(json.dumps([
        {key: snapshot[key] for key in ('headers', 'data', 'metadata')} if snapshot else None
        for snapshot in (source, reference)], ensure_ascii=False, sort_keys=True, allow_nan=False, default=str).encode('utf-8')).hexdigest()
    preview['approval_token'] = hashlib.sha256(json.dumps(preview, ensure_ascii=False, sort_keys=True, allow_nan=False, default=str).encode('utf-8')).hexdigest()
    return preview


def execute(source, reference, plan, export_pdf=False, output_dir=OUTPUT, *, approval_token):
    validate_plan(plan, source, reference)
    reference = reference if plan['action'] == 'lookup' else None
    if approval_token != preview_plan(source, reference, plan, export_pdf, output_dir)['approval_token']:
        raise ValueError('Thiếu chấp nhận hoặc bản xem trước đã đổi; cần kiểm tra kết quả dự kiến lại.')
    for snapshot in (source, reference):
        if snapshot and file_hash(snapshot['path']) != snapshot['hash']:
            raise ValueError('File đã thay đổi sau khi đọc; cần nạp lại trước khi chạy.')
    token = uuid.uuid4().hex[:8]
    sheets, result_name, stats = prepare_sheets(source, reference, plan, token)
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    WORK.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='run-', dir=WORK) as directory:
        directory = Path(directory)
        input_copy = None
        if Path(source['path']).suffix.lower() in ('.xlsx', '.xls'):
            input_copy = directory / ('input' + Path(source['path']).suffix.lower())
            shutil.copyfile(source['path'], input_copy)
            if file_hash(input_copy) != source['hash']:
                raise ValueError('Bản sao nguồn không khớp với dữ liệu đã đọc.')
        staging = directory / 'result.xlsx'
        pdf = directory / 'preview.pdf' if export_pdf else None
        verified = bridge({'operation': 'write', 'input': str(input_copy) if input_copy else None, 'output': str(staging), 'pdf': str(pdf) if pdf else None, 'result_sheet': result_name, 'sheets': sheets, 'tolerance': 0.000001}, directory)
        if not staging.exists() or not zipfile.is_zipfile(staging):
            raise RuntimeError('Excel chưa tạo XLSX hợp lệ.')
        saved_checks = verify_saved_result(staging, sheets)
        preservation = preserve_check(source['path'], staging) if input_copy else {'original_sheets': 0, 'unchanged': True}
        if not preservation['unchanged']:
            raise RuntimeError('Workbook thay đổi nội dung ngoài phạm vi; không bàn giao kết quả: ' + ', '.join(preservation['changed_addresses']))
        for snapshot in (source, reference):
            if snapshot and file_hash(snapshot['path']) != snapshot['hash']:
                raise RuntimeError('File nguồn thay đổi trong lúc chạy; không bàn giao kết quả.')
        expected_checks = sum(len(spec['formulas']) for spec in sheets)
        if verified['formula_checks'] != expected_checks:
            raise RuntimeError('Số phép kiểm tra công thức không khớp.')
        basename = f'{time.strftime("%Y%m%d-%H%M%S")}_{plan["action"]}_{token}'
        target = output_dir / (basename + '.xlsx')
        pdf_target = None
        if pdf:
            if not pdf.exists() or not pdf.read_bytes().startswith(b'%PDF-'):
                raise RuntimeError('PDF không xuất thành công.')
            pdf_target = output_dir / (basename + '.pdf')
        report = dict(stats, **verified, approval=dict(token=approval_token, plan=dict(plan), export_pdf=bool(export_pdf)), saved_cell_checks=saved_checks, preservation=preservation, source_unchanged=True, output=str(target), pdf=str(pdf_target) if pdf_target else None, pdf_layout='Chưa kiểm tra trực quan' if pdf_target else 'Không yêu cầu', security='Local; chưa nghiệm thu mã hóa/egress/AutoRecover của toàn máy', key_policy='NFC + trim + casefold; giữ 0 đầu', output_sha256=file_hash(staging))
        report_staging = directory / 'report.json'
        report_staging.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        files = [(staging, target)]
        if pdf:
            files.append((pdf, pdf_target))
        files.append((report_staging, output_dir / (basename + '.json')))
        published = []
        try:
            for private_path, public_path in files:
                publish_file(private_path, public_path)
                published.append(public_path)
        except Exception:
            for public_path in published:
                public_path.unlink(missing_ok=True)
            raise
    return report


def execute_batch(source_dir, reference_path=None, plan=None, output_dir=OUTPUT, *,
                  export_pdf=False, file_pattern=None, sheet='', header_row=1,
                  auto_heal=True, lenient=True, on_progress=None):
    """Execute an approved plan across all spreadsheets in source_dir with 0 AI tokens.
    Processes each file independently, catches per-file errors, and publishes a summary report.
    """
    import fnmatch
    source_dir = Path(source_dir).resolve(strict=True)
    if not source_dir.is_dir():
        raise NotADirectoryError(f'Thư mục nguồn không hợp lệ: {source_dir}')

    if not plan or not isinstance(plan, dict) or plan.get('action') not in ('lookup', 'summary', 'clean', 'filter'):
        raise ValueError('Kế hoạch thực thi không hợp lệ. Cần một kế hoạch đối chiếu, tổng hợp, lọc hoặc làm sạch đã duyệt.')

    output_dir = Path(output_dir or OUTPUT).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1. Inspect reference once if lookup
    ref_snapshot = None
    if plan['action'] == 'lookup':
        if not reference_path:
            raise ValueError('Thao tác đối chiếu (lookup) cần có file danh mục tham chiếu.')
        ref_snapshot = inspect_file(reference_path, lenient=lenient, auto_heal=auto_heal)
        if plan.get('ref_key') and plan['ref_key'] not in ref_snapshot['headers']:
            matched_rk = ai_client.fuzzy_match_column(plan['ref_key'], ref_snapshot['headers'])
            if matched_rk and matched_rk in ref_snapshot['headers']:
                plan['ref_key'] = matched_rk
        if plan.get('ref_value') and plan['ref_value'] not in ref_snapshot['headers']:
            matched_rv = ai_client.fuzzy_match_column(plan['ref_value'], ref_snapshot['headers'])
            if matched_rv and matched_rv in ref_snapshot['headers']:
                plan['ref_value'] = matched_rv

    # 2. Collect candidate files
    valid_exts = ('.xlsx', '.xls', '.csv')
    candidates = []
    for item in sorted(source_dir.iterdir()):
        if not item.is_file():
            continue
        if item.name.startswith(('~', '.~', '.')):
            continue
        if item.suffix.lower() not in valid_exts:
            continue
        if file_pattern and not fnmatch.fnmatch(item.name.lower(), file_pattern.lower()):
            continue
        if item.parent == output_dir and (bool(re.match(r'^\d{8}-\d{6}_', item.name)) or item.name.startswith('batch_report_')):
            continue
        candidates.append(item)

    if not candidates:
        raise ValueError(f'Không tìm thấy file bảng tính (XLSX, XLS, CSV) nào trong thư mục {source_dir}.')

    start_time = time.time()
    results = []
    succeeded = 0
    failed = 0

    for idx, fpath in enumerate(candidates):
        item_res = {
            'index': idx + 1,
            'file': fpath.name,
            'path': str(fpath),
            'status': 'pending',
            'error': None,
            'output': None,
            'rows': 0,
            'stats': {}
        }
        try:
            h_row = header_row
            try:
                src = inspect_file(fpath, sheet=sheet, header_row=h_row, lenient=lenient, auto_heal=auto_heal)
            except HeaderValidationError:
                src = inspect_file(fpath, sheet=sheet, header_row='auto', lenient=lenient, auto_heal=auto_heal)
            item_res['rows'] = len(src.get('data', []))

            # Adaptive column matching:
            file_plan = dict(plan)
            file_plan.setdefault('clean_output', True)
            file_plan.setdefault('lenient_numbers', True)
            file_plan.setdefault('allow_dates', True)
            req_key = file_plan.get('key', '')
            if req_key and req_key not in src['headers']:
                matched_k = ai_client.fuzzy_match_column(req_key, src['headers'])
                if matched_k and matched_k in src['headers']:
                    file_plan['key'] = matched_k
                else:
                    raise ValueError(f"File thiếu cột khóa '{req_key}'. Các cột hiện có: {', '.join(src['headers'][:5])}")

            if file_plan['action'] == 'summary':
                v_col = file_plan.get('value', '')
                if v_col and v_col not in src['headers']:
                    matched_v = ai_client.fuzzy_match_column(v_col, src['headers'])
                    if matched_v and matched_v in src['headers']:
                        file_plan['value'] = matched_v
                    else:
                        raise ValueError(f"File thiếu cột số '{v_col}'.")

            prev = preview_plan(src, ref_snapshot, file_plan, export_pdf=export_pdf, output_dir=output_dir)
            exec_res = execute(src, ref_snapshot, file_plan, export_pdf=export_pdf, output_dir=output_dir, approval_token=prev['approval_token'])

            item_res['status'] = 'success'
            item_res['output'] = exec_res['output']
            item_res['pdf'] = exec_res.get('pdf')
            item_res['stats'] = exec_res.get('stats', {})
            succeeded += 1
        except Exception as exc:
            item_res['status'] = 'failed'
            item_res['error'] = str(exc)
            failed += 1

        results.append(item_res)
        if callable(on_progress):
            try:
                on_progress(idx + 1, len(candidates), item_res)
            except Exception:
                pass

    elapsed = round(time.time() - start_time, 2)
    batch_token = uuid.uuid4().hex[:8]
    summary_report = {
        'batch_id': batch_token,
        'timestamp': time.strftime('%Y-%m-%d %H:%M:%S'),
        'source_dir': str(source_dir),
        'plan': plan,
        'total_files': len(candidates),
        'succeeded': succeeded,
        'failed': failed,
        'duration_seconds': elapsed,
        'zero_tokens_used': True,
        'results': results
    }
    report_file = output_dir / f"batch_report_{time.strftime('%Y%m%d-%H%M%S')}_{batch_token}.json"
    report_file.write_text(json.dumps(summary_report, ensure_ascii=False, indent=2), encoding='utf-8')
    summary_report['report_file'] = str(report_file)
    return summary_report


def xml_text(node):
    text = ''.join(item.text or '' for item in node.findall('s:t', NS) + node.findall('s:r/s:t', NS))
    return decode_excel_text(text)


def decode_excel_text(text):
    unescaped = re.sub(r'_x([0-9A-Fa-f]{4})_', lambda match: chr(int(match[1], 16)), text)
    return unescaped.encode('utf-16-le', errors='surrogatepass').decode('utf-16-le')


def xml_value(cell, strings):
    kind = cell.get('t', 'n')
    value = cell.findtext('s:v', default='', namespaces=NS)
    if kind == 's':
        return ('text', strings[int(value)])
    if kind == 'inlineStr':
        return ('text', xml_text(cell.find('s:is', NS)))
    if kind == 'b':
        return ('bool', value == '1')
    if kind == 'str':
        return ('text', decode_excel_text(value))
    if kind in ('n', '') and value:
        return ('number', Decimal(value))
    if not value:
        return ('empty', None)
    return (kind, value)


def worksheet_parts(archive):
    book = read_xml(archive, 'xl/workbook.xml')
    rels = {r.get('Id'): r.get('Target') for r in read_xml(archive, 'xl/_rels/workbook.xml.rels')}
    return {sheet.get('name'): rels[sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')] for sheet in book.findall('s:sheets/s:sheet', NS)}


def verify_saved_result(path, specs):
    checks = 0
    with zipfile.ZipFile(path) as archive:
        strings = [xml_text(item) for item in read_xml(archive, 'xl/sharedStrings.xml')] if 'xl/sharedStrings.xml' in archive.namelist() else []
        parts = worksheet_parts(archive)
        for spec in specs:
            target = parts[spec['name']]
            target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            cells = {cell.get('r'): cell for cell in read_xml(archive, target).findall('.//s:sheetData/s:row/s:c', NS)}
            formulas = {(patch['row'], patch['col']): patch for patch in spec['formulas']}
            for r, row in enumerate(spec['rows'], 1):
                for c, expected in enumerate(row, 1):
                    cell = cells.get(column_name(c) + str(r))
                    patch = formulas.get((r, c))
                    if (cell is not None and cell.find('s:f', NS) is not None) != bool(patch):
                        raise RuntimeError('Công thức sau khi lưu không khớp với kế hoạch.')
                    if patch:
                        expected = patch['expected']
                    kind, actual = xml_value(cell, strings) if cell is not None else ('empty', None)
                    if expected is None or expected == '':
                        valid = actual is None or actual == ''
                    elif isinstance(expected, bool):
                        valid = kind == 'bool' and actual == expected
                    elif isinstance(expected, (int, float)):
                        valid = kind == 'number' and actual.is_finite() and abs(actual - Decimal(str(expected))) <= Decimal('0.000001')
                    else:
                        valid = kind == 'text' and actual == str(expected)
                    if not valid:
                        raise RuntimeError(f'Giá trị sau khi lưu không khớp: {spec["name"]}!{column_name(c)}{r}')
                    checks += 1
    return checks


def semantic_sheets(path):
    with zipfile.ZipFile(path) as archive:
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            for item in read_xml(archive, 'xl/sharedStrings.xml'):
                strings.append(xml_text(item))
        book = read_xml(archive, 'xl/workbook.xml')
        rels = {r.get('Id'): r.get('Target') for r in read_xml(archive, 'xl/_rels/workbook.xml.rels')}
        result = {}
        for sheet in book.findall('s:sheets/s:sheet', NS):
            rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target = rels[rid]
            target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            cells = {}
            for cell in read_xml(archive, target).findall('.//s:sheetData/s:row/s:c', NS):
                formula = cell.find('s:f', NS)
                if formula is not None:
                    content = ('formula', formula.text or '', tuple(sorted(formula.attrib.items())))
                else:
                    content = xml_value(cell, strings)
                if content[0] != 'empty':
                    cells[cell.get('r')] = content
            result[sheet.get('name')] = cells
        return result


def preserve_check(original, result):
    if not zipfile.is_zipfile(original):
        return {'original_sheets': 1, 'unchanged': True, 'changed_sheets': [], 'changed_addresses': [], 'scope': 'Chuyển đổi từ định dạng XLS cũ sang XLSX'}
    before = semantic_sheets(original)
    after = semantic_sheets(result)
    changed = []
    addresses = []
    for name, before_cells in before.items():
        after_cells = after.get(name, {})
        sheet_changed = False
        all_keys = set(before_cells) | set(after_cells)
        for address in all_keys:
            val_before = before_cells.get(address)
            val_after = after_cells.get(address)
            if val_before == val_after:
                continue
            # If both are numbers, check numeric tolerance (for Excel float normalization, e.g. 3E-12 on Google Forms timestamps)
            if (val_before and val_after and
                val_before[0] == 'number' and val_after[0] == 'number'):
                diff = abs(val_before[1] - val_after[1])
                if diff <= Decimal('0.00001'):
                    continue
            sheet_changed = True
            addresses.append(f'{name}!{address}')
        if sheet_changed:
            changed.append(name)
    return {'original_sheets': len(before), 'unchanged': not changed, 'changed_sheets': changed, 'changed_addresses': addresses[:10], 'scope': 'Giá trị và công thức của sheet gốc; chưa chứng minh mọi đối tượng/định dạng'}


def ai_plan(prompt, source, reference=None, allow_clarify=False, provider=None, api_key=None, model=None):
    plan = ai_client.request_plan(prompt, source['headers'], reference['headers'] if reference else (), provider=provider, api_key=api_key, model=model)
    return validate_plan(plan, source, reference, allow_clarify)


def gui(smoke=False):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class App:
        def __init__(self, window):
            cleanup_stale_work_dirs()
            self.window = window
            self.source = None
            self.reference = None
            self.jobs = queue.Queue()
            self.busy = False
            self.abort_requested = False
            self.current_worker_id = None
            self.last_output = None
            self.current_step = 1

            cfg = ai_client.get_config()
            self.ai_provider = tk.StringVar(value=cfg.get('provider', 'smart'))

            window.title('Excel Local Expert · Trợ Lý Bảng Tính Cục Bộ & AI v0.3.1')
            window.protocol('WM_DELETE_WINDOW', self.close)
            width = min(1160, window.winfo_screenwidth() - 60)
            height = min(870, window.winfo_screenheight() - 80)
            window.geometry(f'{width}x{height}+{max(0, (window.winfo_screenwidth() - width) // 2)}+12')
            window.minsize(min(980, width), min(750, height))

            style = ttk.Style()
            style.theme_use('vista')
            style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), foreground='#0d6efd')
            style.configure('Sub.TLabel', font=('Segoe UI', 9), foreground='#495057')
            style.configure('StepActive.TLabel', font=('Segoe UI', 10, 'bold'), foreground='#0d6efd')
            style.configure('StepDone.TLabel', font=('Segoe UI', 10, 'bold'), foreground='#198754')
            style.configure('StepPending.TLabel', font=('Segoe UI', 10), foreground='#adb5bd')
            style.configure('Highlight.TButton', font=('Segoe UI', 10, 'bold'))

            outer = ttk.Frame(window, padding=16)
            outer.pack(fill='both', expand=True)

            # Header
            header_frame = ttk.Frame(outer)
            header_frame.pack(fill='x', pady=(0, 6))
            ttk.Label(header_frame, text='Excel Local Expert v0.3', style='Title.TLabel').pack(side='left')
            ttk.Label(header_frame, text='Dual-Engine AI (OpenAI & Ollama) · Tự chữa lành ô gộp & dòng tiêu đề · Đối soát độc lập', style='Sub.TLabel').pack(side='left', padx=12, pady=6)
            ttk.Button(header_frame, text='⚙ Cài đặt AI & API Key', command=self.open_ai_settings).pack(side='right')

            # Stepper Bar
            self.stepper_frame = ttk.Frame(outer, padding=6)
            self.stepper_frame.pack(fill='x', pady=(0, 8))
            self.step_labels = []
            steps = ['① Nạp Dữ Liệu', '② AI & Thiết Lập', '③ Xem Trước Dự Kiến', '④ Thực Thi & Hoàn Thành']
            for i, st in enumerate(steps):
                lbl = ttk.Label(self.stepper_frame, text=st, style='StepActive.TLabel' if i == 0 else 'StepPending.TLabel')
                lbl.pack(side='left', padx=10)
                self.step_labels.append(lbl)
                if i < len(steps) - 1:
                    ttk.Label(self.stepper_frame, text='➔', foreground='#ced4da').pack(side='left')

            # Data Section
            self.fields = {}
            files = ttk.LabelFrame(outer, text='1. Dữ liệu đầu vào (Tự động nhận diện tiêu đề & ô gộp)', padding=8)
            files.pack(fill='x', pady=(0, 6))
            for row, kind, label in ((0, 'source', 'File nguồn:'), (1, 'reference', 'Danh mục:')):
                ttk.Label(files, text=label).grid(row=row, column=0, sticky='w', padx=(0, 6), pady=3)
                path = tk.StringVar()
                sheet = tk.StringVar()
                header = tk.IntVar(value=1)
                self.fields[kind] = (path, sheet, header)
                ttk.Entry(files, textvariable=path, width=50).grid(row=row, column=1, sticky='ew')
                ttk.Button(files, text='Chọn…', command=lambda k=kind: self.choose(k)).grid(row=row, column=2, padx=4)
                combo = ttk.Combobox(files, textvariable=sheet, state='readonly', width=16)
                combo.grid(row=row, column=3, padx=4)
                combo.bind('<<ComboboxSelected>>', lambda event, k=kind: self.load(k))
                setattr(self, kind + '_sheet', combo)
                ttk.Label(files, text='Dòng tiêu đề:').grid(row=row, column=4, padx=4)
                ttk.Spinbox(files, from_=1, to=20000, textvariable=header, width=4).grid(row=row, column=5)
                ttk.Button(files, text='Đọc', command=lambda k=kind: self.load(k)).grid(row=row, column=6, padx=(4, 2))
                ttk.Button(files, text='🔍 Tự tìm dòng', command=lambda k=kind: self.auto_detect(k)).grid(row=row, column=7, padx=(2, 0))
            files.columnconfigure(1, weight=1)

            # AI & Controls Section
            controls = ttk.LabelFrame(outer, text='2. Lập kế hoạch thao tác & Hỗ trợ AI thông minh', padding=8)
            controls.pack(fill='x', pady=(0, 6))

            engine_box = ttk.Frame(controls)
            engine_box.pack(fill='x', pady=(0, 4))
            ttk.Label(engine_box, text='Động cơ AI:').pack(side='left', padx=(0, 8))
            ttk.Radiobutton(engine_box, text='⚡ Siêu tốc 0-Token (Cục bộ ➔ Cloud khi cần)', variable=self.ai_provider, value='smart_local', command=self.save_engine_choice).pack(side='left', padx=(0, 12))
            ttk.Radiobutton(engine_box, text='🌐 Smart Router (Cloud Cascade)', variable=self.ai_provider, value='smart', command=self.save_engine_choice).pack(side='left', padx=(0, 12))
            ttk.Radiobutton(engine_box, text='🏠 100% Cục bộ (Offline)', variable=self.ai_provider, value='heuristic', command=self.save_engine_choice).pack(side='left', padx=(0, 12))

            op_box = ttk.Frame(controls)
            op_box.pack(fill='x', pady=4)
            self.action = tk.StringVar(value='lookup')
            ttk.Label(op_box, text='Thao tác:').pack(side='left', padx=(0, 8))
            for label, value in (('Đối chiếu (Lookup)', 'lookup'), ('Tổng hợp (Summary)', 'summary'), ('Làm sạch (Clean)', 'clean'), ('🎯 Lọc & Trích xuất (Filter)', 'filter')):
                ttk.Radiobutton(op_box, text=label, variable=self.action, value=value, command=self.labels).pack(side='left', padx=(0, 14))

            config = ttk.Frame(controls)
            config.pack(fill='x', pady=4)
            self.vars = {key: tk.StringVar() for key in ('key', 'value', 'ref_key', 'ref_value')}
            self.combos = {}
            self.config_labels = {}
            for i, (key, label) in enumerate((('key', 'Khóa nguồn'), ('value', 'Cột cần cộng'), ('ref_key', 'Khóa danh mục'), ('ref_value', 'Giá trị trả về'))):
                col = i * 2
                self.config_labels[key] = ttk.Label(config, text=label)
                self.config_labels[key].grid(row=0, column=col, sticky='w', padx=(0, 4))
                combo = ttk.Combobox(config, textvariable=self.vars[key], state='readonly', width=16)
                combo.grid(row=0, column=col + 1, sticky='ew', padx=(0, 8))
                self.combos[key] = combo
                config.columnconfigure(col + 1, weight=1)

            # Quick suggestions
            chip_box = ttk.Frame(controls)
            chip_box.pack(fill='x', pady=(2, 4))
            ttk.Label(chip_box, text='Gợi ý nhanh:', font=('Segoe UI', 8, 'italic'), foreground='#6c757d').pack(side='left', padx=(0, 6))

            def set_prompt_template(text, act):
                self.prompt.delete('1.0', 'end')
                self.prompt.insert('1.0', text)
                self.action.set(act)
                self.labels()

            ttk.Button(chip_box, text='🎯 Lọc Ban Cán Sự (Lớp trưởng, Bí thư...)', command=lambda: set_prompt_template('Lọc danh sách các cán sự, lớp trưởng, tổ trưởng, bí thư', 'filter')).pack(side='left', padx=2)
            ttk.Button(chip_box, text='⭐ Lọc Điểm cao (>= 8.0)', command=lambda: set_prompt_template('Lọc danh sách sinh viên có Điểm lớn hơn hoặc bằng 8.0', 'filter')).pack(side='left', padx=2)
            ttk.Button(chip_box, text='🎓 Đối chiếu Sinh viên', command=lambda: set_prompt_template('Đối chiếu theo Mã sinh viên để lấy thông tin từ danh mục', 'lookup')).pack(side='left', padx=2)
            ttk.Button(chip_box, text='💰 Tổng hợp Học phí/Tiền', command=lambda: set_prompt_template('Tính tổng số tiền hoặc học phí theo từng nhóm/lớp', 'summary')).pack(side='left', padx=2)
            ttk.Button(chip_box, text='🧹 Chuẩn hóa Họ tên & Chuỗi', command=lambda: set_prompt_template('Làm sạch khoảng trắng thừa và chuẩn hóa Unicode cột Họ tên', 'clean')).pack(side='left', padx=2)

            # Prompt
            ttk.Label(controls, text='🔒 Bảo mật Zero-Leakage: Tự động khử định danh PII (CCCD, SĐT, Email) trước khi gọi AI · Dữ liệu ô tính không bao giờ gửi ra ngoài', font=('Segoe UI', 8, 'italic'), foreground='#198754').pack(anchor='w', pady=(2, 2))
            self.prompt = tk.Text(controls, height=2, font=('Segoe UI', 10), wrap='word')
            self.prompt.pack(fill='x', pady=(2, 6))
            self.prompt.insert('1.0', 'Ví dụ: Đối chiếu cột Mã sinh viên với danh mục để lấy Họ tên và Quê quán.')

            actions = ttk.Frame(controls)
            actions.pack(fill='x')
            ttk.Button(actions, text='✨ AI Phân tích & Lập đề xuất', command=self.plan).pack(side='left')
            self.pdf = tk.BooleanVar(value=False)
            ttk.Checkbutton(actions, text='Xuất PDF của sheet kết quả', variable=self.pdf).pack(side='left', padx=12)
            self.auto_heal = tk.BooleanVar(value=True)
            ttk.Checkbutton(actions, text='Tự chữa lành ô gộp & khoảng trắng', variable=self.auto_heal).pack(side='left', padx=8)
            self.clean_output = tk.BooleanVar(value=True)
            ttk.Checkbutton(actions, text='Xuất bảng gọn gàng (Không chèn cột trạng thái thừa)', variable=self.clean_output).pack(side='left', padx=8)
            self.run_button = ttk.Button(actions, text='3. Kiểm tra kết quả dự kiến (Preview)', style='Highlight.TButton', command=self.run)
            self.run_button.pack(side='right')

            # Preview treeview
            preview_frame = ttk.LabelFrame(outer, text='Xem trước bảng dữ liệu (Tối đa 20 dòng / 10 cột)', padding=4)
            preview_frame.pack(fill='both', expand=True, pady=(0, 6))
            self.preview = ttk.Treeview(preview_frame, show='headings', height=6)
            h_scroll = ttk.Scrollbar(preview_frame, orient='horizontal', command=self.preview.xview)
            v_scroll = ttk.Scrollbar(preview_frame, orient='vertical', command=self.preview.yview)
            self.preview.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
            self.preview.grid(row=0, column=0, sticky='nsew')
            h_scroll.grid(row=1, column=0, sticky='ew')
            v_scroll.grid(row=0, column=1, sticky='ns')
            preview_frame.rowconfigure(0, weight=1)
            preview_frame.columnconfigure(0, weight=1)

            # Live Progress & Status
            progress_frame = ttk.Frame(outer)
            progress_frame.pack(fill='x', pady=(0, 4))
            self.status = tk.StringVar(value='Sẵn sàng. Chọn file nguồn để bắt đầu.')
            ttk.Label(progress_frame, textvariable=self.status, font=('Segoe UI', 9, 'bold')).pack(side='left')
            self.stop_button = tk.Button(progress_frame, text='⏹ Dừng lại', command=self.stop_current_work, bg='#dc3545', fg='white', activebackground='#bb2d3b', activeforeground='white', font=('Segoe UI', 9, 'bold'), relief='flat', padx=10, pady=1, cursor='hand2')
            self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=160)
            self.progress.pack(side='right')

            self.log = tk.Text(outer, height=3, font=('Consolas', 9), wrap='word', state='disabled')
            self.log.pack(fill='x', pady=(0, 6))

            # Bottom Actions
            bottom = ttk.Frame(outer)
            bottom.pack(fill='x')
            ttk.Button(bottom, text='📂 Mở workbook kết quả', command=self.open_result).pack(side='left')
            ttk.Button(bottom, text='📁 Mở thư mục kết quả', command=lambda: os.startfile(str(OUTPUT))).pack(side='left', padx=6)
            ttk.Button(bottom, text='⚡ Xử lý hàng loạt (Batch)', command=self.open_batch_dialog).pack(side='left', padx=6)
            ttk.Button(bottom, text='🔄 Làm mới phiên', command=self.reset_session).pack(side='left')
            self.engine_badge = ttk.Label(bottom, text='', foreground='#0d6efd', font=('Segoe UI', 9, 'bold'))
            self.engine_badge.pack(side='right')

            OUTPUT.mkdir(exist_ok=True)
            self.labels()
            self.update_engine_badge()
            window.after(150, self.poll)

        def set_step(self, step):
            self.current_step = step
            for i, lbl in enumerate(self.step_labels):
                if i + 1 < step:
                    lbl.configure(style='StepDone.TLabel')
                elif i + 1 == step:
                    lbl.configure(style='StepActive.TLabel')
                else:
                    lbl.configure(style='StepPending.TLabel')

        def update_engine_badge(self):
            prov = self.ai_provider.get()
            cfg = ai_client.get_config()
            if prov in ('smart_local', 'local_first'):
                status_text = "Động cơ: ⚡ Siêu Tốc 0-Token (Cục bộ 0ms ➔ Cloud khi cần)"
            elif prov == 'heuristic':
                status_text = "Động cơ: 🏠 100% Cục Bộ Nội Bộ (0 Token, Offline 0ms)"
            elif prov == 'smart':
                cands = ai_client.get_active_candidates(cfg)
                if cands:
                    cand_names = [c['name'].split('(')[0].strip() for c in cands]
                    status_text = f"Động cơ: ⚡ Smart Router ({' ➔ '.join(cand_names)} ➔ Ollama)"
                else:
                    status_text = "Động cơ: ⚡ Smart Router (Chưa cấu hình Cloud Key · Tự dùng Ollama & Heuristic)"
            elif prov == 'ollama':
                status_text = "Động cơ: 🏠 Ollama Cục bộ (100% Offline)"
            elif prov in ai_client.PROVIDER_PROFILES:
                prof = ai_client.PROVIDER_PROFILES[prov]
                k = cfg.get('provider_keys', {}).get(prov) or (cfg.get('openai_api_key') if prov == 'openai' else '')
                status_text = f"Động cơ: {prof['name']} · {'Đã có Key (DPAPI)' if k else 'Chưa có Key'}"
            else:
                has_key = bool(cfg.get('openai_api_key') or os.environ.get('OPENAI_API_KEY'))
                status_text = f"Động cơ: OpenAI API ({cfg.get('openai_model', 'gpt-4o-mini')}) · {'Đã có Key (DPAPI)' if has_key else 'Chưa có Key'}"
            self.engine_badge.configure(text=status_text)

        def save_engine_choice(self):
            ai_client.save_config({'provider': self.ai_provider.get()})
            self.update_engine_badge()
            self.note(f"Đã chuyển động cơ AI sang: {self.ai_provider.get().upper()}")

        def open_ai_settings(self):
            cfg = ai_client.get_config()
            prov_keys = cfg.get('provider_keys', {})

            dlg = tk.Toplevel(self.window)
            dlg.title('Cài đặt AI & Smart Router (Luân phiên Model Miễn phí / Trả phí)')
            dlg.geometry('650x550')
            dlg.transient(self.window)
            dlg.grab_set()

            f = ttk.Frame(dlg, padding=16)
            f.pack(fill='both', expand=True)

            ttk.Label(f, text='⚡ Cấu hình Smart Multi-Model Router', font=('Segoe UI', 13, 'bold'), foreground='#0d6efd').pack(anchor='w', pady=(0, 2))
            ttk.Label(f, text='🔒 Khóa API được mã hóa bằng Windows DPAPI an toàn tuyệt đối. Hệ thống tự động luân phiên giữa các model khi hết hạn mức (quota) hoặc sự cố kết nối, và tự động chuyển về Ollama Offline khi mất mạng.', font=('Segoe UI', 8, 'italic'), foreground='#198754', wraplength=610).pack(anchor='w', pady=(0, 8))

            top_frame = ttk.LabelFrame(f, text='Chế độ hoạt động', padding=8)
            top_frame.pack(fill='x', pady=(0, 8))

            row_top = ttk.Frame(top_frame)
            row_top.pack(fill='x', pady=2)
            ttk.Label(row_top, text='Động cơ ưu tiên:').pack(side='left', padx=(0, 8))

            prov_var = tk.StringVar(value=cfg.get('provider', 'smart_local'))
            prov_cb = ttk.Combobox(row_top, textvariable=prov_var, values=['smart_local', 'smart', 'heuristic', 'gemini', 'groq', 'openai', 'openrouter', 'ollama'], state='readonly', width=16)
            prov_cb.pack(side='left', padx=(0, 16))

            fallback_var = tk.BooleanVar(value=cfg.get('auto_fallback', True))
            ttk.Checkbutton(row_top, text='Tự động dự phòng Ollama khi mất mạng', variable=fallback_var).pack(side='left')

            keys_frame = ttk.LabelFrame(f, text='Quản lý Khóa API các Nhà cung cấp (Tự động xoay vòng)', padding=8)
            keys_frame.pack(fill='x', pady=(0, 8))

            grid = ttk.Frame(keys_frame)
            grid.pack(fill='x')

            # 1. Google Gemini
            ttk.Label(grid, text='Google Gemini Key (Free):', font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky='w', pady=3)
            gemini_key_var = tk.StringVar(value=prov_keys.get('gemini', ''))
            gemini_entry = ttk.Entry(grid, textvariable=gemini_key_var, width=34, show='*')
            gemini_entry.grid(row=0, column=1, sticky='ew', pady=3, padx=4)
            ttk.Label(grid, text='(aistudio.google.com)', font=('Segoe UI', 8, 'italic'), foreground='#6c757d').grid(row=0, column=2, sticky='w', padx=4)

            # 2. Groq Cloud
            ttk.Label(grid, text='Groq Cloud Key (Free 70B):', font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky='w', pady=3)
            groq_key_var = tk.StringVar(value=prov_keys.get('groq', ''))
            groq_entry = ttk.Entry(grid, textvariable=groq_key_var, width=34, show='*')
            groq_entry.grid(row=1, column=1, sticky='ew', pady=3, padx=4)
            ttk.Label(grid, text='(console.groq.com)', font=('Segoe UI', 8, 'italic'), foreground='#6c757d').grid(row=1, column=2, sticky='w', padx=4)

            # 3. OpenAI Official
            ttk.Label(grid, text='OpenAI API Key:').grid(row=2, column=0, sticky='w', pady=3)
            openai_key_var = tk.StringVar(value=prov_keys.get('openai', '') or cfg.get('openai_api_key', ''))
            openai_entry = ttk.Entry(grid, textvariable=openai_key_var, width=34, show='*')
            openai_entry.grid(row=2, column=1, sticky='ew', pady=3, padx=4)
            ttk.Label(grid, text='(platform.openai.com)', font=('Segoe UI', 8, 'italic'), foreground='#6c757d').grid(row=2, column=2, sticky='w', padx=4)

            # Toggle show keys
            show_keys_var = tk.BooleanVar(value=False)
            def toggle_show():
                sh = '' if show_keys_var.get() else '*'
                gemini_entry.configure(show=sh)
                groq_entry.configure(show=sh)
                openai_entry.configure(show=sh)
            ttk.Checkbutton(grid, text='Hiện mã khóa API', variable=show_keys_var, command=toggle_show).grid(row=3, column=1, sticky='w', pady=2, padx=4)

            grid.columnconfigure(1, weight=1)

            # Advanced Custom Endpoint (OpenAI-compatible)
            adv_frame = ttk.LabelFrame(f, text='Cổng Tùy chỉnh (OpenAI-compatible / Custom Base URL)', padding=8)
            adv_frame.pack(fill='x', pady=(0, 6))

            adv_grid = ttk.Frame(adv_frame)
            adv_grid.pack(fill='x')

            ttk.Label(adv_grid, text='Base URL:').grid(row=0, column=0, sticky='w', pady=2)
            url_var = tk.StringVar(value=cfg.get('openai_base_url', 'https://api.openai.com/v1'))
            ttk.Entry(adv_grid, textvariable=url_var, width=34).grid(row=0, column=1, sticky='ew', pady=2, padx=4)

            ttk.Label(adv_grid, text='Model:').grid(row=0, column=2, sticky='w', pady=2, padx=(6, 2))
            model_var = tk.StringVar(value=cfg.get('openai_model', 'gpt-4o-mini'))
            ttk.Entry(adv_grid, textvariable=model_var, width=16).grid(row=0, column=3, sticky='ew', pady=2)

            adv_grid.columnconfigure(1, weight=1)

            status_lbl = ttk.Label(f, text='', font=('Segoe UI', 9), wraplength=610)
            status_lbl.pack(fill='x', pady=4)

            btn_box = ttk.Frame(f)
            btn_box.pack(fill='x', pady=(6, 0))

            def test_conn():
                status_lbl.configure(text='Đang kiểm tra kết nối...', foreground='#0d6efd')
                dlg.update()
                p = prov_var.get()
                k = ''
                u = None
                if p == 'gemini':
                    k = gemini_key_var.get().strip()
                    u = ai_client.PROVIDER_PROFILES['gemini']['base_url']
                elif p == 'groq':
                    k = groq_key_var.get().strip()
                    u = ai_client.PROVIDER_PROFILES['groq']['base_url']
                elif p == 'openai':
                    k = openai_key_var.get().strip()
                    u = url_var.get().strip()
                elif p in ('smart', 'smart_local'):
                    k = gemini_key_var.get().strip() or groq_key_var.get().strip() or openai_key_var.get().strip()
                    if gemini_key_var.get().strip():
                        u = ai_client.PROVIDER_PROFILES['gemini']['base_url']
                    elif groq_key_var.get().strip():
                        u = ai_client.PROVIDER_PROFILES['groq']['base_url']
                    else:
                        u = url_var.get().strip()
                ok, msg = ai_client.test_ai_connection(p, api_key=k, base_url=u)
                status_lbl.configure(text=msg, foreground='#198754' if ok else '#dc3545')

            def save_settings():
                g_key = gemini_key_var.get().strip()
                gr_key = groq_key_var.get().strip()
                o_key = openai_key_var.get().strip()
                ai_client.save_config({
                    'provider': prov_var.get(),
                    'auto_fallback': fallback_var.get(),
                    'openai_api_key': o_key,
                    'openai_model': model_var.get().strip(),
                    'openai_base_url': url_var.get().strip(),
                    'provider_keys': {
                        'gemini': g_key,
                        'groq': gr_key,
                        'openai': o_key,
                    }
                })
                self.ai_provider.set(prov_var.get())
                self.update_engine_badge()
                dlg.destroy()
                self.note('Đã lưu cấu hình Smart Router & mã hóa DPAPI an toàn.')

            ttk.Button(btn_box, text='Kiểm tra kết nối', command=test_conn).pack(side='left')
            ttk.Button(btn_box, text='Lưu cấu hình', style='Highlight.TButton', command=save_settings).pack(side='right')
            ttk.Button(btn_box, text='Hủy', command=dlg.destroy).pack(side='right', padx=6)

        def close(self):
            if self.busy:
                messagebox.showinfo('Đang xử lý', 'Hãy đợi tác vụ hoàn thành trước khi đóng để Excel và file tạm được dọn sạch.')
            else:
                self.window.destroy()

        def current(self, kind, snapshot):
            if not snapshot:
                return
            path, sheet, header = self.fields[kind]
            try:
                matches = (Path(path.get()).resolve() == Path(snapshot['path']) and sheet.get() == snapshot['sheet'] and int(header.get()) == snapshot['header_row'])
            except (ValueError, tk.TclError):
                matches = False
            if not matches:
                raise ValueError('File, sheet hoặc hàng tiêu đề đã đổi. Bấm Đọc trước khi tiếp tục.')

        def labels(self):
            action = self.action.get()
            self.config_labels['key'].configure(text={'lookup': 'Khóa nguồn', 'summary': 'Cột nhóm', 'clean': 'Cột làm sạch', 'filter': 'Cột cần lọc'}.get(action, 'Cột nguồn'))
            self.config_labels['value'].configure(text='Cột cần cộng' if action == 'summary' else ('Từ khóa / Tiêu chí lọc' if action == 'filter' else ''))
            self.config_labels['ref_key'].configure(text='Cột trích xuất (Tùy chọn)' if action == 'filter' else 'Khóa danh mục')
            self.config_labels['ref_value'].configure(text='' if action == 'filter' else 'Giá trị trả về')
            for key, combo in self.combos.items():
                enabled = key == 'key' or (key == 'value' and action in ('summary', 'filter')) or (key == 'ref_key' and action in ('lookup', 'filter')) or (key == 'ref_value' and action == 'lookup')
                combo.configure(state='normal' if (key in ('value', 'ref_key') and action == 'filter') else ('readonly' if enabled else 'disabled'))

        def note(self, text):
            self.log.configure(state='normal')
            ts = time.strftime('%H:%M:%S')
            self.log.insert('end', f'[{ts}] {text}\n')
            self.log.see('end')
            self.log.configure(state='disabled')

        def stop_current_work(self):
            if not self.busy:
                return
            self.abort_requested = True
            self.current_worker_id = None
            terminate_active_bridges()
            self.busy = False
            self.progress.stop()
            try:
                self.stop_button.pack_forget()
            except Exception:
                pass
            self.run_button.configure(state='normal')
            self.status.set('⏹ Đã dừng thao tác.')
            self.note('⏹ Đã dừng tác vụ đọc/xử lý file theo yêu cầu của bạn.')

        def work(self, title, function, callback):
            if self.busy:
                messagebox.showinfo('Đang xử lý', 'Hãy đợi thao tác hiện tại hoàn thành hoặc bấm "⏹ Dừng lại".')
                return
            self.busy = True
            self.abort_requested = False
            self.status.set(title)
            self.progress.start()
            try:
                self.stop_button.pack(side='right', padx=(0, 8))
            except Exception:
                pass
            self.run_button.configure(state='disabled')
            worker_id = uuid.uuid4().hex
            self.current_worker_id = worker_id

            def worker():
                try:
                    res = function()
                    if self.abort_requested or self.current_worker_id != worker_id:
                        return
                    self.jobs.put((callback, res, None, worker_id))
                except Exception as error:
                    if self.abort_requested or self.current_worker_id != worker_id:
                        return
                    self.jobs.put((callback, None, str(error), worker_id))

            threading.Thread(target=worker, daemon=True).start()

        def poll(self):
            try:
                item = self.jobs.get_nowait()
                if len(item) == 4:
                    callback, value, error, worker_id = item
                    if self.current_worker_id != worker_id or self.abort_requested:
                        return
                else:
                    callback, value, error = item
                self.busy = False
                self.progress.stop()
                try:
                    self.stop_button.pack_forget()
                except Exception:
                    pass
                self.run_button.configure(state='normal')
                if error:
                    self.status.set('Chưa hoàn thành.')
                    self.note(error)
                    messagebox.showerror('Không thể hoàn thành', error)
                else:
                    callback(value)
            except queue.Empty:
                pass
            except Exception as error:
                self.busy = False
                self.progress.stop()
                try:
                    self.stop_button.pack_forget()
                except Exception:
                    pass
                self.run_button.configure(state='normal')
                self.status.set('Không thể hiển thị kết quả; hãy kiểm tra lại.')
                self.note(str(error))
                messagebox.showerror('Giao diện', str(error))
            finally:
                self.window.after(150, self.poll)

        def choose(self, kind):
            if self.busy:
                return
            path = filedialog.askopenfilename(title='Chọn bảng tính local', filetypes=[('Bảng tính Excel & CSV', '*.xlsx *.xls *.csv'), ('Excel Workbook (*.xlsx)', '*.xlsx'), ('Excel 97-2003 (*.xls)', '*.xls'), ('Tệp CSV (*.csv)', '*.csv'), ('Tất cả tệp', '*.*')])
            if path:
                self.fields[kind][0].set(path)
                self.fields[kind][1].set('')
                self.fields[kind][2].set('1')
                self.load(kind, auto_detect=True)

        def auto_detect(self, kind):
            path = self.fields[kind][0].get()
            if not path or not Path(path).exists():
                messagebox.showinfo('Thiếu file', 'Vui lòng chọn file trước khi dò tìm dòng tiêu đề.')
                return
            self.load(kind, auto_detect=True)

        def load(self, kind, auto_detect=False):
            if self.busy:
                messagebox.showinfo('Đang xử lý', 'Hãy đợi thao tác hiện tại hoàn thành.')
                return
            path, sheet, header = self.fields[kind]
            try:
                header_raw = str(header.get()).strip() if header.get() is not None else '1'
                header_val = 0 if auto_detect else int(header_raw)
                args = (path.get(), sheet.get(), header_val)
            except Exception as error:
                messagebox.showerror('Hàng tiêu đề', str(error))
                return
            if not args[0]:
                return
            setattr(self, kind, None)
            auto_heal_val = self.auto_heal.get()

            def loaded(snapshot):
                failure = snapshot if isinstance(snapshot, HeaderValidationError) else None
                if failure:
                    snapshot = failure.snapshot
                else:
                    setattr(self, kind, snapshot)
                    self.set_step(2)
                self.fields[kind][1].set(snapshot['sheet'])
                getattr(self, kind + '_sheet').configure(values=snapshot['sheets'])
                if snapshot.get('header_row'):
                    self.fields[kind][2].set(str(snapshot['header_row']))
                    if auto_detect and snapshot['header_row'] > 1:
                        self.note(f"🔍 Tự động phát hiện dòng tiêu đề bảng: Dòng {snapshot['header_row']} (đã bỏ qua các dòng banner phía trên).")
                if kind == 'source':
                    self.combos['ref_key'].configure(values=[] if failure else snapshot['headers'])
                cols = ('key', 'value') if kind == 'source' else ('ref_key', 'ref_value')
                for key in cols:
                    self.combos[key].configure(values=[] if failure else snapshot['headers'])
                    if failure:
                        self.vars[key].set('')
                    elif self.vars[key].get() not in snapshot['headers']:
                        self.vars[key].set(snapshot['headers'][0])
                if failure:
                    columns = ['Dòng Excel'] + [column_name(c + 1) for c in range(min(9, len(snapshot['rows'][0])))]
                    start = min(max(0, args[2] - 5), max(0, len(snapshot['rows']) - 20))
                    preview_rows = [[r + 1] + snapshot['rows'][r][:9] for r in range(start, min(start + 20, len(snapshot['rows'])))]
                else:
                    columns = snapshot['headers'][:10]
                    preview_rows = snapshot['data'][:20]
                self.preview.delete(*self.preview.get_children())
                ids = [f'c{i}' for i in range(len(columns))]
                self.preview.configure(columns=ids)
                for column, label in zip(ids, columns):
                    self.preview.heading(column, text=label)
                    self.preview.column(column, width=120, minwidth=70)
                for row in preview_rows:
                    self.preview.insert('', 'end', values=[str(value)[:90] if value is not None else '' for value in row[:10]])
                if failure:
                    self.status.set('Đã đọc file. Chọn đúng sheet và Dòng tên cột rồi bấm Đọc.')
                    self.note(str(failure))
                    messagebox.showerror('Chọn lại dòng tên cột', str(failure))
                    return
                heal_info = ' (Đã tự chữa lành ô gộp/cột khuyết)' if snapshot.get('auto_healed') else ''
                self.status.set(f'Đã đọc {len(snapshot["data"])} dòng · Sheet {snapshot["sheet"]}{heal_info}')
                self.note(f'Đã kiểm tra và đọc {kind}: {len(snapshot["data"])} dòng, {len(snapshot["headers"])} cột{heal_info}.')

            def inspect():
                try:
                    return inspect_file(args[0], args[1], args[2], lenient=True, auto_heal=auto_heal_val)
                except HeaderValidationError as error:
                    return error

            self.work('Đang kiểm tra bảo mật file và nạp dữ liệu…', inspect, loaded)

        def plan(self):
            if self.busy:
                return
            if not self.source:
                messagebox.showinfo('Thiếu nguồn', 'Nạp file nguồn trước.')
                return
            try:
                self.current('source', self.source)
                self.current('reference', self.reference)
            except ValueError as error:
                messagebox.showerror('Cần nạp lại', str(error))
                return
            prompt = self.prompt.get('1.0', 'end').strip()
            source, reference = self.source, self.reference
            prov = self.ai_provider.get()
            cfg = ai_client.get_config()
            if prov == 'openai' and not cfg.get('openai_api_key'):
                self.note('💡 Gợi ý: Chưa có OpenAI API Key. Hệ thống đang dùng Ollama Cục bộ. Để đạt độ chính xác >95% và hiểu tiếng Việt sâu sắc nhất, hãy nhập key tại "⚙ Cài đặt AI & API Key".')
            elif prov == 'smart':
                cands = ai_client.get_active_candidates(cfg)
                if not cands:
                    self.note('💡 Gợi ý Smart Router: Chưa cấu hình Cloud Key (Gemini, Groq, OpenAI). Hệ thống sẽ dùng Ollama Cục bộ hoặc Bộ suy luận ngữ nghĩa.')

            clean_test = ai_client.PrivacyGuard.anonymize_prompt(prompt)
            if clean_test != prompt:
                self.note('🔒 Zero-Leakage: Đã nhận diện và khử định danh dữ liệu nhạy cảm (PII) trong yêu cầu trước khi gửi AI.')

            def proposed(plan):
                if self.prompt.get('1.0', 'end').strip() != prompt:
                    self.status.set('Yêu cầu đã đổi. Bấm AI đề xuất lại.')
                    return
                engine_name = plan.get('engine_used', prov.upper())
                if plan['action'] == 'clarify':
                    self.status.set('Yêu cầu cần bổ sung: ' + plan.get('message', 'Yêu cầu chưa đủ rõ.'))
                    self.note(f'💬 Trợ lý AI ({engine_name}) phản hồi:\n' + plan.get('message', 'Yêu cầu chưa đủ rõ.'))
                    messagebox.showinfo('💬 Trợ lý AI Giải Đáp & Hướng Dẫn', f'[{engine_name}]\n\n' + plan.get('message', 'Yêu cầu chưa đủ rõ.'))
                    return
                self.action.set(plan['action'])
                for key in self.vars:
                    self.vars[key].set(plan.get(key, ''))
                self.labels()
                self.status.set(f'AI ({engine_name}) đã đề xuất kế hoạch. Đang chờ bạn kiểm tra và chấp nhận.')
                action_label = {'lookup': 'Đối chiếu', 'summary': 'Tổng hợp', 'clean': 'Làm sạch', 'filter': 'Lọc & Trích xuất'}[plan['action']]
                self.note(f'[{engine_name}] ' + plan.get('message', '') + f'\nĐề xuất: {action_label}; cột nguồn: {plan["key"]}. Hãy bấm "Kiểm tra kết quả dự kiến".')
                self.set_step(2)
                self.run(plan)

            self.work(f'AI ({prov.upper()}) đang phân tích yêu cầu tiếng Việt…', lambda: ai_plan(prompt, source, reference, allow_clarify=True, provider=prov), proposed)

        def run(self, proposal=None):
            if self.busy:
                return
            if not self.source:
                messagebox.showinfo('Thiếu nguồn', 'Nạp file nguồn trước.')
                return
            plan = proposal or {'action': self.action.get(), **{key: value.get() for key, value in self.vars.items()}, 'message': ''}
            plan['lenient_numbers'] = True
            plan['allow_dates'] = True
            plan['clean_output'] = self.clean_output.get()
            source = self.source
            reference = self.reference if plan['action'] == 'lookup' else None
            export_pdf = self.pdf.get()
            try:
                self.current('source', source)
                self.current('reference', reference)
                preview = preview_plan(source, reference, plan, export_pdf)
            except (ValueError, OSError) as error:
                messagebox.showerror('Tham số', str(error))
                return
            self.set_step(3)
            self.review(source, reference, plan, preview, export_pdf)

        def review(self, source, reference, plan, preview, export_pdf):
            dialog = tk.Toplevel(self.window)
            dialog.title('Kiểm tra kết quả dự kiến · Chưa thực hiện')
            dialog.transient(self.window)
            dialog.geometry(f'{min(1080, self.window.winfo_screenwidth() - 60)}x{min(680, self.window.winfo_screenheight() - 80)}')
            dialog.grab_set()

            frame = ttk.Frame(dialog, padding=16)
            frame.pack(fill='both', expand=True)
            ttk.Label(frame, text='Kiểm tra trước khi chấp nhận và thực thi', font=('Segoe UI', 15, 'bold'), foreground='#0d6efd').pack(anchor='w')

            # Summary Badges
            stats = preview.get('stats', {})
            badge_bar = ttk.Frame(frame)
            badge_bar.pack(fill='x', pady=(6, 4))

            def add_badge(parent, label, val, bg='#e9ecef', fg='#212529'):
                card = tk.Label(parent, text=f"{label}: {val}", font=('Segoe UI', 9, 'bold'), bg=bg, fg=fg, padx=8, pady=3, relief='solid', borderwidth=1)
                card.pack(side='left', padx=(0, 8))

            if plan['action'] == 'lookup':
                add_badge(badge_bar, '✓ Khớp', stats.get('matched', 0), bg='#d1e7dd', fg='#0f5132')
                add_badge(badge_bar, '⚠ Thiếu / Trống', stats.get('missing', 0), bg='#fff3cd', fg='#664d03')
                add_badge(badge_bar, '⚡ Khóa trùng', stats.get('duplicate', 0), bg='#f8d7da', fg='#842029')
            elif plan['action'] == 'summary':
                add_badge(badge_bar, '📊 Số nhóm', stats.get('groups', 0), bg='#cfe2ff', fg='#084298')
                add_badge(badge_bar, '💰 Tổng dự kiến', stats.get('independent_total', 0), bg='#d1e7dd', fg='#0f5132')
            elif plan['action'] == 'clean':
                add_badge(badge_bar, '🧹 Ô làm sạch', stats.get('changed_cells', 0), bg='#d1e7dd', fg='#0f5132')
            elif plan['action'] == 'filter':
                add_badge(badge_bar, '🎯 Dòng trích xuất', f"{stats.get('filtered_rows', 0)} / {stats.get('source_rows', 0)} dòng", bg='#d1e7dd', fg='#0f5132')
                add_badge(badge_bar, '🔍 Tiêu chí', stats.get('criteria', ''), bg='#cfe2ff', fg='#084298')
            add_badge(badge_bar, '🔍 Công thức đối soát', preview.get('formula_checks', 0), bg='#e2e3e5', fg='#41464b')

            labels = {'source_rows': 'Dòng nguồn', 'filtered_rows': 'Dòng trích xuất', 'criteria': 'Tiêu chí lọc', 'matched': 'Khớp', 'missing': 'Thiếu / khóa trống', 'duplicate': 'Khóa trùng trên nguồn', 'reference_duplicate_keys': 'Khóa trùng trong danh mục', 'blank_keys': 'Khóa trống', 'groups': 'Số nhóm', 'independent_total': 'Tổng dự kiến', 'changed_cells': 'Ô được làm sạch'}
            details = [plan.get('message', ''), f'Nguồn: {source["path"]} · Sheet {source["sheet"]}', f'Thao tác: {dict(lookup="Đối chiếu", summary="Tổng hợp", clean="Làm sạch", filter="Lọc & Trích xuất")[plan["action"]]}; Cột nguồn: {plan["key"]}']
            if plan['action'] == 'lookup':
                details.append(f'Danh mục: {plan["ref_key"]}; Trả về: {plan["ref_value"]}')
            if plan['action'] == 'summary':
                details.append('Cột cộng: ' + plan['value'])
            if plan['action'] == 'filter':
                details.append('Tiêu chí lọc: ' + str(plan.get('value', '')))
            details += [f'{labels.get(key, key)}: {value}' for key, value in preview['stats'].items() if key != 'action']
            details += [f'Công thức sẽ đối soát: {preview["formula_checks"]}; Xuất PDF: {"Có" if export_pdf else "Không"}', preview['scope'], preview['status']]

            text = tk.Text(frame, height=5, wrap='word', font=('Segoe UI', 9))
            text.pack(fill='x', pady=6)
            text.insert('1.0', '\n'.join(filter(None, details)))
            text.configure(state='disabled')

            ttk.Label(frame, text='Mẫu kết quả dự kiến (Tối đa 10 dòng đầu):', font=('Segoe UI', 9, 'bold')).pack(anchor='w')
            table_frame = ttk.Frame(frame)
            table_frame.pack(fill='both', expand=True, pady=6)
            ids = [f'c{i}' for i in range(len(preview['columns']))]
            table = ttk.Treeview(table_frame, columns=ids, show='headings', height=8)
            for column, label in zip(ids, preview['columns']):
                table.heading(column, text=label)
                table.column(column, width=150, minwidth=100, stretch=False)
            for row in preview['sample_rows']:
                table.insert('', 'end', values=['' if value is None else str(value)[:200] for value in row])
            horizontal = ttk.Scrollbar(table_frame, orient='horizontal', command=table.xview)
            vertical = ttk.Scrollbar(table_frame, orient='vertical', command=table.yview)
            table.configure(xscrollcommand=horizontal.set, yscrollcommand=vertical.set)
            table.grid(row=0, column=0, sticky='nsew')
            horizontal.grid(row=1, column=0, sticky='ew')
            vertical.grid(row=0, column=1, sticky='ns')
            table_frame.rowconfigure(0, weight=1)
            table_frame.columnconfigure(0, weight=1)

            buttons = ttk.Frame(frame)
            buttons.pack(fill='x', pady=(6, 0))
            def cancel():
                dialog.destroy()
                self.status.set('Chưa chấp nhận. Bạn có thể sửa yêu cầu hoặc các cột.')
                self.set_step(2)
            def accept():
                try:
                    self.current('source', source)
                    self.current('reference', reference)
                    if preview_plan(source, reference, plan, export_pdf)['approval_token'] != preview['approval_token']:
                        raise ValueError('Kết quả dự kiến đã đổi; cần kiểm tra lại.')
                except (ValueError, OSError) as error:
                    dialog.destroy()
                    messagebox.showerror('Bản xem trước hết hiệu lực', str(error))
                    return
                dialog.destroy()
                self.perform(source, reference, plan, export_pdf, preview['approval_token'])
            def apply_batch():
                b_dir = filedialog.askdirectory(title='Chọn thư mục chứa các file cần chạy hàng loạt theo kế hoạch này')
                if not b_dir:
                    return
                dialog.destroy()
                ref_path = reference['path'] if reference else None
                self.start_batch_run(b_dir, ref_path, plan, export_pdf)

            ttk.Button(buttons, name='cancel', text='Chưa chấp nhận · Sửa yêu cầu', command=cancel).pack(side='left')
            ttk.Button(buttons, name='approve', text='Chấp nhận và thực hiện', style='Highlight.TButton', command=accept).pack(side='right')
            ttk.Button(buttons, text='📂 Áp dụng hàng loạt cho Thư mục…', command=apply_batch).pack(side='right', padx=8)
            dialog.protocol('WM_DELETE_WINDOW', cancel)
            self.status.set('Đang chờ chấp nhận; chưa tạo workbook hoặc PDF.')

        def open_batch_dialog(self):
            b_dir = filedialog.askdirectory(title='Chọn thư mục chứa các bảng tính cần xử lý hàng loạt')
            if not b_dir:
                return
            plan = {'action': self.action.get(), **{key: value.get() for key, value in self.vars.items()}, 'message': ''}
            plan['lenient_numbers'] = True
            plan['allow_dates'] = True
            plan['clean_output'] = self.clean_output.get()
            ref_path = self.reference['path'] if self.reference else None
            if not plan.get('key'):
                messagebox.showinfo('Thiếu kế hoạch', 'Hãy nạp file mẫu và chọn thao tác/cột hoặc để AI đề xuất trước, sau đó bấm Chạy hàng loạt.')
                return
            self.start_batch_run(b_dir, ref_path, plan, self.pdf.get())

        def start_batch_run(self, b_dir, ref_path, plan, export_pdf):
            self.set_step(4)
            b_path = Path(b_dir)
            files = [f for f in b_path.iterdir() if f.is_file() and f.suffix.lower() in ('.xlsx', '.xls', '.csv') and not f.name.startswith(('~', '.'))]
            if not files:
                messagebox.showinfo('Không có file', f'Không tìm thấy file XLSX/CSV nào trong thư mục {b_path.name}.')
                return

            b_dlg = tk.Toplevel(self.window)
            b_dlg.title(f'Xử lý Hàng loạt (Batch Mode) · {len(files)} file · 0 Token')
            b_dlg.geometry('820x540')
            b_dlg.transient(self.window)
            b_dlg.grab_set()

            b_frame = ttk.Frame(b_dlg, padding=16)
            b_frame.pack(fill='both', expand=True)

            ttk.Label(b_frame, text=f'📂 Đang xử lý hàng loạt ({len(files)} file)', font=('Segoe UI', 13, 'bold'), foreground='#0d6efd').pack(anchor='w')
            ttk.Label(b_frame, text=f'Áp dụng kế hoạch: {plan["action"].upper()} | Cột: {plan.get("key", "")} | Thư mục: {b_path.name}\n💡 0 AI Token tiêu tốn · Đối soát độc lập 100% từng file.', font=('Segoe UI', 9), foreground='#495057').pack(anchor='w', pady=(0, 6))

            b_progress = ttk.Progressbar(b_frame, maximum=len(files), value=0)
            b_progress.pack(fill='x', pady=6)
            b_status_lbl = ttk.Label(b_frame, text='Đang chuẩn bị...', font=('Segoe UI', 9, 'italic'))
            b_status_lbl.pack(anchor='w', pady=(0, 6))

            tbl_f = ttk.Frame(b_frame)
            tbl_f.pack(fill='both', expand=True, pady=4)
            b_tree = ttk.Treeview(tbl_f, columns=('idx', 'file', 'status', 'rows', 'detail'), show='headings', height=10)
            b_tree.heading('idx', text='#')
            b_tree.column('idx', width=40, anchor='center')
            b_tree.heading('file', text='Tên file')
            b_tree.column('file', width=220)
            b_tree.heading('status', text='Trạng thái')
            b_tree.column('status', width=110, anchor='center')
            b_tree.heading('rows', text='Số dòng')
            b_tree.column('rows', width=80, anchor='center')
            b_tree.heading('detail', text='Chi tiết')
            b_tree.column('detail', width=320)
            b_scroll = ttk.Scrollbar(tbl_f, orient='vertical', command=b_tree.yview)
            b_tree.configure(yscrollcommand=b_scroll.set)
            b_tree.pack(side='left', fill='both', expand=True)
            b_scroll.pack(side='right', fill='y')

            b_btn_box = ttk.Frame(b_frame)
            b_btn_box.pack(fill='x', pady=(8, 0))
            close_btn = ttk.Button(b_btn_box, text='Đóng', state='disabled', command=b_dlg.destroy)
            close_btn.pack(side='right')
            open_folder_btn = ttk.Button(b_btn_box, text='📁 Mở thư mục kết quả', command=lambda: os.startfile(str(OUTPUT)))
            open_folder_btn.pack(side='right', padx=8)

            def on_prog(cur, total, item):
                self.window.after(0, lambda: _update_row(cur, total, item))

            def _update_row(cur, total, item):
                b_progress.configure(value=cur)
                b_status_lbl.configure(text=f"Đang xử lý [{cur}/{total}]: {item['file']}...")
                st_text = '✓ Thành công' if item['status'] == 'success' else '✗ Lỗi'
                det_text = Path(item['output']).name if item['output'] else (item['error'] or '')
                b_tree.insert('', 'end', values=(item['index'], item['file'], st_text, item['rows'], det_text))
                b_tree.yview_moveto(1.0)

            def run_worker():
                try:
                    summary = execute_batch(b_path, reference_path=ref_path, plan=plan, export_pdf=export_pdf, on_progress=on_prog)
                    self.window.after(0, lambda: _done(summary, None))
                except Exception as exc:
                    self.window.after(0, lambda: _done(None, str(exc)))

            def _done(summary, err):
                close_btn.configure(state='normal')
                if err:
                    b_status_lbl.configure(text=f"Lỗi: {err}", foreground='#dc3545')
                    messagebox.showerror('Lỗi xử lý hàng loạt', err)
                else:
                    b_status_lbl.configure(text=f"Hoàn thành: {summary['succeeded']}/{summary['total_files']} file thành công trong {summary['duration_seconds']}s (0 token)!", foreground='#198754')
                    self.note(f"Xử lý hàng loạt hoàn tất: {summary['succeeded']}/{summary['total_files']} file thành công. Báo cáo: {summary['report_file']}")
                    messagebox.showinfo('Hoàn thành hàng loạt', f"Đã xử lý xong {summary['total_files']} file!\nThành công: {summary['succeeded']}\nLỗi: {summary['failed']}\nThời gian: {summary['duration_seconds']}s\n0 token tiêu tốn.")

            threading.Thread(target=run_worker, daemon=True).start()

        def perform(self, source, reference, plan, export_pdf, approval_token):
            self.set_step(4)
            self.note('Đã chấp nhận bản xem trước ' + approval_token[:12] + '. Bắt đầu thực thi và đối soát Excel Desktop.')
            def finished(report):
                self.last_output = report['output']
                self.status.set('Đã tạo workbook và đối soát thành công!')
                self.note(f'Kết quả: {report["output"]}\nCông thức đối soát: {report["formula_checks"]}; file nguồn nguyên vẹn.\nPDF: {report["pdf_layout"]}')
                if report['action'] == 'lookup':
                    self.note(f'Khớp: {report["matched"]}; Thiếu/Trống: {report["missing"]}; Khóa trùng: {report["duplicate"]}.')
                if report['action'] == 'summary':
                    self.note(f'Nhóm: {report["groups"]}; Tổng độc lập: {report["independent_total"]}.')
                if report['action'] == 'clean':
                    self.note(f'Đã làm sạch {report["changed_cells"]} ô trên sheet kết quả.')
                messagebox.showinfo('Hoàn thành', f'Đã xuất bản bảng tính thành công:\n{report["output"]}')
            self.work('Đang tạo workbook, tính lại và đối soát trong Excel…', lambda: execute(source, reference, plan, export_pdf, approval_token=approval_token), finished)

        def open_result(self):
            if self.last_output and Path(self.last_output).exists():
                os.startfile(self.last_output)
            else:
                messagebox.showinfo('Kết quả', 'Chưa có workbook kết quả.')

        def reset_session(self):
            self.source = None
            self.reference = None
            for kind in ('source', 'reference'):
                self.fields[kind][0].set('')
                self.fields[kind][1].set('')
                self.fields[kind][2].set(1)
                getattr(self, kind + '_sheet').configure(values=[])
            for key in self.vars:
                self.vars[key].set('')
                self.combos[key].configure(values=[])
            self.preview.delete(*self.preview.get_children())
            self.preview.configure(columns=[])
            self.status.set('Đã làm mới phiên làm việc. Hãy chọn file để bắt đầu.')
            self.set_step(1)
            self.note('Đã làm mới phiên làm việc.')

    window = tk.Tk()
    window.controller = App(window)
    if smoke:
        window.update()
        outer = window.winfo_children()[0]
        clipped = [str(widget) for widget in outer.winfo_children() if widget.winfo_height() <= 1 or widget.winfo_y() + widget.winfo_height() > outer.winfo_height()]
        print(json.dumps({'title': window.title(), 'width': window.winfo_width(), 'height': window.winfo_height(), 'clipped_sections': clipped, 'status': 'GUI constructed'}, ensure_ascii=False))
        window.after(300, window.destroy)
    window.mainloop()


def main():
    cleanup_stale_work_dirs()
    parser = argparse.ArgumentParser(description='Excel Local Expert')
    parser.add_argument('--source')
    parser.add_argument('--sheet', default='')
    parser.add_argument('--reference')
    parser.add_argument('--ref-sheet', default='')
    parser.add_argument('--ref-header-row', type=int, default=1)
    parser.add_argument('--header-row', type=int, default=1)
    parser.add_argument('--plan', help='JSON file with fixed action/column names')
    parser.add_argument('--pdf', action='store_true')
    parser.add_argument('--ai-prompt')
    parser.add_argument('--provider', choices=['smart_local', 'smart', 'heuristic', 'gemini', 'groq', 'openai', 'openrouter', 'ollama'], help='Chọn AI Provider')
    parser.add_argument('--api-key', help='OpenAI API Key')
    parser.add_argument('--approve', metavar='TOKEN', help='Token của bản xem trước đã được người dùng kiểm tra và chấp nhận')
    parser.add_argument('--full-audit-columns', action='store_true', help='Chèn thêm các cột trạng thái KHOP và khóa chuẩn hóa vào sheet kết quả')
    parser.add_argument('--batch-dir', help='Thư mục chứa các bảng tính nguồn để xử lý hàng loạt theo kế hoạch')
    parser.add_argument('--pattern', default=None, help='Mẫu lọc tên file trong thư mục batch (VD: *.xlsx, lop_*.xlsx)')
    parser.add_argument('--gui-smoke', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.batch_dir:
        if not args.plan and not args.ai_prompt:
            parser.error('--batch-dir cần kèm theo --plan hoặc --ai-prompt để xác định kế hoạch xử lý.')
        reference = inspect_file(args.reference, args.ref_sheet, args.ref_header_row, lenient=True, auto_heal=True) if args.reference else None
        if args.plan:
            plan = json.loads(Path(args.plan).read_text(encoding='utf-8-sig'))
        else:
            b_src = Path(args.batch_dir).resolve(strict=True)
            valid_exts = ('.xlsx', '.xls', '.csv')
            candidates = [f for f in sorted(b_src.iterdir()) if f.is_file() and f.suffix.lower() in valid_exts and not f.name.startswith(('~', '.'))]
            if not candidates:
                parser.error(f'Không tìm thấy file bảng tính nào trong thư mục {b_src}')
            sample_src = inspect_file(candidates[0], args.sheet, args.header_row, lenient=True, auto_heal=True)
            plan = ai_plan(args.ai_prompt, sample_src, reference, provider=args.provider, api_key=args.api_key)
        batch_report = execute_batch(args.batch_dir, reference_path=args.reference, plan=plan, export_pdf=args.pdf, file_pattern=args.pattern)
        print(json.dumps(batch_report, ensure_ascii=False, indent=2))
        return

    if not args.source:
        gui(args.gui_smoke)
        return
    if not args.plan and not args.ai_prompt:
        parser.error('--source cần --plan hoặc --ai-prompt')
    source = inspect_file(args.source, args.sheet, args.header_row, lenient=True, auto_heal=True)
    reference = inspect_file(args.reference, args.ref_sheet, args.ref_header_row, lenient=True, auto_heal=True) if args.reference else None
    plan = ai_plan(args.ai_prompt, source, reference, provider=args.provider, api_key=args.api_key) if args.ai_prompt else json.loads(Path(args.plan).read_text(encoding='utf-8-sig'))
    if 'clean_output' not in plan:
        plan['clean_output'] = not args.full_audit_columns
    preview = preview_plan(source, reference, plan, args.pdf)
    if not args.approve:
        print(json.dumps({'status': 'awaiting_approval', 'preview': preview}, ensure_ascii=False, indent=2))
        return
    if args.approve != preview['approval_token']:
        parser.error('Bản xem trước đã đổi hoặc token không khớp; hãy kiểm tra lại trước khi chấp nhận.')
    print(json.dumps(execute(source, reference, plan, args.pdf, approval_token=args.approve), ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
