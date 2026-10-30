"""Round 26: fuzz the pure-Python XLSX reading paths.

load_xlsx_fast (openpyxl-based lenient reader) contract: return a snapshot
dict or None, never raise. read_xml contract: parse or raise ValueError on
DTD/entity content. Plus one end-to-end lenient inspect on a real workbook.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import datetime
import tempfile
import zipfile
from pathlib import Path

import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def run():
    try:
        import openpyxl
    except ImportError:
        # load_xlsx_fast itself degrades to None without openpyxl; the fast
        # path is optional, so this file SKIPs instead of failing (round 4 rule).
        print('[SKIP] openpyxl chua duoc cai - bo qua fuzz duong doc XLSX nhanh.')
        return
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        d = Path(td)

        fake = d / 'fake.xlsx'
        fake.write_text('this is not a zip at all', encoding='utf-8')
        ok('X-01 file khong phai ZIP -> None, khong nem loi', app.load_xlsx_fast(fake) is None)

        notwb = d / 'notwb.xlsx'
        with zipfile.ZipFile(notwb, 'w') as z:
            z.writestr('random/part.xml', '<x/>')
        ok('X-02 ZIP khong co cau truc workbook -> None', app.load_xlsx_fast(notwb) is None)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Data'
        ws.append(['MSSV', 'Ho ten', 'Ngay', 'Diem', 'Dat'])
        ws.append(['001', 'An', datetime.date(2024, 1, 15), 8.5, True])
        ws.append(['002', 'Binh', datetime.datetime(2024, 2, 1, 10, 30, 0), None, False])
        ws2 = wb.create_sheet('Phu')
        ws2.append(['X'])
        good = d / 'good.xlsx'
        wb.save(good)

        snap = app.load_xlsx_fast(good)
        ok('X-03 file hop le doc duoc, dung sheet dau', snap is not None and snap['sheet'] == 'Data' and len(snap['rows']) == 3)
        ok('X-04 ngay thang dang date -> YYYY-MM-DD', snap['rows'][1][2] == '2024-01-15')
        ok('X-05 datetime co gio -> YYYY-MM-DD HH:MM:SS', snap['rows'][2][2] == '2024-02-01 10:30:00')
        ok('X-06 boolean giu nguyen, o trong la None', snap['rows'][1][4] is True and snap['rows'][2][3] is None)

        snap2 = app.load_xlsx_fast(good, sheet_name='Phu')
        ok('X-07 sheet_name ton tai -> doc dung sheet do', snap2['sheet'] == 'Phu' and snap2['rows'] == [['X']])
        snap3 = app.load_xlsx_fast(good, sheet_name='KhongCo')
        ok('X-08 sheet_name khong ton tai -> lui ve sheet dau', snap3['sheet'] == 'Data')

        wb2 = openpyxl.Workbook()
        wb2.active.title = 'Empty'
        empty = d / 'empty.xlsx'
        wb2.save(empty)
        snap4 = app.load_xlsx_fast(empty)
        ok('X-09 sheet rong -> dict voi rows rong (khong None, khong loi)', snap4 is not None and snap4['rows'] == [])

        banner = d / 'banner.xlsx'
        wb3 = openpyxl.Workbook()
        ws3 = wb3.active
        ws3.title = 'DS'
        ws3.append(['TRUONG DAI HOC TON DUC THANG'])
        ws3.append(['DANH SACH SINH VIEN'])
        ws3.append(['MSSV', 'Ho ten'])
        ws3.append(['001', 'An'])
        wb3.save(banner)
        insp = app.inspect_file(banner, header_row='auto', lenient=True)
        ok('X-10 inspect lenient + auto header tren XLSX that: nhan dung header dong 3',
           insp['headers'] == ['MSSV', 'Ho ten'] and insp['data'] == [['001', 'An']])

        dtd = d / 'dtd.xlsx'
        with zipfile.ZipFile(dtd, 'w') as z:
            z.writestr('xl/workbook.xml', '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e "boom">]><workbook/>')
        with zipfile.ZipFile(dtd) as z:
            try:
                app.read_xml(z, 'xl/workbook.xml')
                ok('X-11 read_xml chan DTD/entity', False)
            except ValueError:
                ok('X-11 read_xml chan DTD/entity', True)

        deep = d / 'deep.xlsx'
        with zipfile.ZipFile(deep, 'w') as z:
            z.writestr('a.xml', '<r>' + '<n>' * 200 + 'x' + '</n>' * 200 + '</r>')
        with zipfile.ZipFile(deep) as z:
            root = app.read_xml(z, 'a.xml')
        ok('X-12 XML long nhau 200 tang van doc duoc', root.tag == 'r')

    print(f'\nAll {len(CHECKS)} xlsx-fuzz checks passed.')

if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    run()
