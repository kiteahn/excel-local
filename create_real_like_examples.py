"""Generate synthetic Vietnamese office workbooks for Q4 (no real user data).

Every value is deterministic and invented. Files include common real-world
warts: administrative banners, merged title cells, zero-padded codes, money
stored as text with dot thousands, mixed date formats, NFC/NFD names, a total
row, an error cell and duplicate keys. They live in examples/real_like/.
"""
from datetime import date, timedelta
import unicodedata
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'examples' / 'real_like'

SURNAMES = ['Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Vũ', 'Đặng', 'Bùi']
GIVEN = ['Văn An', 'Thị Bình', 'Minh Châu', 'Quốc Dũng', 'Thu Hà', 'Hải Long', 'Ngọc Mai', 'Thanh Nam']
DEPTS = ['Kế toán', 'Kinh doanh', 'Kỹ thuật', 'Hành chính']
CLASSES = ['CNTT01', 'KTPM02', 'HTTT03']
GROUPS = ['Đồ uống', 'Bánh kẹo', 'Gia vị', 'Văn phòng phẩm']
PRODUCTS = ['Cà phê rang', 'Trà xanh', 'Bánh quy', 'Nước mắm', 'Bút bi', 'Giấy A4', 'Mì gói', 'Sữa hộp']


def nfd(text):
    return unicodedata.normalize('NFD', text)


def money(value):
    return f'{value:,}'.replace(',', '.')


def banner(ws, title, subtitle, last_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.cell(1, 1, title).font = Font(bold=True, size=14)
    ws.cell(1, 1).alignment = Alignment(horizontal='center')
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws.cell(2, 1, subtitle).font = Font(italic=True, size=10)
    ws.cell(2, 1).alignment = Alignment(horizontal='center')


def write_table(ws, header_row, headers, rows):
    for c, name in enumerate(headers, 1):
        cell = ws.cell(header_row, c, name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', fgColor='DDEBF7')
    for r, row in enumerate(rows, header_row + 1):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
    for c, name in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c)].width = min(max(12, len(name) + 4), 34)


def make_students():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Danh sách'
    banner(ws, 'TRƯỜNG ĐẠI HỌC MẪU', 'PHÒNG CÔNG TÁC SINH VIÊN — NĂM HỌC 2025-2026', 6)
    headers = ['STT', 'MSSV', 'Họ và tên', 'Lớp', 'Ngày sinh', 'Điểm TB']
    rows = []
    for i in range(1, 121):
        name = f"{SURNAMES[i % len(SURNAMES)]} {GIVEN[i % len(GIVEN)]}"
        if i % 7 == 0:
            name = nfd(name)
        birth = [date(2004, 1, 1) + timedelta(days=i * 13),
                 f'{1 + i % 28:02d}/{1 + i % 12:02d}/2004',
                 f'2004-{1 + i % 12:02d}-{1 + i % 28:02d}'][i % 3]
        rows.append([i, f'{i:07d}', name, CLASSES[i % 3], birth, round(5 + (i % 50) / 10, 1)])
    write_table(ws, 4, headers, rows)
    wb.save(OUT / 'sinh_vien.xlsx')


def make_payroll():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Bảng lương'
    banner(ws, 'CÔNG TY TNHH MẪU VIỆT NAM', 'BẢNG LƯƠNG THÁNG 10/2026 — ĐƠN VỊ: ĐỒNG', 7)
    headers = ['Mã NV', 'Họ tên', 'Phòng ban', 'Ngày công', 'Lương cơ bản', 'Phụ cấp', 'Thực lĩnh']
    rows = []
    for i in range(1, 81):
        base = 8_000_000 + (i % 8) * 750_000
        allowance = 300_000 + (i % 5) * 100_000
        net = base + allowance
        rows.append([f'NV{i:04d}', nfd(f'{SURNAMES[i % 8]} {GIVEN[i % 8]}') if i % 9 == 0 else f'{SURNAMES[i % 8]} {GIVEN[i % 8]}',
                     DEPTS[i % 4], 18 + i % 6, money(base), money(allowance), money(net)])
    write_table(ws, 3, headers, rows)
    wb.save(OUT / 'bang_luong.xlsx')


def make_warehouse():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Kho'
    banner(ws, 'KHO HÀNG MẪU', 'SỔ NHẬP XUẤT TỒN — THÁNG 10/2026', 7)
    headers = ['Mã hàng', 'Tên hàng', 'Nhóm hàng', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Ngày nhập']
    rows = []
    total = 0
    for i in range(1, 101):
        qty = 1 + i % 40
        price = 5_000 + (i % 20) * 1_500
        amount = qty * price
        total += amount
        rows.append([f'SP{i:04d}', PRODUCTS[i % len(PRODUCTS)], GROUPS[i % 4], qty,
                     money(price), money(amount), f'{1 + i % 28:02d}/{1 + i % 12:02d}/2026'])
    rows.append(['', 'TỔNG CỘNG', '', sum(1 + i % 40 for i in range(1, 101)), '', money(total), ''])
    write_table(ws, 3, headers, rows)
    wb.save(OUT / 'kho_hang.xlsx')


def make_attendance():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Điểm danh'
    banner(ws, 'LỚP CNTT01 — HỌC PHẦN CƠ SỞ DỮ LIỆU', 'ĐIỂM DANH BUỔI 1–10', 6)
    headers = ['Mã SV', 'Họ tên', 'Lớp', 'Buổi', 'Trạng thái', 'Có mặt', 'Ghi chú']
    rows = []
    for i in range(1, 101):
        rows.append([f'{i:07d}', f'{SURNAMES[i % 8]} {GIVEN[i % 8]}', CLASSES[i % 3],
                     f'Buổi {1 + i % 10}', 'Vắng' if i % 7 == 0 else 'Có mặt',
                     0 if i % 7 == 0 else 1,
                     '  ' + ('Có phép' if i % 2 else '') + '  '])
    write_table(ws, 3, headers, rows)
    wb.save(OUT / 'diem_danh.xlsx')


def make_cashbook():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sổ quỹ'
    banner(ws, 'VĂN PHÒNG ĐOÀN THANH NIÊN', 'SỔ QUỸ TIỀN MẶT — QUÝ IV/2026', 6)
    headers = ['Ngày', 'Số phiếu', 'Diễn giải', 'Thu', 'Chi', 'Số dư']
    rows = []
    balance = 2_000_000
    for i in range(1, 61):
        thu = 500_000 + (i % 4) * 250_000 if i % 3 else 0
        chi = 200_000 + (i % 5) * 100_000 if i % 3 else 0
        balance += thu - chi
        rows.append([f'{1 + i % 28:02d}/{1 + i % 12:02d}/2026', f'PT{i:03d}',
                     'Thu quỹ' if thu else 'Chi hoạt động', money(thu) if thu else '',
                     money(chi) if chi else '', money(balance)])
    rows[17][5] = '#N/A'
    rows.append(['', '', 'TỔNG CỘNG', money(0), money(0), money(balance)])
    write_table(ws, 3, headers, rows)
    wb.save(OUT / 'so_quy.xlsx')


def make_orders():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Đơn hàng'
    banner(ws, 'CỬA HÀNG MẪU', 'DANH SÁCH ĐƠN HÀNG THÁNG 10/2026', 7)
    headers = ['Mã đơn', 'Ngày đặt', 'Khách hàng', 'Mã hàng', 'Số lượng', 'Đơn giá', 'Thành tiền']
    rows = []
    for i in range(1, 121):
        qty = 1 + i % 12
        price = 10_000 + (i % 15) * 2_000
        rows.append([f'DH{i // 3 + 1:04d}', f'{1 + i % 28:02d}/{1 + i % 12:02d}/2026',
                     f'Khách {chr(65 + i % 8)}', f'SP{i % 100 + 1:04d}', qty, money(price), money(qty * price)])
    write_table(ws, 3, headers, rows)
    wb.save(OUT / 'don_hang.xlsx')


def make_catalogs():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Danh mục hàng'
    headers = ['Mã hàng', 'Tên hàng', 'Nhóm hàng', 'Đơn giá']
    rows = [[f'SP{i:04d}', PRODUCTS[i % len(PRODUCTS)], GROUPS[i % 4], money(10_000 + i * 500)]
            for i in range(1, 121)]
    write_table(ws, 1, headers, rows)
    wb.save(OUT / 'danh_muc_hang.xlsx')

    wb = Workbook()
    ws = wb.active
    ws.title = 'Nhân viên'
    headers = ['Mã NV', 'Họ tên chuẩn', 'Phòng ban chuẩn']
    rows = [[f'NV{i:04d}', f'{SURNAMES[i % 8]} {GIVEN[i % 8]}', DEPTS[i % 4]]
            for i in range(1, 121)]
    write_table(ws, 1, headers, rows)
    wb.save(OUT / 'danh_muc_nhan_vien.xlsx')

    wb = Workbook()
    ws = wb.active
    ws.title = 'Lớp'
    headers = ['Mã lớp', 'Tên lớp', 'Giảng viên']
    rows = [[CLASSES[i], f'Lớp {CLASSES[i]} — Khoa CNTT', f'GV {i + 1:02d}'] for i in range(3)]
    write_table(ws, 1, headers, rows)
    wb.save(OUT / 'danh_muc_lop.xlsx')


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    make_students()
    make_payroll()
    make_warehouse()
    make_attendance()
    make_cashbook()
    make_orders()
    make_catalogs()
    print(f'Generated {len(list(OUT.glob("*.xlsx")))} synthetic workbooks in {OUT}')


if __name__ == '__main__':
    main()
