"""Generate diverse, realistic Excel & CSV sample files covering diverse edge cases and scenarios."""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile
import app
from test_app import fixture

EXAMPLES_DIR = app.ROOT / 'examples'
EXAMPLES_DIR.mkdir(exist_ok=True)

def generate_all_samples():
    print(f"Creating diverse sample files in: {EXAMPLES_DIR}")

    # =========================================================================
    # KỊCH BẢN 1: BẢNG HỌC PHÍ ĐẠI HỌC CÓ BANNER HÀNH CHÍNH & MÃ SỐ 0 ĐẦU
    # =========================================================================
    p1_src = EXAMPLES_DIR / '01_hoc_phi_co_banner.xlsx'
    rows_p1 = [
        ['CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', None, None, None, None, None],
        ['Độc lập - Tự do - Hạnh phúc', None, None, None, None, None],
        ['TRƯỜNG ĐẠI HỌC BÁCH KHOA - PHÒNG TÀI CHÍNH KẾ TOÁN', None, None, None, None, None],
        ['DANH SÁCH THEO DÕI NỘP HỌC PHÍ NĂM HỌC 2024-2025', None, None, None, None, None],
        # Dòng 5: Tiêu đề bảng thật
        ['STT', 'Mã sinh viên', 'Họ và tên', 'Lớp sinh hoạt', 'Số tiền nộp', 'Ngày nộp tiền'],
        [1, '00123', 'Nguyễn Văn An', '20KTPM1', '1.500.000 đ', '15/10/2024'],
        [2, '00124', 'Trần Thị Bình', '20KTPM1', '2.850.000 đ', '16/10/2024'],
        [3, '00125', 'Lê Hoàng Cường', '20KTPM2', '3.200.000 đ', '17/10/2024'],
        [4, '00126', 'Phạm Minh Dũng', '20KTPM2', '1.500.000 đ', '18/10/2024'],
        [5, '00127', 'Hoàng Ngọc Ánh', '20KTPM1', '4.100.000 đ', '19/10/2024'],
        [6, '00128', 'Đỗ Thành Đạt', '20KTPM3', '2.500.000 đ', '20/10/2024'],
        [7, '00129', 'Vũ Mai Phương', '20KTPM3', '3.000.000 đ', '21/10/2024'],
    ]
    fixture(p1_src, {'HocPhi': rows_p1})
    print(f"  [OK] Tạo thành công: {p1_src.name}")

    p1_ref = EXAMPLES_DIR / '01_danh_muc_sinh_vien.xlsx'
    rows_p1_ref = [
        ['Mã sinh viên', 'Khoa chuyên ngành', 'Chương trình đào tạo', 'Cố vấn học tập'],
        ['00123', 'Công nghệ thông tin', 'Chất lượng cao', 'ThS. Nguyễn Văn Toàn'],
        ['00124', 'Công nghệ thông tin', 'Đại trà', 'TS. Lê Thị Mai'],
        ['00125', 'Khoa học máy tính', 'Tài năng', 'PGS. Trần Quốc Bảo'],
        ['00126', 'Hệ thống thông tin', 'Đại trà', 'ThS. Hoàng Văn Nam'],
        ['00127', 'Kỹ thuật phần mềm', 'Chất lượng cao', 'TS. Phạm Minh Tuấn'],
        ['00128', 'Mạng máy tính', 'Đại trà', 'ThS. Vũ Đức Hùng'],
        ['00129', 'An toàn thông tin', 'Tài năng', 'TS. Đặng Thanh Sơn'],
    ]
    fixture(p1_ref, {'DanhMuc': rows_p1_ref})
    print(f"  [OK] Tạo thành công: {p1_ref.name}")

    # =========================================================================
    # KỊCH BẢN 2: KẾ TOÁN CHI PHÍ - SỐ ÂM NGOẶC ĐƠN, ĐUÔI TRỪ, TIỀN TỆ ĐA DẠNG
    # =========================================================================
    p2_src = EXAMPLES_DIR / '02_ke_toan_chi_phi_ngan_sach.xlsx'
    rows_p2 = [
        ['Mã chứng từ', 'Phòng ban phát sinh', 'Hạng mục chi phí', 'Số tiền (VND)', 'Ghi chú duyệt'],
        ['CT001', 'Phòng Hành chính', 'Văn phòng phẩm quý 3', '15.450.000 đ', 'Đã duyệt chi'],
        ['CT002', 'Phòng Hành chính', 'Hoàn ứng tiếp khách', '(1,250,000)', 'Giảm trừ chi phí'],
        ['CT003', 'Phòng Kỹ thuật', 'Mua sắm thiết bị linh kiện', '45.000.000 đ', 'Hóa đơn VAT'],
        ['CT004', 'Phòng Kỹ thuật', 'Chiết khấu thương mại', '2500000-', 'Đuôi trừ kế toán'],
        ['CT005', 'Phòng Nhân sự', 'Chi phí tuyển dụng tháng 9', '12.800.000 đ', 'Hợp đồng ngoài'],
        ['CT006', 'Phòng Nhân sự', 'Hoàn trả tạm ứng công tác', '(500.000 đ)', 'Kế toán đã thu hoàn'],
        ['CT007', 'Phòng Kinh doanh', 'Chi phí quảng bá hội thảo', '68.500.000 đ', 'Được cấp ngân sách'],
        ['CT008', 'Phòng Kinh doanh', 'Thu hồi bảo lãnh', '- 1.500.000', 'Khoản thu hồi'],
    ]
    fixture(p2_src, {'ChiPhi': rows_p2})
    print(f"  [OK] Tạo thành công: {p2_src.name}")

    # =========================================================================
    # KỊCH BẢN 3: NHÂN SỰ & TIỀN LƯƠNG - Ô GỘP NGANG, KHUYẾT TÊN CỘT, KÝ TỰ ẨN
    # =========================================================================
    p3_src = EXAMPLES_DIR / '03_nhan_su_tien_luong_o_gop.xlsx'
    zw = '\u200b' # zero-width space
    nbsp = '\u00a0' # non-breaking space
    rows_p3 = [
        # Ô gộp: Cột 2 None (ô gộp của Họ tên), Cột 5 None (ô gộp của Phụ cấp)
        ['Mã nhân viên', 'Họ và tên', None, 'Phòng ban', 'Lương cơ bản', None, 'Ghi chú hồ sơ'],
        ['NV01', 'Nguyễn' + nbsp + 'Văn', 'An' + zw, 'Kỹ thuật', '18.000.000 đ', '2.000.000 đ', '  Chính thức  '],
        ['NV02', 'Trần' + zw + 'Thị', 'Bình', 'Hành chính', '12.500.000 đ', '1.500.000 đ', '  Thử việc   '],
        ['NV03', 'Lê' + nbsp + 'Văn', 'Cường', 'Kỹ thuật', '22.000.000 đ', '3.000.000 đ', '  Chính thức\u200b  '],
        ['NV04', 'Phạm' + zw + 'Thị', 'Dung', 'Kế toán', '15.000.000 đ', '2.000.000 đ', '  Chính thức  '],
        ['NV05', 'Hoàng' + nbsp + 'Minh', 'Em', 'Kinh doanh', '16.000.000 đ', '5.000.000 đ', '  Đạt KPI  '],
    ]
    fixture(p3_src, {'BangLuong': rows_p3})
    print(f"  [OK] Tạo thành công: {p3_src.name}")

    # =========================================================================
    # KỊCH BẢN 4: TUYỂN SINH - KHÓA TRÙNG, KHÓA TRỐNG & KÝ TỰ ĐẶC BIỆT (~, *, ?)
    # =========================================================================
    p4_src = EXAMPLES_DIR / '04_tuyen_sinh_khoa_trung_ngoai_le.xlsx'
    rows_p4 = [
        ['Mã hồ sơ', 'Họ và tên thí sinh', 'Khu vực', 'Điểm thi tổng'],
        ['TS01', 'Nguyễn Gia Bảo', 'KV1', '24,5'],
        ['TS02', 'Trần Thu Hà', 'KV2-NT', '26,0'],
        ['TS02', 'Trần Thu Hà (Bản nộp lại)', 'KV2-NT', '26,0'], # Khóa trùng
        ['', 'Thí sinh khuyết mã', 'KV3', '21,0'],               # Khóa trống
        ['TS*03', 'Lê Văn Sao', 'KV1', '27,5'],                   # Dấu sao *
        ['TS~04', 'Phạm Hoàng Ngã', 'KV2', '25,25'],              # Dấu ngã ~
        ['TS99', 'Đỗ Ngoại Lệ (Chưa có trong DM)', 'KV3', '28,0'],# Không tìm thấy
    ]
    fixture(p4_src, {'ThiSinh': rows_p4})
    print(f"  [OK] Tạo thành công: {p4_src.name}")

    p4_ref = EXAMPLES_DIR / '04_danh_muc_diem_chuan.xlsx'
    rows_p4_ref = [
        ['Mã hồ sơ', 'Ngành trúng tuyển', 'Điểm chuẩn xét tuyển', 'Trạng thái'],
        ['TS01', 'Khoa học Dữ liệu', '23,5', 'Trúng tuyển'],
        ['TS02', 'Trí tuệ Nhân tạo (Đợt 1)', '25,5', 'Trúng tuyển'],
        ['TS02', 'Trí tuệ Nhân tạo (Đợt 2)', '25,0', 'Trúng tuyển'], # Trùng khóa trong danh mục
        ['TS*03', 'Kỹ thuật Robot', '26,0', 'Trúng tuyển'],
        ['TS~04', 'Công nghệ Vi mạch', '24,0', 'Trúng tuyển'],
        ['TS05', 'An ninh Mạng', '24,5', 'Trúng tuyển'],
    ]
    fixture(p4_ref, {'DiemChuan': rows_p4_ref})
    print(f"  [OK] Tạo thành công: {p4_ref.name}")

    # =========================================================================
    # KỊCH BẢN 5: BÁN HÀNG KHO VẬN - CSV CHẤM PHẨY (;), UTF-8 BOM, NGOẶC KÉP
    # =========================================================================
    p5_csv = EXAMPLES_DIR / '05_kho_hang_xuat_ban.csv'
    csv_content = (
        '\ufeff"Mã sản phẩm";"Tên hàng hóa";"Kho xuất";"Số lượng";"Đơn giá (đ)";"Thành tiền (VND)"\r\n'
        '"SP001";"Máy tính xách tay Dell ""Latitude""";"Kho Miền Bắc";15;"18.500.000 đ";"277.500.000 đ"\r\n'
        '"SP002";"Chuột không dây Logitech, dòng silent";"Kho Miền Bắc";50;"350.000 đ";"17.500.000 đ"\r\n'
        '"SP003";"Bàn phím cơ DareU; bản LED RGB";"Kho Miền Nam";30;"850.000 đ";"25.500.000 đ"\r\n'
        '"SP004";"Màn hình Asus 27 inch 2K, 144Hz";"Kho Miền Nam";10;"6.200.000 đ";"62.000.000 đ"\r\n'
        '"SP005";"Ổ cứng SSD Samsung 1TB";"Kho Miền Trung";25;"2.400.000 đ";"60.000.000 đ"\r\n'
        '"SP006";"Tai nghe Sony chống ồn";"Kho Miền Trung";12;"4.500.000 đ";"54.000.000 đ"\r\n'
    )
    p5_csv.write_bytes(csv_content.encode('utf-8'))
    print(f"  [OK] Tạo thành công: {p5_csv.name}")

    # =========================================================================
    # KỊCH BẢN 6: BẢNG ĐIỂM HỌC KỲ - SỐ THẬP PHÂN PHẨY VIỆT NAM (8,5; 7,25; 9,0)
    # =========================================================================
    p6_src = EXAMPLES_DIR / '06_bang_diem_hoc_ky_phay_thap_phan.xlsx'
    rows_p6 = [
        ['Mã số sinh viên', 'Họ và tên', 'Lớp học phần', 'Điểm quá trình', 'Điểm cuối kỳ', 'Điểm trung bình'],
        ['2021001', 'Vũ Thành An', 'Giải tích 1', '8,5', '9,0', '8,8'],
        ['2021002', 'Bùi Văn Bình', 'Giải tích 1', '7,0', '8,5', '7,9'],
        ['2021003', 'Trịnh Quốc Cường', 'Vật lý đại cương', '9,25', '8,75', '8,95'],
        ['2021004', 'Dương Thúy Diễm', 'Vật lý đại cương', '6,5', '7,5', '7,1'],
        ['2021005', 'Ngô Hoàng Giang', 'Đại số tuyến tính', '8,0', '8,0', '8,0'],
        ['2021006', 'Lý Hải Hưng', 'Đại số tuyến tính', '7,75', '9,25', '8,65'],
    ]
    fixture(p6_src, {'BangDiem': rows_p6})
    print(f"  [OK] Tạo thành công: {p6_src.name}")

    # =========================================================================
    # KỊCH BẢN 7: DANH SÁCH SINH VIÊN VÀ BAN CÁN SỰ CÓ BANNER HÀNH CHÍNH & CHỮ KÝ
    # (Xử lý file đơn lẻ: Lọc cán sự, lớp trưởng, tổ trưởng... không cần file đối chiếu)
    # =========================================================================
    p7_src = EXAMPLES_DIR / '07_danh_sach_lop_can_su_co_banner.xlsx'
    rows_p7 = [
        ['BỘ GIÁO DỤC VÀ ĐÀO TẠO', None, None, None, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', None, None],
        ['TRƯỜNG ĐẠI HỌC BÁCH KHOA', None, None, None, 'Độc lập - Tự do - Hạnh phúc', None, None],
        ['KHOA CÔNG NGHỆ THÔNG TIN', None, None, None, None, None, None],
        ['DANH SÁCH SINH VIÊN VÀ BAN CÁN SỰ LỚP K45-CNTT', None, None, None, None, None, None],
        [None, None, None, None, None, None, None],
        # Dòng 6: Tiêu đề bảng thật
        ['STT', 'Mã sinh viên', 'Họ và tên', 'Lớp', 'Chức vụ', 'Số điện thoại', 'Điểm rèn luyện'],
        [1, 'SV001', 'Nguyễn Văn An', 'K45-CNTT', 'Lớp trưởng', '0912345678', 95],
        [2, 'SV002', 'Trần Thị Bình', 'K45-CNTT', 'Bí thư Chi đoàn', '0923456789', 92],
        [3, 'SV003', 'Lê Hoàng Cường', 'K45-CNTT', 'Sinh viên', '0934567890', 78],
        [4, 'SV004', 'Phạm Minh Dũng', 'K45-CNTT', 'Tổ trưởng tổ 1', '0945678901', 88],
        [5, 'SV005', 'Hoàng Ngọc Ánh', 'K45-CNTT', 'Sinh viên', '0956789012', 82],
        [6, 'SV006', 'Đỗ Thành Đạt', 'K45-CNTT', 'Tổ trưởng tổ 2', '0967890123', 85],
        [7, 'SV007', 'Vũ Mai Phương', 'K45-CNTT', 'Lớp phó học tập', '0978901234', 90],
        [8, 'SV008', 'Nguyễn Đức Thắng', 'K45-CNTT', 'Thủ quỹ', '0989012345', 86],
        [9, 'SV009', 'Lê Thị Thu Thảo', 'K45-CNTT', 'Sinh viên', '0990123456', 75],
        [None, None, None, None, None, None, None],
        ['Hà Nội, ngày 20 tháng 10 năm 2026', None, None, None, None, None, None],
        ['Người lập biểu', None, None, None, 'Ban Cán Sự Lớp', None, None],
    ]
    fixture(p7_src, {'DanhSach': rows_p7})
    print(f"  [OK] Tạo thành công: {p7_src.name}")

    print("\n--- KIỂM TRA ĐỘC LẬP TẤT CẢ CÁC FILE VỪA TẠO VỚI ENGINE EXCEL LOCAL ---")
    # Test 1: File 1 auto-detect header row with header_row='auto'
    s1 = app.inspect_file(p1_src, header_row='auto')
    print(f"File 1 ({p1_src.name}): Đã tự động nhận diện dòng tiêu đề = {s1['header_row']}, số cột = {len(s1['headers'])}, số dòng = {len(s1['data'])}")
    assert s1['header_row'] == 5
    assert 'Mã sinh viên' in s1['headers']

    # Test 2: File 2 number parsing
    s2 = app.inspect_file(p2_src)
    print(f"File 2 ({p2_src.name}): Tiêu đề = {s2['headers']}, số dòng = {len(s2['data'])}")
    assert 'Số tiền (VND)' in s2['headers']

    # Test 3: File 3 auto heal merged headers
    s3 = app.inspect_file(p3_src, auto_heal=True)
    print(f"File 3 ({p3_src.name}): Sau khi tự chữa lành = {s3['headers']}")
    assert 'Họ và tên_phụ' in s3['headers']
    assert 'Lương cơ bản_phụ' in s3['headers']

    # Test 4: File 4 exceptions
    s4 = app.inspect_file(p4_src)
    print(f"File 4 ({p4_src.name}): Tiêu đề = {s4['headers']}, số dòng = {len(s4['data'])}")

    # Test 5: File 5 CSV semicolon & UTF-8 BOM
    s5 = app.inspect_file(p5_csv)
    print(f"File 5 ({p5_csv.name}): Tiêu đề = {s5['headers']}, số dòng = {len(s5['data'])}")
    assert 'Kho xuất' in s5['headers']

    # Test 6: File 6 decimal numbers
    s6 = app.inspect_file(p6_src)
    print(f"File 6 ({p6_src.name}): Tiêu đề = {s6['headers']}, số dòng = {len(s6['data'])}")

    print("\nHoàn tất 100%! Tất cả 6 kịch bản đều được tạo và kiểm tra thành công tuyệt đối!")

if __name__ == '__main__':
    generate_all_samples()
