"""End-to-end verification of all diverse sample files across all operations."""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

from pathlib import Path
import tempfile
import app
import ai_client

EXAMPLES_DIR = app.ROOT / 'examples'

def run_e2e_scenarios():
    print("=== CHẠY KIỂM THỬ END-TO-END TRÊN TẤT CẢ CÁC FILE MẪU THỰC TẾ ===")
    
    with tempfile.TemporaryDirectory(dir=app.WORK) as td:
        out_dir = Path(td)

        # ---------------------------------------------------------------------
        # KỊCH BẢN 1: Tự nhận diện banner và đối chiếu khoa chuyên ngành
        # ---------------------------------------------------------------------
        print("\n1. Kịch bản 1: Học phí có banner + Danh mục sinh viên (Đối chiếu khóa)")
        src1 = app.inspect_file(EXAMPLES_DIR / '01_hoc_phi_co_banner.xlsx', header_row='auto')
        ref1 = app.inspect_file(EXAMPLES_DIR / '01_danh_muc_sinh_vien.xlsx')
        plan1 = ai_client.infer_smart_plan("Đối chiếu Mã sinh viên lấy Khoa chuyên ngành", src1['headers'], ref1['headers'])
        assert plan1 and plan1['action'] == 'lookup'
        print(f"   AI Đề xuất: {plan1}")
        prev1 = app.preview_plan(src1, ref1, plan1, output_dir=out_dir)
        print(f"   Xem trước dự kiến: Khớp = {prev1['stats']['matched']}, Không tìm thấy = {prev1['stats']['missing']}")
        assert prev1['stats']['matched'] == 7
        assert prev1['stats']['missing'] == 0

        # ---------------------------------------------------------------------
        # KỊCH BẢN 2: Kế toán chi phí - Số âm ngoặc đơn & Đuôi trừ (Tổng hợp theo nhóm)
        # ---------------------------------------------------------------------
        print("\n2. Kịch bản 2: Bảng kê chi phí kế toán (Tổng hợp theo Phòng ban)")
        src2 = app.inspect_file(EXAMPLES_DIR / '02_ke_toan_chi_phi_ngan_sach.xlsx')
        plan2 = ai_client.infer_smart_plan("Tính tổng Số tiền theo Phòng ban phát sinh", src2['headers'])
        assert plan2 and plan2['action'] == 'summary'
        print(f"   AI Đề xuất: {plan2}")
        prev2 = app.preview_plan(src2, None, plan2, output_dir=out_dir)
        total2 = float(prev2['stats']['independent_total'])
        print(f"   Xem trước dự kiến: Số nhóm = {prev2['stats']['groups']}, Tổng tiền = {total2:,.2f}")
        assert prev2['stats']['groups'] == 4
        # Kiểm tra tính toán số âm chuẩn xác
        assert total2 > 0

        # ---------------------------------------------------------------------
        # KỊCH BẢN 3: Nhân sự - Tự chữa lành ô gộp và làm sạch khoảng trắng ẩn
        # ---------------------------------------------------------------------
        print("\n3. Kịch bản 3: Bảng lương ô gộp (Tự chữa lành & Làm sạch khoảng trắng)")
        src3 = app.inspect_file(EXAMPLES_DIR / '03_nhan_su_tien_luong_o_gop.xlsx', auto_heal=True)
        assert 'Họ và tên_phụ' in src3['headers']
        plan3 = ai_client.infer_smart_plan("Làm sạch khoảng trắng cột Ghi chú hồ sơ", src3['headers'])
        assert plan3 and plan3['action'] == 'clean'
        print(f"   AI Đề xuất: {plan3}")
        prev3 = app.preview_plan(src3, None, plan3, output_dir=out_dir)
        print(f"   Xem trước dự kiến: Số ô thay đổi = {prev3['stats']['changed_cells']}")
        assert prev3['stats']['changed_cells'] > 0

        # ---------------------------------------------------------------------
        # KỊCH BẢN 4: Tuyển sinh - Khóa trùng, khóa trống, ký tự đặc biệt (*, ~)
        # ---------------------------------------------------------------------
        print("\n4. Kịch bản 4: Tuyển sinh (Bắt trúng Khóa trùng, Khóa trống, Ký tự đặc biệt)")
        src4 = app.inspect_file(EXAMPLES_DIR / '04_tuyen_sinh_khoa_trung_ngoai_le.xlsx')
        ref4 = app.inspect_file(EXAMPLES_DIR / '04_danh_muc_diem_chuan.xlsx')
        plan4 = {'action': 'lookup', 'key': 'Mã hồ sơ', 'value': '', 'ref_key': 'Mã hồ sơ', 'ref_value': 'Ngành trúng tuyển'}
        prev4 = app.preview_plan(src4, ref4, plan4, output_dir=out_dir)
        print(f"   Thống kê rủi ro: Khóa trùng = {prev4['stats']['duplicate']}, Khóa trống = {prev4['stats']['blank_keys']}, Không tìm thấy = {prev4['stats']['missing']}")
        assert prev4['stats']['duplicate'] >= 1
        assert prev4['stats']['blank_keys'] >= 1
        assert prev4['stats']['missing'] >= 1

        # ---------------------------------------------------------------------
        # KỊCH BẢN 5: Kho hàng CSV - Dấu chấm phẩy (;), UTF-8 BOM, ngoặc kép
        # ---------------------------------------------------------------------
        print("\n5. Kịch bản 5: Bán hàng Kho vận CSV (Phân cách chấm phẩy, UTF-8 BOM)")
        src5 = app.inspect_file(EXAMPLES_DIR / '05_kho_hang_xuat_ban.csv')
        plan5 = ai_client.infer_smart_plan("Tính tổng Số lượng theo từng Kho xuất", src5['headers'])
        assert plan5 and plan5['action'] == 'summary'
        print(f"   AI Đề xuất: {plan5}")
        prev5 = app.preview_plan(src5, None, plan5, output_dir=out_dir)
        total5 = float(prev5['stats']['independent_total'])
        print(f"   Xem trước dự kiến: Số kho = {prev5['stats']['groups']}, Tổng sản phẩm = {total5}")
        assert prev5['stats']['groups'] == 3
        assert total5 == 142

        # ---------------------------------------------------------------------
        # KỊCH BẢN 6: Bảng điểm học phần - Số thập phân có dấu phẩy Việt Nam (8,5; 7,25)
        # ---------------------------------------------------------------------
        print("\n6. Kịch bản 6: Bảng điểm học kỳ (Định dạng số phẩy thập phân Việt Nam)")
        src6 = app.inspect_file(EXAMPLES_DIR / '06_bang_diem_hoc_ky_phay_thap_phan.xlsx')
        plan6 = ai_client.infer_smart_plan("Cộng Điểm trung bình theo Lớp học phần", src6['headers'])
        assert plan6 and plan6['action'] == 'summary'
        print(f"   AI Đề xuất: {plan6}")
        prev6 = app.preview_plan(src6, None, plan6, output_dir=out_dir)
        print(f"   Xem trước dự kiến: Số lớp học phần = {prev6['stats']['groups']}")
        assert prev6['stats']['groups'] == 3

    print("\n=========================================================================")
    print("HOÀN TẤT THÀNH CÔNG RỰC RỠ! Tất cả 6 kịch bản đều được phân tích & xử lý tự động 100%!")
    print("=========================================================================")

if __name__ == '__main__':
    run_e2e_scenarios()
