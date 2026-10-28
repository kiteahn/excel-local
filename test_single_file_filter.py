"""Test single-file filter & extract: auto-header detection, natural language intent, preview, and Excel Desktop COM execution."""
import json
from pathlib import Path
import tempfile
import app
import ai_client

checks = []

def check(name, condition):
    assert condition, name
    checks.append(name)
    print(f"  [PASS] {name}")

def run_tests():
    print("=== Testing Single-File Intelligent Filter & Extract ===")
    sample_file = app.ROOT / 'examples' / '07_danh_sach_lop_can_su_co_banner.xlsx'
    assert sample_file.exists(), f"File {sample_file} must exist."

    # 1. Test Auto-Header Detection with Vietnamese Administrative Banners
    snapshot = app.inspect_file(sample_file, header_row='auto')
    check("Auto-detect identifies row 6 (skipping 'Cộng hòa', 'Độc lập' and banner titles)",
          snapshot['header_row'] == 6)
    check("Headers correctly parsed",
          snapshot['headers'] == ['STT', 'Mã sinh viên', 'Họ và tên', 'Lớp', 'Chức vụ', 'Số điện thoại', 'Điểm rèn luyện'])
    check("Data rows read from source table",
          len(snapshot['data']) >= 9)

    # 2. Test AI / Semantic Heuristic for Natural Language Filter Requests
    headers = snapshot['headers']
    prompt1 = "Lọc danh sách các cán sự, lớp trưởng, tổ trưởng, bí thư"
    plan1 = ai_client.infer_smart_plan(prompt1, headers)
    check("Prompt 1 infers action='filter'", plan1['action'] == 'filter')
    check("Prompt 1 selects key='Chức vụ'", plan1['key'] == 'Chức vụ')
    check("Prompt 1 extracts officer roles in value",
          any(r in plan1['value'] for r in ['lớp trưởng', 'tổ trưởng', 'bí thư']))

    # 3. Test exact user query with banners mentioned
    user_prompt = "ví dụ tôi yêu cầu đọc file đó, có danh sách các cán sự, lớp trưởng, tổ trưởng, v.v, Trong file excel đầu vào, có thể được viết rằng độc lập cộng hòa..., nhưng nội dung chính bên dưới, tôi cần nó đọc thế xong lấy ra cái tôi yêu cầu, k phải lúc nào cũng cần file đối chiếu thì mới làm được, mà chỉ cần một file đơn lẻ và lấy đúng yêu cầu ra mà làm"
    user_plan = ai_client.infer_smart_plan(user_prompt, headers)
    check("User's exact query infers action='filter'", user_plan['action'] == 'filter')
    check("User's exact query maps to key='Chức vụ'", user_plan['key'] == 'Chức vụ')

    # 4. Test Numeric Filter Inference
    prompt_num = "Lọc danh sách sinh viên có điểm rèn luyện lớn hơn hoặc bằng 85"
    plan_num = ai_client.infer_smart_plan(prompt_num, headers)
    check("Numeric prompt infers action='filter'", plan_num['action'] == 'filter')
    check("Numeric prompt selects key='Điểm rèn luyện'", plan_num['key'] == 'Điểm rèn luyện')
    check("Numeric condition parses >= 85", '>= 85' in plan_num['value'] or '>= 85.0' in plan_num['value'])

    # 5. Test Plan Validation without Reference Table
    validated1 = app.validate_plan(plan1, snapshot, reference=None)
    check("validate_plan passes with reference=None for filter action",
          validated1['action'] == 'filter' and validated1['key'] == 'Chức vụ')

    # 6. Test Preview Plan in RAM
    with tempfile.TemporaryDirectory(dir=app.WORK) as tmp_dir:
        out_dir = Path(tmp_dir) / 'output'
        preview = app.preview_plan(snapshot, None, validated1, output_dir=out_dir)
        check("Preview calculates 6 filtered officer rows",
              preview['stats']['filtered_rows'] == 6)
        check("Preview records total source rows",
              preview['stats']['source_rows'] >= 9)
        check("Preview approval token generated",
              bool(preview['approval_token']) and len(preview['approval_token']) == 64)

        # Verify sample rows in preview
        sample_officers = [r[headers.index('Chức vụ')] for r in preview['sample_rows']]
        check("Sample rows contain only officers",
              all(any(role in o.lower() for role in ['lớp trưởng', 'bí thư', 'tổ trưởng', 'lớp phó', 'thủ quỹ']) for o in sample_officers))
        check("Regular students excluded from preview sample",
              'Sinh viên' not in sample_officers)

        # 7. Test Real Excel Desktop COM Execution & Audit Verification
        report = app.execute(snapshot, None, validated1, approval_token=preview['approval_token'], output_dir=out_dir)
        target_path = Path(report['output'])
        check("Target output file exists and is valid XLSX",
              target_path.exists() and target_path.suffix.lower() == '.xlsx')
        check("Report records filter action and 6 rows",
              report['action'] == 'filter' and report['filtered_rows'] == 6)
        check("Source file preservation verified unchanged",
              report['preservation']['unchanged'] is True)

        # Inspect generated workbook
        res_snapshot = app.inspect_file(target_path, report['result_sheet'])
        check("Result sheet contains 6 data rows",
              len(res_snapshot['data']) == 6)
        check("Result sheet retains original headers",
              res_snapshot['headers'] == headers)

        # 8. Test Numeric Filter Execution E2E
        validated_num = app.validate_plan(plan_num, snapshot, reference=None)
        preview_num = app.preview_plan(snapshot, None, validated_num, output_dir=out_dir)
        check("Numeric filter preview matches scores >= 85 (6 students: 95, 92, 88, 85, 90, 86)",
              preview_num['stats']['filtered_rows'] == 6)
        report_num = app.execute(snapshot, None, validated_num, approval_token=preview_num['approval_token'], output_dir=out_dir)
        check("Numeric filter executed successfully via Excel Desktop COM",
              report_num['action'] == 'filter' and report_num['filtered_rows'] == 6)

        # 9. Test Column Selection / Projection
        prompt_proj = "Lọc cán sự, chỉ lấy Họ và tên, Chức vụ, Số điện thoại"
        plan_proj = ai_client.infer_smart_plan(prompt_proj, headers)
        check("Column projection plan infers action='filter'", plan_proj['action'] == 'filter')
        check("Column projection plan extracts selected_columns",
              'Họ và tên' in plan_proj['selected_columns'] and 'Chức vụ' in plan_proj['selected_columns'] and 'Số điện thoại' in plan_proj['selected_columns'])
        validated_proj = app.validate_plan(plan_proj, snapshot, reference=None)
        preview_proj = app.preview_plan(snapshot, None, validated_proj, output_dir=out_dir)
        check("Preview columns only contain the 3 requested columns",
              set(preview_proj['columns']) == {'Họ và tên', 'Chức vụ', 'Số điện thoại'} and len(preview_proj['columns']) == 3)
        check("Preview sample rows contain 3 columns",
              len(preview_proj['sample_rows'][0]) == 3)
        report_proj = app.execute(snapshot, None, validated_proj, approval_token=preview_proj['approval_token'], output_dir=out_dir)
        check("Projected filter executed successfully",
              report_proj['action'] == 'filter' and report_proj['filtered_rows'] == 6)
        res_proj_snapshot = app.inspect_file(Path(report_proj['output']), report_proj['result_sheet'])
        check("Result workbook has exactly 3 columns",
              set(res_proj_snapshot['headers']) == {'Họ và tên', 'Chức vụ', 'Số điện thoại'} and len(res_proj_snapshot['headers']) == 3)

    print(f"\nAll {len(checks)} single-file filter & extract tests PASSED perfectly!")
    return len(checks)

if __name__ == '__main__':
    run_tests()

