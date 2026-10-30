# CYCLE-00 — Nghiệm thu nền Windows

Commit kiểm thử: a9a1c5a (nhánh muse-upgrade), checkout cô lập từ bundle trong .work/cycle00.
Môi trường: Windows, Python 3.14; Excel Desktop được cài. COM bị sandbox từ chối (0x80070005), nhưng chạy ngoài sandbox thành công.
Đã kiểm tra STOP.txt trước các bước lớn; file không tồn tại.

## Tóm tắt cập nhật trước bàn giao

- Nhóm 1 sau sửa UC-33: 20/20 lệnh exit 0, 437 check có bộ đếm và 6 kịch bản end-to-end hoàn tất.
- Nhóm 2 sau sửa: 2/2 lệnh, 50/50 check đạt (DPAPI thật).
- Nhóm 3 trước sửa: toàn bộ 4 lệnh đạt; sau sửa đã chạy lại `test_app.py --integration` 97/97, `test_gui.py` 14/14, `test_read_boundary.py` 4/4. Benchmark COM không chạy lại sau sửa theo yêu cầu dừng tại bước bàn giao; số liệu baseline trước sửa nằm ở `results/benchmark-current-1000.json`.
- UC-33 phát hiện app đọc được workbook khi đang mở trong Excel. Đã thêm guard theo owner marker `~$` và regression; probe với Excel thật sau sửa chặn rõ: “Tệp đang được mở trong Excel. Hãy đóng tệp rồi thử lại.”
- W-1 đã kiểm bằng Excel; không có Sheet1 trống. PDF integration đã render/xem, không cắt/chồng.
- Handoff còn lại: quan sát trực tiếp luồng UI, tooltip, phím tắt và hộp thoại lỗi. Không chuyển Chu kỳ 1 trước khi các mục đó được xác nhận. Có thể có false positive nếu owner marker Excel còn sót sau crash; log giới hạn ở NEXT.md.
- Sau test đã khôi phục fixtures/results; không còn tiến trình Excel do test. Checkout đang ở nhánh `muse-upgrade`; không push.
## Output test nguyên văn


===== python -X utf8 test_app.py =====
{
  "passed": 79,
  "integration": false,
  "ai": false,
  "example_output": null,
  "security_scope": "Kiểm tra cục bộ chức năng và trust boundary; chưa nghiệm thu bảo mật toàn máy"
}

EXIT CODE: 0

===== python -X utf8 test_all_edge_cases.py =====
Starting Comprehensive Edge-Case Stress Testing...

Completed! Total Checks: 61
PASSED: 61
FAILED: 0

EXIT CODE: 0

===== python -X utf8 test_upgrades.py =====
All 28 upgrade checks passed successfully!

EXIT CODE: 0

===== python -X utf8 test_cli.py =====
{
  "passed": 6,
  "checks": [
    "default CLI previews without execution",
    "wrong token blocks execution",
    "matching token executes exactly once",
    "changed source blocks previous token",
    "CLI --batch-dir with --plan executes batch successfully",
    "CLI --batch-dir with --ai-prompt infers plan and executes batch"
  ],
  "scope": "CLI gate; execution mocked"
}

EXIT CODE: 0

===== python -X utf8 test_headers.py =====
Traceback (most recent call last):
  File "D:\Excel_update\.work\cycle00\test_headers.py", line 49, in <module>
    try: app.inspect_file(source)
         ~~~~~~~~~~~~~~~~^^^^^^^^
  File "D:\Excel_update\.work\cycle00\app.py", line 484, in inspect_file
    snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
  File "D:\Excel_update\.work\cycle00\app.py", line 279, in bridge
    raise RuntimeError('Excel: ' + response.get('error', 'Thao tác thất bại'))
RuntimeError: Excel: Retrieving the COM class factory for component with CLSID {00024500-0000-0000-C000-000000000046} failed due to the following error: 80070005 Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED)).

EXIT CODE: 1

===== python -X utf8 test_usecases.py =====
--- CSV inspect variants (UC-01..UC-06) ---
  [PASS] UC-01 CSV UTF-16 BOM doc duoc
  [PASS] UC-02 CSV dau cham phay
  [PASS] UC-03 CSV dau tab
  [PASS] UC-04 CSV banner hanh chinh + tu tim dong tieu de
  [PASS] UC-05 CSV 61 cot bi tu choi
  [PASS] UC-06 CSV dung 20000 dong duoc chap nhan
--- Header handling (UC-07..UC-10) ---
  [PASS] UC-07 detect_header_row bo qua banner
  [PASS] UC-08 cot trung ten sau khi cat khoang trang bi tu choi
  [PASS] UC-09 dong tieu de vuot qua du lieu bi tu choi
  [PASS] UC-10 auto_heal xu ly cot trung/khuyet
--- Lookup key semantics via prepare_sheets (UC-11..UC-16) ---
  [PASS] UC-11 khoa thua khoang trang + khac hoa thuong van khop
  [PASS] UC-12 khoa NFD khop khoa NFC
  [PASS] UC-13 so 123 (General) khong khop chuoi 00123
  [PASS] UC-14 so 123 dinh dang 00000 khop chuoi 00123
  [PASS] UC-15 khoa trong duoc dem rieng
  [PASS] UC-16 ky tu * trong khoa duoc hieu theo nghia den
--- Filter (UC-17..UC-19) ---
  [PASS] UC-17 loc theo tu khoa chuc vu
  [PASS] UC-18 loc dieu kien so >=8
  [PASS] UC-19 loc khac rong
--- Summary numbers (UC-20..UC-24) ---
  [PASS] UC-20 tong tien VN dang 1.234.567 o che do lenient
  [PASS] UC-20b 500.000 voi tuy chon vn_dot_thousands=TAT -> thap phan 500
  [PASS] UC-20c 500.000 voi tuy chon vn_dot_thousands=BAT -> 500 nghin
  [PASS] UC-21 so dang 1,234 bi tu choi o che do chat
  [PASS] UC-22 tong vuot 1e14 bi tu choi
  [PASS] UC-23 tong co so am
  [PASS] UC-24 o so trong khong tu doi thanh 0
  [PASS] UC-29 tran 1e14 cua parse_lenient_number (dot bien MUT-5 tung lot)
  [PASS] UC-29b tran 1e14 dang so thang cua parse_lenient_number
--- Dates, clean, token gate (UC-25..UC-28) ---
  [PASS] UC-25 key_value ngay serial khi allow_dates
  [PASS] UC-26 khoa so dinh dang ngay bi tu choi khi chua cho phep ngay
  [PASS] UC-27 clean cat khoang trang dau/cuoi + NFC
  [PASS] UC-28 execute voi token sai bi chan truoc khi cham Excel
  [PASS] SEC-01 validate_plan tu choi action ngoai danh sach (dot bien M3-4 tung lot)
  [PASS] SEC-02 validate_plan tu choi truong la trong ke hoach
--- Preflight OOXML (UC-29..UC-31) ---
  [PASS] UC-29 preflight file sach duoc chap nhan
  [PASS] UC-30 workbook chua macro bi preflight tu choi
  [PASS] UC-31 workbook chua lien ket ngoai bi preflight tu choi

All 37 use-case checks passed.

EXIT CODE: 0

===== python -X utf8 test_fuzz_light.py =====
  [PASS] F-01 CRLF doc duoc
  [PASS] F-02 thieu newline cuoi van doc duoc
  [PASS] F-03 dong trong xen ke: strict tu choi bang ValueError
  [PASS] F-04 dong trong xen ke: lenient bo qua dong trong
  [PASS] F-05 o dung 32767 ky tu duoc chap nhan
  [PASS] F-06 o 32768 ky tu bi tu choi
  [PASS] F-07 ky tu dieu khien o vi tri dau bi tu choi
  [PASS] F-07 ky tu dieu khien o vi tri giua bi tu choi
  [PASS] F-07 ky tu dieu khien o vi tri cuoi bi tu choi
  [PASS] F-08 phan cach hon hop: doc duoc hoac ValueError, khong sap
  [PASS] F-09 file rong bi tu choi bang ValueError
  [PASS] F-10 UTF-8 BOM + cham phay
  [PASS] F-11 CSV cp1258 doc duoc theo hanh vi code hien tai
  [PASS] F-12 nhay kep giua o khong ngoac: doc duoc hoac ValueError
--- header fuzz ---
  [PASS] F-13 header toan khoang trang bi tu choi
  [PASS] F-14 ten cot 1000 ky tu bi tu choi (gioi han 200)
  [PASS] F-14b ten cot dung 200 ky tu duoc chap nhan
  [PASS] F-15 cap header NFC/NFD giong het nhau bi tu choi la trung ten (VD-14)
  [PASS] F-16 detect_header_row voi moi dong giong nhau khong sap
  [PASS] F-17 detect_header_row voi danh sach rong: tra ve int hoac ValueError

All 20 fuzz checks passed (no unexpected exception types).

EXIT CODE: 0

===== python -X utf8 test_fuzz_headers.py =====
  [PASS] H-01 o header None bi tu choi nhu o trong
  [PASS] H-02 dong dau toan so: detect nhay toi dong chu
  [PASS] H-03 moi dong deu la so: khong sap, tra ve int
  [PASS] H-04 chi mot dong duy nhat: khong sap
  [PASS] H-05 dong du lieu lech do dai: doc duoc hoac ValueError ro rang
  [PASS] H-06 NBSP quanh ten cot duoc cat bo
  [PASS] H-07 cap ID/id khac hoa thuong: hanh vi hien tai la chap nhan 2 cot rieng
  [PASS] H-08 ky tu zero-width trong ten cot: doc duoc hoac tu choi, khong sap
  [PASS] H-09 snapshot 61 cot qua select_header: chap nhan (gioi han 60 nam o inspect_file, khong o select_header)
  [PASS] H-10 auto_heal voi cot trung: ten thanh duy nhat
  [PASS] H-11 banner hanh chinh + dong STT/MSSV: detect dung dong 3
  [PASS] H-12 header rong hoan toan bi tu choi
  [PASS] H-13 auto_heal dien o trong o giua bang ten _phu
  [PASS] H-13b auto_heal khong cuu dong chi co 1 ten cot (populated_count <= 1): tu choi nhu che do chat
  [PASS] H-14 dong dau toan None: detect bo qua toi dong chu
  [PASS] H-15 diem thuong du lieu ben duoi quyet dinh dong header that (M5-4)

All 16 header-fuzz checks passed (no unexpected exception types).

EXIT CODE: 0

===== python -X utf8 test_infer_golden.py =====
All 28 infer_smart_plan golden cases passed.

EXIT CODE: 0

===== python -X utf8 test_sanitize_plan.py =====
  [PASS] S-01 action la bi ep ve clarify (M2-3)
  [PASS] S-02 action hop le duoc giu nguyen
  [PASS] S-03 key khong dau duoc fuzzy match ve ten cot that (M2-5)
  [PASS] S-04 value cua summary duoc fuzzy match
  [PASS] S-05 ref_key/ref_value fuzzy match theo cot tham chieu
  [PASS] S-06 filter selected_columns duoc chuan hoa ve cot that
  [PASS] S-07 plan khong phai dict bi tu choi
  [PASS] S-08 cac truong bat buoc duoc dien mac dinh
--- Plan cache key (round 25, M5-1) ---
  [PASS] C-01 cache key khac nhau khi cot nguon khac nhau (M5-1)
  [PASS] C-02 cache key khong phu thuoc thu tu cot
  [PASS] C-03 cache key khac nhau khi cot tham chieu khac nhau
  [PASS] C-04 request_plan khong dung lai plan cache cua bo cot khac

All 12 sanitize_plan checks passed.

EXIT CODE: 0

===== python -X utf8 test_policy_defaults.py =====
  [PASS] P-01 mac dinh khi chua co config: provider=ollama
  [PASS] P-02 mac dinh vn_dot_thousands=True
  [PASS] P-03 config da luu cua nguoi dung duoc ton trong (smart giu nguyen)
  [PASS] P-04 500.000 voi dot_thousands=True -> 500000
  [PASS] P-05 500.000 voi dot_thousands=False -> 500
  [PASS] P-06 12.345 voi True -> 12345
  [PASS] P-07 0.500 luon la thap phan 0.5 ke ca khi True
  [PASS] P-08 1.2345 (4 chu so sau cham) van la thap phan
  [PASS] P-09 -500.000 voi True -> -500000
  [PASS] P-10 da dau cham 1.234.567 khong doi: 1234567
  [PASS] P-11 prepare summary theo config True: tong 500100
  [PASS] P-12 prepare summary theo config False: tong 600

All 12 policy checks passed.

EXIT CODE: 0

===== python -X utf8 test_safety_limits.py =====
  [PASS] L-01 preflight tu choi file 81MB
  [PASS] L-02 preflight chap nhan file dung 80MB
  [PASS] L-03 preflight tu choi XLSX hon 10000 thanh phan
  [PASS] L-04 validate_plan tu choi chuoi dai hon 2000 ky tu
  [PASS] L-05 validate_plan chap nhan chuoi dung 2000 ky tu
  [PASS] L-06 inspect_file phat hien file thay doi giua luc doc (hash)

All 6 safety-limit checks passed.

EXIT CODE: 0

===== python -X utf8 test_fuzz_xlsx.py =====
  [PASS] X-01 file khong phai ZIP -> None, khong nem loi
  [PASS] X-02 ZIP khong co cau truc workbook -> None
  [PASS] X-03 file hop le doc duoc, dung sheet dau
  [PASS] X-04 ngay thang dang date -> YYYY-MM-DD
  [PASS] X-05 datetime co gio -> YYYY-MM-DD HH:MM:SS
  [PASS] X-06 boolean giu nguyen, o trong la None
  [PASS] X-07 sheet_name ton tai -> doc dung sheet do
  [PASS] X-08 sheet_name khong ton tai -> lui ve sheet dau
  [PASS] X-09 sheet rong -> dict voi rows rong (khong None, khong loi)
  [PASS] X-10 inspect lenient + auto header tren XLSX that: nhan dung header dong 3
  [PASS] X-11 read_xml chan DTD/entity
  [PASS] X-12 XML long nhau 200 tang van doc duoc

All 12 xlsx-fuzz checks passed.

EXIT CODE: 0

===== python -X utf8 test_usecases2.py =====
  [PASS] UC2-01 CSV chi co dong header: prepare tu choi vi khong co dong du lieu
  [PASS] UC2-02 tu doi chieu voi chinh no: moi khoa deu khop, missing=0
  [PASS] UC2-03 filter selected_columns chua cot khong ton tai bi validate chan
  [PASS] UC2-04 lookup vua trung khoa tham chieu vua thieu khoa: duplicate=1, missing=1
  [PASS] UC2-05 clean tren cot so: changed=0 va gia tri giu nguyen
  [PASS] UC2-06 summary moi khoa nhom deu trong: bi tu choi
  [PASS] UC2-07 filter >= 8 tren cot toan chu: 0 dong khop, khong loi
  [PASS] UC2-07b filter >= 8 tren cot hon hop: chi dong so thoa man duoc giu
  [PASS] UC2-08 ten sheet rat dai + ky tu la: van xu ly va ghi dung vao sheet kiem tra
  [PASS] UC2-09 lookup thieu ref_value trong ke hoach bi validate chan

All 10 unusual-use-case checks passed.

EXIT CODE: 0

===== python -X utf8 test_batch_summary.py =====
  [PASS] B-01 batch dem dung: 3 file, 2 thanh cong, 1 loi (thieu cot)
  [PASS] B-02 trang thai tung file dung, loi duoc ghi ly do
  [PASS] B-03 file bao cao batch ton tai va khop so dem trong bo nho
  [PASS] B-04 plan action clarify bi tu choi o batch
  [PASS] B-05 thu muc khong co file bang tinh bi tu choi
  [PASS] B-06 lookup thieu file tham chieu bi tu choi

All 6 batch-summary checks passed.

EXIT CODE: 0

===== python -X utf8 test_privacy_guard_cases.py =====
PII cases: 33 | leaked (unexpected): 0
Non-PII cases: 20 | falsely redacted (unexpected): 0
Known limitations hit: 5
  KNOWN [standalone full name (no title)] 'Cho tôi hồ sơ của Nguyễn Văn An' -> 'Cho tôi hồ sơ của Nguyễn Văn An'
  KNOWN [standalone name unaccented lowercase] 'Cho tôi hồ sơ của nguyen van an' -> 'Cho tôi hồ sơ của nguyen van an'
  KNOWN [student id 9 digits] 'Tìm sinh viên MSSV 524002081' -> 'Tìm sinh viên MSSV [CMND_AN_DANH]'
  KNOWN [9-digit amount] 'Tính tổng 123456789 đồng cho quỹ lớp' -> 'Tính tổng [CMND_AN_DANH] đồng cho quỹ lớp'
  KNOWN [13-digit business code] 'Đối chiếu theo mã 0012345678901' -> 'Đối chiếu theo mã [STK_AN_DANH]'
PrivacyGuard corpus: all unexpected-failure counts are zero.

EXIT CODE: 0

===== python -X utf8 test_single_file_filter.py =====
=== Testing Single-File Intelligent Filter & Extract ===
Traceback (most recent call last):
  File "D:\Excel_update\.work\cycle00\test_single_file_filter.py", line 126, in <module>
    run_tests()
    ~~~~~~~~~^^
  File "D:\Excel_update\.work\cycle00\test_single_file_filter.py", line 24, in run_tests
    snapshot = app.inspect_file(sample_file, header_row='auto')
  File "D:\Excel_update\.work\cycle00\app.py", line 484, in inspect_file
    snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
  File "D:\Excel_update\.work\cycle00\app.py", line 279, in bridge
    raise RuntimeError('Excel: ' + response.get('error', 'Thao tác thất bại'))
RuntimeError: Excel: Retrieving the COM class factory for component with CLSID {00024500-0000-0000-C000-000000000046} failed due to the following error: 80070005 Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED)).

EXIT CODE: 1

===== python -X utf8 test_diverse_scenarios_e2e.py =====
=== CHẠY KIỂM THỬ END-TO-END TRÊN TẤT CẢ CÁC FILE MẪU THỰC TẾ ===

1. Kịch bản 1: Học phí có banner + Danh mục sinh viên (Đối chiếu khóa)
Traceback (most recent call last):
  File "D:\Excel_update\.work\cycle00\test_diverse_scenarios_e2e.py", line 109, in <module>
    run_e2e_scenarios()
    ~~~~~~~~~~~~~~~~~^^
  File "D:\Excel_update\.work\cycle00\test_diverse_scenarios_e2e.py", line 27, in run_e2e_scenarios
    src1 = app.inspect_file(EXAMPLES_DIR / '01_hoc_phi_co_banner.xlsx', header_row='auto')
  File "D:\Excel_update\.work\cycle00\app.py", line 484, in inspect_file
    snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
  File "D:\Excel_update\.work\cycle00\app.py", line 279, in bridge
    raise RuntimeError('Excel: ' + response.get('error', 'Thao tác thất bại'))
RuntimeError: Excel: Retrieving the COM class factory for component with CLSID {00024500-0000-0000-C000-000000000046} failed due to the following error: 80070005 Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED)).

EXIT CODE: 1

===== python -X utf8 test_batch_processing.py =====
=== KIỂM THỬ XỬ LÝ HÀNG LOẠT (BATCH PROCESSING) - 0 AI TOKEN ===

1. Kiểm tra Batch Summary (Tổng hợp theo nhóm):
  [PASS] Batch report ghi nhận đúng 3 file
Traceback (most recent call last):
  File "D:\Excel_update\.work\cycle00\test_batch_processing.py", line 118, in <module>
    run_batch_tests()
    ~~~~~~~~~~~~~~~^^
  File "D:\Excel_update\.work\cycle00\test_batch_processing.py", line 68, in run_batch_tests
    check("Tất cả 3 file đều thành công 100%", rep1['succeeded'] == 3 and rep1['failed'] == 0)
    ~~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "D:\Excel_update\.work\cycle00\test_batch_processing.py", line 18, in check
    assert condition, name
           ^^^^^^^^^
AssertionError: Tất cả 3 file đều thành công 100%

EXIT CODE: 1

===== python -X utf8 benchmark_python_stages.py =====
===== rows=100 =====
preflight_xlsx: 0.018s
inspect_source_lenient_fast: 0.012s
inspect_reference_lenient_fast: 0.021s
preview_lookup: 0.002s
preview_summary: 0.002s
preview_clean: 0.001s
inspect_csv: 0.016s
===== rows=1000 =====
preflight_xlsx: 0.028s
inspect_source_lenient_fast: 0.07s
inspect_reference_lenient_fast: 0.057s
preview_lookup: 0.008s
preview_summary: 0.006s
preview_clean: 0.004s
inspect_csv: 0.016s
===== rows=5000 =====
preflight_xlsx: 0.095s
inspect_source_lenient_fast: 0.348s
inspect_reference_lenient_fast: 0.233s
preview_lookup: 0.043s
preview_summary: 0.023s
preview_clean: 0.02s
inspect_csv: 0.029s
===== rows=20000 =====
preflight_xlsx: 0.311s
inspect_source_lenient_fast: 1.41s
inspect_reference_lenient_fast: 0.851s
preview_lookup: 0.167s
preview_summary: 0.098s
preview_clean: 0.08s
inspect_csv: 0.065s

EXIT CODE: 0

===== python -X utf8 test_router.py =====
=== 1. Test DPAPI Encryption for Provider Keys ===
PASS: All keys encrypted on disk! (Windows DPAPI)
PASS: Decryption on the fly works perfectly!

=== 2. Test Smart Router Candidate Generation ===
Active Candidates in priority order: ['Google Gemini (Free Tier)', 'Groq Cloud (Free Llama 70B)', 'OpenAI Official']
PASS: Priority order is Gemini ➔ Groq ➔ OpenAI!

=== 3. Test Q&A / Semantic Instant Answering (0ms, 0 Network Calls) ===
PASS: Instant Q&A response for error explanation!
PASS: Instant response for column count inquiries!

=== 4. Test Auto-Rotation Cascade (Gemini Fails ➔ Groq Takes Over) ===
Plan returned: {'action': 'lookup', 'key': 'MSSV', 'value': '', 'ref_key': 'MSSV', 'ref_value': 'Quê quán', 'message': 'Đã đối chiếu thành công qua Groq Cloud.', 'engine_used': 'Groq Cloud (Free Llama 70B)'}
PASS: Automatic failover from Gemini to Groq succeeded seamlessly!

=== 5. Test Fallback to Ollama Local When All Cloud Quotas Fail ===
Plan returned: {'action': 'summary', 'key': 'Lớp', 'value': 'Học phí', 'ref_key': '', 'ref_value': '', 'message': 'Tổng hợp học phí theo lớp qua Ollama.', 'engine_used': 'Ollama Cục bộ (Dự phòng sau khi 3 Cloud model bận/hết quota)'}
PASS: Automatic failover to local offline Ollama succeeded!

=== 6. Test Semantic Heuristic Rescue When Everything Is Down ===
Plan returned: {'action': 'summary', 'key': 'Lớp', 'value': 'Tiền', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Tiền' theo từng nhóm 'Lớp'.", 'engine_used': 'Smart Heuristic Engine'}
PASS: Semantic Heuristic Rescue rescued the plan with zero network and zero daemons!

=== 7. Test save_config Refuses To Persist Unencrypted Keys (VD-11) ===
PASS: save_config raised a clear error and wrote nothing when DPAPI failed!

All 7 Smart Multi-Model Router architecture tests PASSED brilliantly!

EXIT CODE: 0

===== python -X utf8 test_security_privacy.py =====
All 43 security, privacy, and robustness checks passed successfully!

EXIT CODE: 0

===== python -X utf8 test_app.py --integration =====
Traceback (most recent call last):
  File "D:\Excel_update\.work\cycle00\test_app.py", line 321, in <module>
    if __name__=='__main__': main()
                             ~~~~^^
  File "D:\Excel_update\.work\cycle00\test_app.py", line 300, in main
    source,reference,report=integration_tests()
                            ~~~~~~~~~~~~~~~~~^^
  File "D:\Excel_update\.work\cycle00\test_app.py", line 241, in integration_tests
    source=app.inspect_file(source_path,'Ho_so')
  File "D:\Excel_update\.work\cycle00\app.py", line 484, in inspect_file
    snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
  File "D:\Excel_update\.work\cycle00\app.py", line 279, in bridge
    raise RuntimeError('Excel: ' + response.get('error', 'Thao tác thất bại'))
RuntimeError: Excel: Retrieving the COM class factory for component with CLSID {00024500-0000-0000-C000-000000000046} failed due to the following error: 80070005 Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED)).

EXIT CODE: 1

===== python -X utf8 test_gui.py =====
{
  "passed": 14,
  "checks": [
    "header #0 safely displayed",
    "manual preview waits without executing",
    "cancel produces no execution",
    "explicit acceptance executes exactly once",
    "AI proposal waits for acceptance",
    "AI clarification waits without execution",
    "changed file blocks acceptance",
    "invalid header input reports an error",
    "busy window keeps worker alive",
    "callback error is handled",
    "header error shows raw row numbers and blocks execution",
    "sheet selector remains available after header error",
    "correcting header row recovers same file",
    "reset session clears all inputs and returns to step 1"
  ],
  "scope": "Actual Tk controls; Excel execution mocked"
}

EXIT CODE: 0

===== python -X utf8 test_read_boundary.py =====
Traceback (most recent call last):
  File "D:\Excel_update\.work\cycle00\test_read_boundary.py", line 39, in <module>
    snapshot=app.inspect_file(source)
  File "D:\Excel_update\.work\cycle00\app.py", line 484, in inspect_file
    snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
  File "D:\Excel_update\.work\cycle00\app.py", line 279, in bridge
    raise RuntimeError('Excel: ' + response.get('error', 'Thao tác thất bại'))
RuntimeError: Excel: Retrieving the COM class factory for component with CLSID {00024500-0000-0000-C000-000000000046} failed due to the following error: 80070005 Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED)).

EXIT CODE: 1

===== python -X utf8 benchmark.py =====
preflight 0.033
Traceback (most recent call last):
  File "D:\Excel_update\.work\cycle00\benchmark.py", line 55, in <module>
    main()
    ~~~~^^
  File "D:\Excel_update\.work\cycle00\benchmark.py", line 35, in main
    source = timed('inspect_source', lambda:app.inspect_file(source_path))
  File "D:\Excel_update\.work\cycle00\benchmark.py", line 22, in timed
    value = function()
  File "D:\Excel_update\.work\cycle00\benchmark.py", line 35, in <lambda>
    source = timed('inspect_source', lambda:app.inspect_file(source_path))
                                            ~~~~~~~~~~~~~~~~^^^^^^^^^^^^^
  File "D:\Excel_update\.work\cycle00\app.py", line 484, in inspect_file
    snapshot = bridge({'operation': 'inspect', 'input': str(copy), 'sheet': sheet, 'lenient': bool(lenient)}, directory)
  File "D:\Excel_update\.work\cycle00\app.py", line 279, in bridge
    raise RuntimeError('Excel: ' + response.get('error', 'Thao tác thất bại'))
RuntimeError: Excel: Retrieving the COM class factory for component with CLSID {00024500-0000-0000-C000-000000000046} failed due to the following error: 80070005 Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED)).

EXIT CODE: 1


## Kiểm tra tự động tương ứng các quyết định sản phẩm

- Mặc định Ollama và cấu hình cũ được giữ: test_policy_defaults.py, P-01..P-03 — ĐẠT.
- `500.000`: bật hiểu 500000, tắt hiểu 500; tổng hợp theo config cho kết quả 500100/600 — test_policy_defaults.py P-04, P-05, P-11, P-12 — ĐẠT.
- API key không ghi thô: test_router.py chạy với Windows DPAPI thật, 7/7 đạt; bao gồm lưu/đọc khóa và từ chối ghi nếu DPAPI mô phỏng lỗi — ĐẠT.

## Tình trạng tiến trình Excel

Trước và sau các lần chạy integration/test Excel đã kiểm tra `Get-Process EXCEL`; không phát hiện tiến trình Excel trước lần chạy đầu và không có tiến trình còn lại sau từng lần chạy.

### Số phép kiểm chứng xác nhận

- Nhóm 1: 16/20 lệnh kết thúc thành công; 392 check độc lập được xác nhận (bao gồm 6 check header thuần logic trước khi test_headers thử Excel; các check chưa chạy do COM không tính là đạt). `test_batch_processing.py` có 1 check đầu đạt trước khi check thứ hai thất bại; không gộp check đó vào 392 để tránh trộn một bài chưa hoàn tất.
- Nhóm 2: 50/50 check đạt (7 router + 43 security/privacy).
- Nhóm 3: 14/14 check GUI đạt; 3 bài cần Excel COM (integration, read boundary, benchmark) chưa hoàn tất.

===== Retry outside sandbox: python -X utf8 test_headers.py =====
{
  "passed": 11,
  "checks": [
    "title row reports actual empty cell addresses",
    "empty table margins excluded",
    "original cell format offset preserved",
    "raw source rows unchanged",
    "populated unnamed middle column is not removed",
    "duplicate header identifies both addresses",
    "out-of-range row keeps raw inspection available",
    "native Excel sheet choices retained on header error",
    "native table with leading margin reads correctly",
    "native output keeps code format and source sheets",
    "lookup reads zero-padded code from original offset"
  ],
  "scope": "Header parser + native Excel read/write; synthetic files only",
  "app_sha256": "f713cf3e0011bc9502ab9b7ab94ba171d4bc1a91db1bb74ae293c512b959d3a5"
}

EXIT CODE: 0

===== Retry outside sandbox: python -X utf8 test_single_file_filter.py =====
=== Testing Single-File Intelligent Filter & Extract ===
  [PASS] Auto-detect identifies row 6 (skipping 'Cộng hòa', 'Độc lập' and banner titles)
  [PASS] Headers correctly parsed
  [PASS] Data rows read from source table
  [PASS] Prompt 1 infers action='filter'
  [PASS] Prompt 1 selects key='Chức vụ'
  [PASS] Prompt 1 extracts officer roles in value
  [PASS] User's exact query infers action='filter'
  [PASS] User's exact query maps to key='Chức vụ'
  [PASS] Numeric prompt infers action='filter'
  [PASS] Numeric prompt selects key='Điểm rèn luyện'
  [PASS] Numeric condition parses >= 85
  [PASS] validate_plan passes with reference=None for filter action
  [PASS] Preview calculates 6 filtered officer rows
  [PASS] Preview records total source rows
  [PASS] Preview approval token generated
  [PASS] Sample rows contain only officers
  [PASS] Regular students excluded from preview sample
  [PASS] Target output file exists and is valid XLSX
  [PASS] Report records filter action and 6 rows
  [PASS] Source file preservation verified unchanged
  [PASS] Result sheet contains 6 data rows
  [PASS] Result sheet retains original headers
  [PASS] Numeric filter preview matches scores >= 85 (6 students: 95, 92, 88, 85, 90, 86)
  [PASS] Numeric filter executed successfully via Excel Desktop COM
  [PASS] Column projection plan infers action='filter'
  [PASS] Column projection plan extracts selected_columns
  [PASS] Preview columns only contain the 3 requested columns
  [PASS] Preview sample rows contain 3 columns
  [PASS] Projected filter executed successfully
  [PASS] Result workbook has exactly 3 columns

All 30 single-file filter & extract tests PASSED perfectly!

EXIT CODE: 0

===== Retry outside sandbox: python -X utf8 test_diverse_scenarios_e2e.py =====
=== CHẠY KIỂM THỬ END-TO-END TRÊN TẤT CẢ CÁC FILE MẪU THỰC TẾ ===

1. Kịch bản 1: Học phí có banner + Danh mục sinh viên (Đối chiếu khóa)
   AI Đề xuất: {'action': 'lookup', 'key': 'Mã sinh viên', 'value': '', 'ref_key': 'Mã sinh viên', 'ref_value': 'Khoa chuyên ngành', 'message': "Đối chiếu cột 'Mã sinh viên' với 'Mã sinh viên' để lấy giá trị 'Khoa chuyên ngành' từ danh mục."}
   Xem trước dự kiến: Khớp = 7, Không tìm thấy = 0

2. Kịch bản 2: Bảng kê chi phí kế toán (Tổng hợp theo Phòng ban)
   AI Đề xuất: {'action': 'summary', 'key': 'Phòng ban phát sinh', 'value': 'Số tiền (VND)', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Số tiền (VND)' theo từng nhóm 'Phòng ban phát sinh'."}
   Xem trước dự kiến: Số nhóm = 4, Tổng tiền = 136,000,000.00

3. Kịch bản 3: Bảng lương ô gộp (Tự chữa lành & Làm sạch khoảng trắng)
   AI Đề xuất: {'action': 'clean', 'key': 'Ghi chú hồ sơ', 'value': '', 'ref_key': '', 'ref_value': '', 'message': "Làm sạch khoảng trắng thừa và chuẩn hóa Unicode cột 'Ghi chú hồ sơ'."}
   Xem trước dự kiến: Số ô thay đổi = 5

4. Kịch bản 4: Tuyển sinh (Bắt trúng Khóa trùng, Khóa trống, Ký tự đặc biệt)
   Thống kê rủi ro: Khóa trùng = 2, Khóa trống = 1, Không tìm thấy = 2

5. Kịch bản 5: Bán hàng Kho vận CSV (Phân cách chấm phẩy, UTF-8 BOM)
   AI Đề xuất: {'action': 'summary', 'key': 'Kho xuất', 'value': 'Số lượng', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Số lượng' theo từng nhóm 'Kho xuất'."}
   Xem trước dự kiến: Số kho = 3, Tổng sản phẩm = 142.0

6. Kịch bản 6: Bảng điểm học kỳ (Định dạng số phẩy thập phân Việt Nam)
   AI Đề xuất: {'action': 'summary', 'key': 'Lớp học phần', 'value': 'Điểm trung bình', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Điểm trung bình' theo từng nhóm 'Lớp học phần'."}
   Xem trước dự kiến: Số lớp học phần = 3

=========================================================================
HOÀN TẤT THÀNH CÔNG RỰC RỠ! Tất cả 6 kịch bản đều được phân tích & xử lý tự động 100%!
=========================================================================

EXIT CODE: 0

===== Retry outside sandbox: python -X utf8 test_batch_processing.py =====
=== KIỂM THỬ XỬ LÝ HÀNG LOẠT (BATCH PROCESSING) - 0 AI TOKEN ===

1. Kiểm tra Batch Summary (Tổng hợp theo nhóm):
  [PASS] Batch report ghi nhận đúng 3 file
  [PASS] Tất cả 3 file đều thành công 100%
  [PASS] Tiêu tốn đúng 0 AI token
  [PASS] Callback tiến độ được gọi đủ 3 lần
  [PASS] Tệp báo cáo tổng kết JSON đã được lưu

2. Kiểm tra Batch Lookup (Đối chiếu danh mục):
  [PASS] Batch Lookup thành công 3/3 file

3. Kiểm tra khả năng cách ly lỗi (File hỏng không làm dừng toàn bộ batch):
  [PASS] Nhận diện 4 file trong thư mục
  [PASS] 3 file chuẩn thành công, 1 file lỗi được cô lập an toàn
  [PASS] File lỗi ghi rõ nguyên nhân thiếu cột

4. Kiểm tra lọc file theo pattern:
  [PASS] Lọc pattern lop_ktpm*.xlsx bỏ qua file lỗi

=======================================================
TẤT CẢ 10 KIỂM TRA BATCH ĐỀU PASS 100%!
=======================================================

EXIT CODE: 0

Post-test cleanup note: `test_batch_processing.py` left Excel PID 33008 running. There was no Excel process before the test; PID, executable path (`C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE`) and start time matched a process created during the test. Stopped only PID 33008, then verified no Excel process remained.

===== Retry outside sandbox: python -X utf8 test_read_boundary.py =====
{
  "passed": 3,
  "checks": [
    "Changed copy rejected before Excel opens",
    "Exact copied XML is checked before Excel opens",
    "Verified one-column workbook reads in native Excel"
  ],
  "scope": "Two adversarial swaps + one real Excel read; synthetic files only"
}

EXIT CODE: 0

===== Retry outside sandbox: python -X utf8 benchmark.py =====
preflight 0.03
inspect_source 4.464
inspect_reference 4.772
lookup 7.984
summary 8.212
clean 7.202
D:\Excel_update\.work\cycle00\results\benchmark-current-1000.json

EXIT CODE: 0

Post-test cleanup note: `benchmark.py` left Excel PID 30932 running. There was no Excel process before the benchmark; PID, executable path and start time matched a process created during that test. Stopped only PID 30932 and verified no Excel process remained.

Post-check cleanup note: opening the integration result read-only for W-1 left Excel PID 16784 running. There was no Excel process before that check; the process path and start time matched the check. Stopped only PID 16784 and verified no Excel process remained.

## Tiếp tục sau yêu cầu người dùng

- Thử khởi tạo Computer Use để xác nhận các mục giao diện; node_repl báo `trusted Node process exited unexpectedly` cả trước và sau reset kernel, nên không điều khiển UI.
- Chuẩn bị thử UC-33 bằng workbook tổng hợp, nhưng trước khi tạo file phát hiện một tiến trình Excel đã mở từ bên ngoài lượt test. Không mở/đọc/đóng hay tương tác với workbook đó; UC-33 vẫn chờ phiên test riêng không có Excel người dùng mở.

===== UC-33 synthetic lock scenario =====
Open the generated workbook read/write in a test-owned Excel instance, then run app.inspect_file while it remains open.
READ_WHILE_OPEN=SUCCESS ['ID', 'Value'] 1

PROBE_EXIT_CODE: 0

===== Regression rerun: python -X utf8 test_read_boundary.py =====
{
  "passed": 4,
  "checks": [
    "Excel owner file blocks reading with a clear error",
    "Changed copy rejected before Excel opens",
    "Exact copied XML is checked before Excel opens",
    "Verified one-column workbook reads in native Excel"
  ],
  "scope": "Two adversarial swaps + one real Excel read; synthetic files only"
}

EXIT CODE: 0

===== Post-fix Group 1 rerun =====

===== python -X utf8 test_app.py (post-fix) =====
{
  "passed": 79,
  "integration": false,
  "ai": false,
  "example_output": null,
  "security_scope": "Kiểm tra cục bộ chức năng và trust boundary; chưa nghiệm thu bảo mật toàn máy"
}

EXIT CODE: 0

===== python -X utf8 test_all_edge_cases.py (post-fix) =====
Starting Comprehensive Edge-Case Stress Testing...

Completed! Total Checks: 61
PASSED: 61
FAILED: 0

EXIT CODE: 0

===== python -X utf8 test_upgrades.py (post-fix) =====
All 28 upgrade checks passed successfully!

EXIT CODE: 0

===== python -X utf8 test_cli.py (post-fix) =====
{
  "passed": 6,
  "checks": [
    "default CLI previews without execution",
    "wrong token blocks execution",
    "matching token executes exactly once",
    "changed source blocks previous token",
    "CLI --batch-dir with --plan executes batch successfully",
    "CLI --batch-dir with --ai-prompt infers plan and executes batch"
  ],
  "scope": "CLI gate; execution mocked"
}

EXIT CODE: 0

===== python -X utf8 test_headers.py (post-fix) =====
{
  "passed": 11,
  "checks": [
    "title row reports actual empty cell addresses",
    "empty table margins excluded",
    "original cell format offset preserved",
    "raw source rows unchanged",
    "populated unnamed middle column is not removed",
    "duplicate header identifies both addresses",
    "out-of-range row keeps raw inspection available",
    "native Excel sheet choices retained on header error",
    "native table with leading margin reads correctly",
    "native output keeps code format and source sheets",
    "lookup reads zero-padded code from original offset"
  ],
  "scope": "Header parser + native Excel read/write; synthetic files only",
  "app_sha256": "6b6fee940380c21b88d97f723022cb99f8e498aa65ea673d4e4a3a9140426acd"
}

EXIT CODE: 0

===== python -X utf8 test_usecases.py (post-fix) =====
--- CSV inspect variants (UC-01..UC-06) ---
  [PASS] UC-01 CSV UTF-16 BOM doc duoc
  [PASS] UC-02 CSV dau cham phay
  [PASS] UC-03 CSV dau tab
  [PASS] UC-04 CSV banner hanh chinh + tu tim dong tieu de
  [PASS] UC-05 CSV 61 cot bi tu choi
  [PASS] UC-06 CSV dung 20000 dong duoc chap nhan
--- Header handling (UC-07..UC-10) ---
  [PASS] UC-07 detect_header_row bo qua banner
  [PASS] UC-08 cot trung ten sau khi cat khoang trang bi tu choi
  [PASS] UC-09 dong tieu de vuot qua du lieu bi tu choi
  [PASS] UC-10 auto_heal xu ly cot trung/khuyet
--- Lookup key semantics via prepare_sheets (UC-11..UC-16) ---
  [PASS] UC-11 khoa thua khoang trang + khac hoa thuong van khop
  [PASS] UC-12 khoa NFD khop khoa NFC
  [PASS] UC-13 so 123 (General) khong khop chuoi 00123
  [PASS] UC-14 so 123 dinh dang 00000 khop chuoi 00123
  [PASS] UC-15 khoa trong duoc dem rieng
  [PASS] UC-16 ky tu * trong khoa duoc hieu theo nghia den
--- Filter (UC-17..UC-19) ---
  [PASS] UC-17 loc theo tu khoa chuc vu
  [PASS] UC-18 loc dieu kien so >=8
  [PASS] UC-19 loc khac rong
--- Summary numbers (UC-20..UC-24) ---
  [PASS] UC-20 tong tien VN dang 1.234.567 o che do lenient
  [PASS] UC-20b 500.000 voi tuy chon vn_dot_thousands=TAT -> thap phan 500
  [PASS] UC-20c 500.000 voi tuy chon vn_dot_thousands=BAT -> 500 nghin
  [PASS] UC-21 so dang 1,234 bi tu choi o che do chat
  [PASS] UC-22 tong vuot 1e14 bi tu choi
  [PASS] UC-23 tong co so am
  [PASS] UC-24 o so trong khong tu doi thanh 0
  [PASS] UC-29 tran 1e14 cua parse_lenient_number (dot bien MUT-5 tung lot)
  [PASS] UC-29b tran 1e14 dang so thang cua parse_lenient_number
--- Dates, clean, token gate (UC-25..UC-28) ---
  [PASS] UC-25 key_value ngay serial khi allow_dates
  [PASS] UC-26 khoa so dinh dang ngay bi tu choi khi chua cho phep ngay
  [PASS] UC-27 clean cat khoang trang dau/cuoi + NFC
  [PASS] UC-28 execute voi token sai bi chan truoc khi cham Excel
  [PASS] SEC-01 validate_plan tu choi action ngoai danh sach (dot bien M3-4 tung lot)
  [PASS] SEC-02 validate_plan tu choi truong la trong ke hoach
--- Preflight OOXML (UC-29..UC-31) ---
  [PASS] UC-29 preflight file sach duoc chap nhan
  [PASS] UC-30 workbook chua macro bi preflight tu choi
  [PASS] UC-31 workbook chua lien ket ngoai bi preflight tu choi

All 37 use-case checks passed.

EXIT CODE: 0

===== python -X utf8 test_fuzz_light.py (post-fix) =====
  [PASS] F-01 CRLF doc duoc
  [PASS] F-02 thieu newline cuoi van doc duoc
  [PASS] F-03 dong trong xen ke: strict tu choi bang ValueError
  [PASS] F-04 dong trong xen ke: lenient bo qua dong trong
  [PASS] F-05 o dung 32767 ky tu duoc chap nhan
  [PASS] F-06 o 32768 ky tu bi tu choi
  [PASS] F-07 ky tu dieu khien o vi tri dau bi tu choi
  [PASS] F-07 ky tu dieu khien o vi tri giua bi tu choi
  [PASS] F-07 ky tu dieu khien o vi tri cuoi bi tu choi
  [PASS] F-08 phan cach hon hop: doc duoc hoac ValueError, khong sap
  [PASS] F-09 file rong bi tu choi bang ValueError
  [PASS] F-10 UTF-8 BOM + cham phay
  [PASS] F-11 CSV cp1258 doc duoc theo hanh vi code hien tai
  [PASS] F-12 nhay kep giua o khong ngoac: doc duoc hoac ValueError
--- header fuzz ---
  [PASS] F-13 header toan khoang trang bi tu choi
  [PASS] F-14 ten cot 1000 ky tu bi tu choi (gioi han 200)
  [PASS] F-14b ten cot dung 200 ky tu duoc chap nhan
  [PASS] F-15 cap header NFC/NFD giong het nhau bi tu choi la trung ten (VD-14)
  [PASS] F-16 detect_header_row voi moi dong giong nhau khong sap
  [PASS] F-17 detect_header_row voi danh sach rong: tra ve int hoac ValueError

All 20 fuzz checks passed (no unexpected exception types).

EXIT CODE: 0

===== python -X utf8 test_fuzz_headers.py (post-fix) =====
  [PASS] H-01 o header None bi tu choi nhu o trong
  [PASS] H-02 dong dau toan so: detect nhay toi dong chu
  [PASS] H-03 moi dong deu la so: khong sap, tra ve int
  [PASS] H-04 chi mot dong duy nhat: khong sap
  [PASS] H-05 dong du lieu lech do dai: doc duoc hoac ValueError ro rang
  [PASS] H-06 NBSP quanh ten cot duoc cat bo
  [PASS] H-07 cap ID/id khac hoa thuong: hanh vi hien tai la chap nhan 2 cot rieng
  [PASS] H-08 ky tu zero-width trong ten cot: doc duoc hoac tu choi, khong sap
  [PASS] H-09 snapshot 61 cot qua select_header: chap nhan (gioi han 60 nam o inspect_file, khong o select_header)
  [PASS] H-10 auto_heal voi cot trung: ten thanh duy nhat
  [PASS] H-11 banner hanh chinh + dong STT/MSSV: detect dung dong 3
  [PASS] H-12 header rong hoan toan bi tu choi
  [PASS] H-13 auto_heal dien o trong o giua bang ten _phu
  [PASS] H-13b auto_heal khong cuu dong chi co 1 ten cot (populated_count <= 1): tu choi nhu che do chat
  [PASS] H-14 dong dau toan None: detect bo qua toi dong chu
  [PASS] H-15 diem thuong du lieu ben duoi quyet dinh dong header that (M5-4)

All 16 header-fuzz checks passed (no unexpected exception types).

EXIT CODE: 0

===== python -X utf8 test_infer_golden.py (post-fix) =====
All 28 infer_smart_plan golden cases passed.

EXIT CODE: 0

===== python -X utf8 test_sanitize_plan.py (post-fix) =====
  [PASS] S-01 action la bi ep ve clarify (M2-3)
  [PASS] S-02 action hop le duoc giu nguyen
  [PASS] S-03 key khong dau duoc fuzzy match ve ten cot that (M2-5)
  [PASS] S-04 value cua summary duoc fuzzy match
  [PASS] S-05 ref_key/ref_value fuzzy match theo cot tham chieu
  [PASS] S-06 filter selected_columns duoc chuan hoa ve cot that
  [PASS] S-07 plan khong phai dict bi tu choi
  [PASS] S-08 cac truong bat buoc duoc dien mac dinh
--- Plan cache key (round 25, M5-1) ---
  [PASS] C-01 cache key khac nhau khi cot nguon khac nhau (M5-1)
  [PASS] C-02 cache key khong phu thuoc thu tu cot
  [PASS] C-03 cache key khac nhau khi cot tham chieu khac nhau
  [PASS] C-04 request_plan khong dung lai plan cache cua bo cot khac

All 12 sanitize_plan checks passed.

EXIT CODE: 0

===== python -X utf8 test_policy_defaults.py (post-fix) =====
  [PASS] P-01 mac dinh khi chua co config: provider=ollama
  [PASS] P-02 mac dinh vn_dot_thousands=True
  [PASS] P-03 config da luu cua nguoi dung duoc ton trong (smart giu nguyen)
  [PASS] P-04 500.000 voi dot_thousands=True -> 500000
  [PASS] P-05 500.000 voi dot_thousands=False -> 500
  [PASS] P-06 12.345 voi True -> 12345
  [PASS] P-07 0.500 luon la thap phan 0.5 ke ca khi True
  [PASS] P-08 1.2345 (4 chu so sau cham) van la thap phan
  [PASS] P-09 -500.000 voi True -> -500000
  [PASS] P-10 da dau cham 1.234.567 khong doi: 1234567
  [PASS] P-11 prepare summary theo config True: tong 500100
  [PASS] P-12 prepare summary theo config False: tong 600

All 12 policy checks passed.

EXIT CODE: 0

===== python -X utf8 test_safety_limits.py (post-fix) =====
  [PASS] L-01 preflight tu choi file 81MB
  [PASS] L-02 preflight chap nhan file dung 80MB
  [PASS] L-03 preflight tu choi XLSX hon 10000 thanh phan
  [PASS] L-04 validate_plan tu choi chuoi dai hon 2000 ky tu
  [PASS] L-05 validate_plan chap nhan chuoi dung 2000 ky tu
  [PASS] L-06 inspect_file phat hien file thay doi giua luc doc (hash)

All 6 safety-limit checks passed.

EXIT CODE: 0

===== python -X utf8 test_fuzz_xlsx.py (post-fix) =====
  [PASS] X-01 file khong phai ZIP -> None, khong nem loi
  [PASS] X-02 ZIP khong co cau truc workbook -> None
  [PASS] X-03 file hop le doc duoc, dung sheet dau
  [PASS] X-04 ngay thang dang date -> YYYY-MM-DD
  [PASS] X-05 datetime co gio -> YYYY-MM-DD HH:MM:SS
  [PASS] X-06 boolean giu nguyen, o trong la None
  [PASS] X-07 sheet_name ton tai -> doc dung sheet do
  [PASS] X-08 sheet_name khong ton tai -> lui ve sheet dau
  [PASS] X-09 sheet rong -> dict voi rows rong (khong None, khong loi)
  [PASS] X-10 inspect lenient + auto header tren XLSX that: nhan dung header dong 3
  [PASS] X-11 read_xml chan DTD/entity
  [PASS] X-12 XML long nhau 200 tang van doc duoc

All 12 xlsx-fuzz checks passed.

EXIT CODE: 0

===== python -X utf8 test_usecases2.py (post-fix) =====
  [PASS] UC2-01 CSV chi co dong header: prepare tu choi vi khong co dong du lieu
  [PASS] UC2-02 tu doi chieu voi chinh no: moi khoa deu khop, missing=0
  [PASS] UC2-03 filter selected_columns chua cot khong ton tai bi validate chan
  [PASS] UC2-04 lookup vua trung khoa tham chieu vua thieu khoa: duplicate=1, missing=1
  [PASS] UC2-05 clean tren cot so: changed=0 va gia tri giu nguyen
  [PASS] UC2-06 summary moi khoa nhom deu trong: bi tu choi
  [PASS] UC2-07 filter >= 8 tren cot toan chu: 0 dong khop, khong loi
  [PASS] UC2-07b filter >= 8 tren cot hon hop: chi dong so thoa man duoc giu
  [PASS] UC2-08 ten sheet rat dai + ky tu la: van xu ly va ghi dung vao sheet kiem tra
  [PASS] UC2-09 lookup thieu ref_value trong ke hoach bi validate chan

All 10 unusual-use-case checks passed.

EXIT CODE: 0

===== python -X utf8 test_batch_summary.py (post-fix) =====
  [PASS] B-01 batch dem dung: 3 file, 2 thanh cong, 1 loi (thieu cot)
  [PASS] B-02 trang thai tung file dung, loi duoc ghi ly do
  [PASS] B-03 file bao cao batch ton tai va khop so dem trong bo nho
  [PASS] B-04 plan action clarify bi tu choi o batch
  [PASS] B-05 thu muc khong co file bang tinh bi tu choi
  [PASS] B-06 lookup thieu file tham chieu bi tu choi

All 6 batch-summary checks passed.

EXIT CODE: 0

===== python -X utf8 test_privacy_guard_cases.py (post-fix) =====
PII cases: 33 | leaked (unexpected): 0
Non-PII cases: 20 | falsely redacted (unexpected): 0
Known limitations hit: 5
  KNOWN [standalone full name (no title)] 'Cho tôi hồ sơ của Nguyễn Văn An' -> 'Cho tôi hồ sơ của Nguyễn Văn An'
  KNOWN [standalone name unaccented lowercase] 'Cho tôi hồ sơ của nguyen van an' -> 'Cho tôi hồ sơ của nguyen van an'
  KNOWN [student id 9 digits] 'Tìm sinh viên MSSV 524002081' -> 'Tìm sinh viên MSSV [CMND_AN_DANH]'
  KNOWN [9-digit amount] 'Tính tổng 123456789 đồng cho quỹ lớp' -> 'Tính tổng [CMND_AN_DANH] đồng cho quỹ lớp'
  KNOWN [13-digit business code] 'Đối chiếu theo mã 0012345678901' -> 'Đối chiếu theo mã [STK_AN_DANH]'
PrivacyGuard corpus: all unexpected-failure counts are zero.

EXIT CODE: 0

===== python -X utf8 test_single_file_filter.py (post-fix) =====
=== Testing Single-File Intelligent Filter & Extract ===
  [PASS] Auto-detect identifies row 6 (skipping 'Cộng hòa', 'Độc lập' and banner titles)
  [PASS] Headers correctly parsed
  [PASS] Data rows read from source table
  [PASS] Prompt 1 infers action='filter'
  [PASS] Prompt 1 selects key='Chức vụ'
  [PASS] Prompt 1 extracts officer roles in value
  [PASS] User's exact query infers action='filter'
  [PASS] User's exact query maps to key='Chức vụ'
  [PASS] Numeric prompt infers action='filter'
  [PASS] Numeric prompt selects key='Điểm rèn luyện'
  [PASS] Numeric condition parses >= 85
  [PASS] validate_plan passes with reference=None for filter action
  [PASS] Preview calculates 6 filtered officer rows
  [PASS] Preview records total source rows
  [PASS] Preview approval token generated
  [PASS] Sample rows contain only officers
  [PASS] Regular students excluded from preview sample
  [PASS] Target output file exists and is valid XLSX
  [PASS] Report records filter action and 6 rows
  [PASS] Source file preservation verified unchanged
  [PASS] Result sheet contains 6 data rows
  [PASS] Result sheet retains original headers
  [PASS] Numeric filter preview matches scores >= 85 (6 students: 95, 92, 88, 85, 90, 86)
  [PASS] Numeric filter executed successfully via Excel Desktop COM
  [PASS] Column projection plan infers action='filter'
  [PASS] Column projection plan extracts selected_columns
  [PASS] Preview columns only contain the 3 requested columns
  [PASS] Preview sample rows contain 3 columns
  [PASS] Projected filter executed successfully
  [PASS] Result workbook has exactly 3 columns

All 30 single-file filter & extract tests PASSED perfectly!

EXIT CODE: 0

===== python -X utf8 test_diverse_scenarios_e2e.py (post-fix) =====
=== CHẠY KIỂM THỬ END-TO-END TRÊN TẤT CẢ CÁC FILE MẪU THỰC TẾ ===

1. Kịch bản 1: Học phí có banner + Danh mục sinh viên (Đối chiếu khóa)
   AI Đề xuất: {'action': 'lookup', 'key': 'Mã sinh viên', 'value': '', 'ref_key': 'Mã sinh viên', 'ref_value': 'Khoa chuyên ngành', 'message': "Đối chiếu cột 'Mã sinh viên' với 'Mã sinh viên' để lấy giá trị 'Khoa chuyên ngành' từ danh mục."}
   Xem trước dự kiến: Khớp = 7, Không tìm thấy = 0

2. Kịch bản 2: Bảng kê chi phí kế toán (Tổng hợp theo Phòng ban)
   AI Đề xuất: {'action': 'summary', 'key': 'Phòng ban phát sinh', 'value': 'Số tiền (VND)', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Số tiền (VND)' theo từng nhóm 'Phòng ban phát sinh'."}
   Xem trước dự kiến: Số nhóm = 4, Tổng tiền = 136,000,000.00

3. Kịch bản 3: Bảng lương ô gộp (Tự chữa lành & Làm sạch khoảng trắng)
   AI Đề xuất: {'action': 'clean', 'key': 'Ghi chú hồ sơ', 'value': '', 'ref_key': '', 'ref_value': '', 'message': "Làm sạch khoảng trắng thừa và chuẩn hóa Unicode cột 'Ghi chú hồ sơ'."}
   Xem trước dự kiến: Số ô thay đổi = 5

4. Kịch bản 4: Tuyển sinh (Bắt trúng Khóa trùng, Khóa trống, Ký tự đặc biệt)
   Thống kê rủi ro: Khóa trùng = 2, Khóa trống = 1, Không tìm thấy = 2

5. Kịch bản 5: Bán hàng Kho vận CSV (Phân cách chấm phẩy, UTF-8 BOM)
   AI Đề xuất: {'action': 'summary', 'key': 'Kho xuất', 'value': 'Số lượng', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Số lượng' theo từng nhóm 'Kho xuất'."}
   Xem trước dự kiến: Số kho = 3, Tổng sản phẩm = 142.0

6. Kịch bản 6: Bảng điểm học kỳ (Định dạng số phẩy thập phân Việt Nam)
   AI Đề xuất: {'action': 'summary', 'key': 'Lớp học phần', 'value': 'Điểm trung bình', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Điểm trung bình' theo từng nhóm 'Lớp học phần'."}
   Xem trước dự kiến: Số lớp học phần = 3

=========================================================================
HOÀN TẤT THÀNH CÔNG RỰC RỠ! Tất cả 6 kịch bản đều được phân tích & xử lý tự động 100%!
=========================================================================

EXIT CODE: 0

===== python -X utf8 test_batch_processing.py (post-fix) =====
=== KIỂM THỬ XỬ LÝ HÀNG LOẠT (BATCH PROCESSING) - 0 AI TOKEN ===

1. Kiểm tra Batch Summary (Tổng hợp theo nhóm):
  [PASS] Batch report ghi nhận đúng 3 file
  [PASS] Tất cả 3 file đều thành công 100%
  [PASS] Tiêu tốn đúng 0 AI token
  [PASS] Callback tiến độ được gọi đủ 3 lần
  [PASS] Tệp báo cáo tổng kết JSON đã được lưu

2. Kiểm tra Batch Lookup (Đối chiếu danh mục):
  [PASS] Batch Lookup thành công 3/3 file

3. Kiểm tra khả năng cách ly lỗi (File hỏng không làm dừng toàn bộ batch):
  [PASS] Nhận diện 4 file trong thư mục
  [PASS] 3 file chuẩn thành công, 1 file lỗi được cô lập an toàn
  [PASS] File lỗi ghi rõ nguyên nhân thiếu cột

4. Kiểm tra lọc file theo pattern:
  [PASS] Lọc pattern lop_ktpm*.xlsx bỏ qua file lỗi

=======================================================
TẤT CẢ 10 KIỂM TRA BATCH ĐỀU PASS 100%!
=======================================================

EXIT CODE: 0

===== python -X utf8 benchmark_python_stages.py (post-fix) =====
===== rows=100 =====
preflight_xlsx: 0.021s
inspect_source_lenient_fast: 0.019s
inspect_reference_lenient_fast: 0.028s
preview_lookup: 0.003s
preview_summary: 0.002s
preview_clean: 0.002s
inspect_csv: 0.023s
===== rows=1000 =====
preflight_xlsx: 0.057s
inspect_source_lenient_fast: 0.192s
inspect_reference_lenient_fast: 0.124s
preview_lookup: 0.018s
preview_summary: 0.011s
preview_clean: 0.01s
inspect_csv: 0.027s
===== rows=5000 =====
preflight_xlsx: 0.228s
inspect_source_lenient_fast: 1.655s
inspect_reference_lenient_fast: 1.044s
preview_lookup: 0.141s
preview_summary: 0.152s
preview_clean: 0.074s
inspect_csv: 0.165s
===== rows=20000 =====
preflight_xlsx: 0.841s
inspect_source_lenient_fast: 6.666s
inspect_reference_lenient_fast: 3.147s
preview_lookup: 0.46s
preview_summary: 0.312s
preview_clean: 0.24s
inspect_csv: 0.162s

EXIT CODE: 0
[
  {
    "Test": "test_app.py",
    "ExitCode": 0
  },
  {
    "Test": "test_all_edge_cases.py",
    "ExitCode": 0
  },
  {
    "Test": "test_upgrades.py",
    "ExitCode": 0
  },
  {
    "Test": "test_cli.py",
    "ExitCode": 0
  },
  {
    "Test": "test_headers.py",
    "ExitCode": 0
  },
  {
    "Test": "test_usecases.py",
    "ExitCode": 0
  },
  {
    "Test": "test_fuzz_light.py",
    "ExitCode": 0
  },
  {
    "Test": "test_fuzz_headers.py",
    "ExitCode": 0
  },
  {
    "Test": "test_infer_golden.py",
    "ExitCode": 0
  },
  {
    "Test": "test_sanitize_plan.py",
    "ExitCode": 0
  },
  {
    "Test": "test_policy_defaults.py",
    "ExitCode": 0
  },
  {
    "Test": "test_safety_limits.py",
    "ExitCode": 0
  },
  {
    "Test": "test_fuzz_xlsx.py",
    "ExitCode": 0
  },
  {
    "Test": "test_usecases2.py",
    "ExitCode": 0
  },
  {
    "Test": "test_batch_summary.py",
    "ExitCode": 0
  },
  {
    "Test": "test_privacy_guard_cases.py",
    "ExitCode": 0
  },
  {
    "Test": "test_single_file_filter.py",
    "ExitCode": 0
  },
  {
    "Test": "test_diverse_scenarios_e2e.py",
    "ExitCode": 0
  },
  {
    "Test": "test_batch_processing.py",
    "ExitCode": 0
  },
  {
    "Test": "benchmark_python_stages.py",
    "ExitCode": 0
  }
]

===== Post-fix Group 2 rerun =====

===== python -X utf8 test_router.py (post-fix) =====
=== 1. Test DPAPI Encryption for Provider Keys ===
PASS: All keys encrypted on disk! (Windows DPAPI)
PASS: Decryption on the fly works perfectly!

=== 2. Test Smart Router Candidate Generation ===
Active Candidates in priority order: ['Google Gemini (Free Tier)', 'Groq Cloud (Free Llama 70B)', 'OpenAI Official']
PASS: Priority order is Gemini ➔ Groq ➔ OpenAI!

=== 3. Test Q&A / Semantic Instant Answering (0ms, 0 Network Calls) ===
PASS: Instant Q&A response for error explanation!
PASS: Instant response for column count inquiries!

=== 4. Test Auto-Rotation Cascade (Gemini Fails ➔ Groq Takes Over) ===
Plan returned: {'action': 'lookup', 'key': 'MSSV', 'value': '', 'ref_key': 'MSSV', 'ref_value': 'Quê quán', 'message': 'Đã đối chiếu thành công qua Groq Cloud.', 'engine_used': 'Groq Cloud (Free Llama 70B)'}
PASS: Automatic failover from Gemini to Groq succeeded seamlessly!

=== 5. Test Fallback to Ollama Local When All Cloud Quotas Fail ===
Plan returned: {'action': 'summary', 'key': 'Lớp', 'value': 'Học phí', 'ref_key': '', 'ref_value': '', 'message': 'Tổng hợp học phí theo lớp qua Ollama.', 'engine_used': 'Ollama Cục bộ (Dự phòng sau khi 3 Cloud model bận/hết quota)'}
PASS: Automatic failover to local offline Ollama succeeded!

=== 6. Test Semantic Heuristic Rescue When Everything Is Down ===
Plan returned: {'action': 'summary', 'key': 'Lớp', 'value': 'Tiền', 'ref_key': '', 'ref_value': '', 'lenient_numbers': True, 'message': "Tính tổng cột 'Tiền' theo từng nhóm 'Lớp'.", 'engine_used': 'Smart Heuristic Engine'}
PASS: Semantic Heuristic Rescue rescued the plan with zero network and zero daemons!

=== 7. Test save_config Refuses To Persist Unencrypted Keys (VD-11) ===
PASS: save_config raised a clear error and wrote nothing when DPAPI failed!

All 7 Smart Multi-Model Router architecture tests PASSED brilliantly!

EXIT CODE: 0

===== python -X utf8 test_security_privacy.py (post-fix) =====
All 43 security, privacy, and robustness checks passed successfully!

EXIT CODE: 0
[
  {
    "Test": "test_router.py",
    "ExitCode": 0
  },
  {
    "Test": "test_security_privacy.py",
    "ExitCode": 0
  }
]

===== Post-fix Group 3 rerun =====

===== python -X utf8 test_app.py --integration (post-fix) =====
Unknown option: --integration
usage: C:\Python314\python.exe [option] ... [-c cmd | -m mod | file | -] [arg] ...
Try `python -h' for more information.

EXIT CODE: 2

===== python -X utf8 test_gui.py (post-fix) =====
{
  "passed": 14,
  "checks": [
    "header #0 safely displayed",
    "manual preview waits without executing",
    "cancel produces no execution",
    "explicit acceptance executes exactly once",
    "AI proposal waits for acceptance",
    "AI clarification waits without execution",
    "changed file blocks acceptance",
    "invalid header input reports an error",
    "busy window keeps worker alive",
    "callback error is handled",
    "header error shows raw row numbers and blocks execution",
    "sheet selector remains available after header error",
    "correcting header row recovers same file",
    "reset session clears all inputs and returns to step 1"
  ],
  "scope": "Actual Tk controls; Excel execution mocked"
}

EXIT CODE: 0

===== python -X utf8 test_read_boundary.py (post-fix) =====
{
  "passed": 4,
  "checks": [
    "Excel owner file blocks reading with a clear error",
    "Changed copy rejected before Excel opens",
    "Exact copied XML is checked before Excel opens",
    "Verified one-column workbook reads in native Excel"
  ],
  "scope": "Two adversarial swaps + one real Excel read; synthetic files only"
}

EXIT CODE: 0

===== python -X utf8 benchmark.py (post-fix) =====
usage: benchmark.py [-h] [--rows ROWS] [--label LABEL]
benchmark.py: error: unrecognized arguments: benchmark.py

EXIT CODE: 2
[
  {
    "Test": "test_app.py --integration",
    "ExitCode": 2
  },
  {
    "Test": "test_gui.py",
    "ExitCode": 0
  },
  {
    "Test": "test_read_boundary.py",
    "ExitCode": 0
  },
  {
    "Test": "benchmark.py",
    "ExitCode": 2
  }
]

===== Corrected post-fix command: python -X utf8 test_app.py --integration =====
{
  "passed": 97,
  "integration": true,
  "ai": false,
  "example_output": "D:\\Excel_update\\.work\\cycle00\\output\\20261030-223641_lookup_afb3266b.xlsx",
  "security_scope": "Kiểm tra cục bộ chức năng và trust boundary; chưa nghiệm thu bảo mật toàn máy"
}

EXIT CODE: 0

===== Corrected post-fix command: python -X utf8 benchmark.py =====

===== UC-33 actual Excel lock rerun after fix =====
READ_WHILE_OPEN=BLOCKED Tệp đang được mở trong Excel. Hãy đóng tệp rồi thử lại.

PROBE_EXIT_CODE: 0

Test cleanup: stopped hidden Excel PID 26180 from UC-33 probe.
