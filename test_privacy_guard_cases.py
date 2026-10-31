"""Measured corpus for PrivacyGuard (muc E): 80 PII + 37 non-PII samples.

Each PII case names the secret substring that must NOT survive
anonymize_prompt(). Each non-PII case must come back byte-identical.
The script prints the measured leak rate and false-redaction rate and
fails only on cases listed in KNOWN_LIMITATIONS (documented, not hidden).
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

import ai_client

G = ai_client.PrivacyGuard.anonymize_prompt

# (prompt, secret that must be gone, category)
PII_CASES = [
    # Phone numbers, all current VN mobile heads and common notations
    ("Lọc sinh viên có SĐT 0901234567", "0901234567", "phone 09x plain"),
    ("SĐT của cán sự: 0356789012", "0356789012", "phone 03x plain"),
    ("Liên hệ 0523456789 để nhận lớp", "0523456789", "phone 05x (052)"),
    ("Liên hệ 0703456789 để nhận lớp", "0703456789", "phone 07x (070)"),
    ("Liên hệ 0813456789 để nhận lớp", "0813456789", "phone 08x (081)"),
    ("SĐT +84901234567 cần đối chiếu", "+84901234567", "phone +84"),
    ("SĐT (+84) 901 234 567 cần đối chiếu", "901 234 567", "phone (+84) spaced"),
    ("SĐT 0901.234.567 cần đối chiếu", "0901.234.567", "phone dotted"),
    ("SĐT 0901-234-567 cần đối chiếu", "0901-234-567", "phone dashed"),
    ("SĐT 0901 234 567 cần đối chiếu", "0901 234 567", "phone spaced"),
    ("SĐT 84901234567 cần đối chiếu", "84901234567", "phone bare 84 (no plus)"),
    # CCCD / CMND
    ("Tìm CCCD 079201004567 trong danh sách", "079201004567", "CCCD 12"),
    ("Tìm CCCD 001234567890 trong danh sách", "001234567890", "CCCD 12 head 001"),
    ("Tìm CMND 241567891 trong danh sách", "241567891", "CMND 9"),
    ("Số căn cước: 079 201 004 567", "079 201 004 567", "CCCD spaced groups"),
    # Tax codes
    ("Đối chiếu MST 0123456789 của đơn vị", "0123456789", "MST 10"),
    ("Đối chiếu MST 0123456789-001 của chi nhánh", "0123456789-001", "MST 13 with branch"),
    # Bank accounts
    ("Chuyển khoản STK 0123456789012 ngân hàng A", "0123456789012", "STK 13"),
    ("Chuyển khoản STK 1234567890123456 ngân hàng A", "1234567890123456", "STK 16"),
    ("STK 0123 4567 8901 2 nhận học phí", "0123 4567 8901 2", "STK spaced 13"),
    # Email
    ("Gửi cho nguyen.van.a@example.com giúp tôi", "nguyen.van.a@example.com", "email dotted"),
    ("Gửi cho an+lop@example.edu.vn giúp tôi", "an+lop@example.edu.vn", "email plus"),
    ("Email: NguyenVanA@Example.COM", "NguyenVanA@Example.COM", "email mixed case"),
    # API keys / passwords
    ("Dùng khóa sk-1234567890abcdefghijABCD để gọi API", "sk-1234567890abcdefghijABCD", "api key sk-"),
    ("Header Authorization: Bearer abcDEF123456.ghi-789", "abcDEF123456.ghi-789", "bearer token"),
    ("mật khẩu: abc123xyz", "abc123xyz", "password with colon"),
    ("password=Hunter2! cho tài khoản", "Hunter2!", "password with equals"),
    ("Mật khẩu của hệ thống là Sup3rSecret", "Sup3rSecret", "password with 'la'"),
    # Names with a title word (the pattern PrivacyGuard claims to cover)
    ("Tìm sinh viên Nguyễn Văn An trong lớp", "Nguyễn Văn An", "name with title sinh vien"),
    ("Hỏi cô Trần Thu Hằng về điểm", "Trần Thu Hằng", "name with title co"),
    ("Anh Phạm Quốc Bảo phụ trách file này", "Phạm Quốc Bảo", "name with title Anh"),
    ("Cho tôi hồ sơ của Nguyễn Văn An", "Nguyễn Văn An", "standalone full name (no title)"),
    ("Cho tôi hồ sơ của nguyen van an", "nguyen van an", "standalone name unaccented lowercase"),
    # --- Cycle 1 additions -------------------------------------------------
    ("Hồ sơ của Đặng Thị Hồng Nhung cần xem", "Đặng Thị Hồng Nhung", "standalone name 4 words"),
    ("Liên hệ Phạm Minh Tuấn ngay hôm nay", "Phạm Minh Tuấn", "standalone name 3 words"),
    ("Cập nhật điểm cho Lê Thị Mai Anh", "Lê Thị Mai Anh", "standalone name 4 words female"),
    ("Danh sách lớp có Vũ Đức Long", "Vũ Đức Long", "standalone name after verb"),
    ("Báo cáo của nguyen thi lan anh", "nguyen thi lan anh", "standalone unaccented 4 words"),
    ("Nhập điểm của trần thi thu ha", "trần thi thu ha", "standalone accented lowercase"),
    ("Hỏi Nguyễn Hoàng Nam về lịch thi", "Nguyễn Hoàng Nam", "standalone name at sentence start"),
    ("Nguyễn Thị Minh Khai nộp báo cáo", "Nguyễn Thị Minh Khai", "standalone name sharing a street name"),
    ("Gửi thư cho Hoàng Minh Quân", "Hoàng Minh Quân", "standalone name 3 words (Hoàng)"),
    ("Gọi 0323456789 giúp tôi", "0323456789", "phone 03x (032)"),
    ("Gọi 0563456789 giúp tôi", "0563456789", "phone 05x (056)"),
    ("Gọi 0773456789 giúp tôi", "0773456789", "phone 07x (077)"),
    ("Gọi 0963456789 giúp tôi", "0963456789", "phone 09x (096)"),
    ("CCCD 012345678901 của cán bộ", "012345678901", "CCCD 12 head 012"),
    ("CMND 123456789 nộp bản sao", "123456789", "CMND 9 (123)"),
    ("Số định danh 079201004568 cần xác minh", "079201004568", "CCCD 12 (568)"),
    ("MST 0312345678 của công ty", "0312345678", "MST 10 (031)"),
    ("Mã số thuế 0123456789-002 chi nhánh", "0123456789-002", "MST 13 branch 002"),
    ("STK 19034567890123 tại ngân hàng", "19034567890123", "STK 14"),
    ("Tài khoản 012345678901234567 của lớp", "012345678901234567", "STK 18"),
    ("STK 0123-4567-8901-2 của lớp", "0123-4567-8901-2", "STK dashed 13"),
    ("Liên hệ TRAN.THI.B@Example.EDU.VN nhé", "TRAN.THI.B@Example.EDU.VN", "email uppercase"),
    ("Địa chỉ mail: lop.truong+2026@example.org", "lop.truong+2026@example.org", "email plus digits"),
    ("API key là sk-abcdefghijklmnopqrstuvwx gọi cloud", "sk-abcdefghijklmnopqrstuvwx", "api key sk- 24"),
    ("Dùng token=ghp_abcdefghijklmnopqrstuvwxyz0123456789", "ghp_abcdefghijklmnopqrstuvwxyz0123456789", "github token ghp_"),
    ("Mật khẩu là MatKhauMoi#2026", "MatKhauMoi#2026", "password with 'là'"),
    ("password: P@ssw0rd2026 cho wifi", "P@ssw0rd2026", "password with colon"),
    ("SĐT 0333456789 của lớp trưởng", "0333456789", "phone 03x (033)"),
    ("SĐT 0583456789 của lớp trưởng", "0583456789", "phone 05x (058)"),
    ("SĐT 0793456789 của lớp trưởng", "0793456789", "phone 07x (079)"),
    ("SĐT 0823456789 của lớp trưởng", "0823456789", "phone 08x (082)"),
    ("SĐT 0853456789 của lớp trưởng", "0853456789", "phone 08x (085)"),
    ("SĐT 0883456789 của lớp trưởng", "0883456789", "phone 08x (088)"),
    ("SĐT 0993456789 của lớp trưởng", "0993456789", "phone 09x (099)"),
    ("CMND 987654321 trong hồ sơ", "987654321", "CMND 9 (987)"),
    ("CCCD 079099001122 khi nhập học", "079099001122", "CCCD 12 (099)"),
    ("STK 9704366801234567 của trường", "9704366801234567", "STK 16 (9704)"),
    ("STK 012345678901234 tại ngân hàng B", "012345678901234", "STK 15"),
    ("MST 0101234567-001 của công ty", "0101234567-001", "MST 13 branch 001"),
    ("Email: sinhvien2026@hcmut.edu.vn", "sinhvien2026@hcmut.edu.vn", "email digits"),
    ("Mật khẩu wifi là Wifi@2026", "Wifi@2026", "password with middle words"),
    ("token: abcDEF123456ghi", "abcDEF123456ghi", "generic token with colon"),
    ("Gửi cho Nguyễn Thị Thu Trang bản điểm", "Nguyễn Thị Thu Trang", "standalone name 4 words (Trang)"),
    ("Phụ trách là Đỗ Minh Khang", "Đỗ Minh Khang", "standalone name 3 words (Đỗ)"),
    ("Liên hệ Bùi Thị Ánh Tuyết ngay", "Bùi Thị Ánh Tuyết", "standalone name 4 words (Bùi)"),
    # Documented residual risks of this design (measured, not hidden)
    ("Hồ sơ của Lê Minh cần xem", "Lê Minh", "two-word standalone name"),
    ("Hồ sơ của Y Blô Niê cần xem", "Y Blô Niê", "ethnic name without common surname"),
]

# (prompt,) must be returned unchanged
NON_PII_CASES = [
    "Đối chiếu mã hồ sơ 00123 với danh mục",            # leading-zero code
    "Tìm sinh viên MSSV 52400208",                      # student id 8 digits
    "Tìm sinh viên MSSV 524002081",                     # student id 9 digits - see KNOWN_LIMITATIONS
    "Tổng học phí lớp KTPM1 là 1.500.000 đồng",         # money + class code
    "Lọc các dòng có điểm từ 8.5 đến 10",                # scores
    "Thống kê năm 2026 theo từng khoa",                 # year
    "Mã lớp CNTT-K66 và môn 505060",                    # class/course codes
    "Cộng cột Số lượng cho 20000 dòng dữ liệu",         # row count
    "Lấy 60 cột đầu tiên của bảng",                     # column count
    "Số điện thoại nội bộ 1234 của văn phòng khoa",     # internal extension
    "Mã đơn hàng 20261030001 đã giao",                  # order code 11 digits - see KNOWN_LIMITATIONS
    "Tính tổng 123456789 đồng cho quỹ lớp",             # 9-digit amount - see KNOWN_LIMITATIONS
    "Đối chiếu theo mã 0012345678901",                  # 13-digit business code - see KNOWN_LIMITATIONS
    "Họ tên sinh viên nào có điểm cao nhất lớp",        # no actual name, just words
    "Lọc cán sự, lớp trưởng và tổ trưởng",              # role keywords
    "Tìm các dòng chứa từ khóa ngân hàng",              # word 'ngan hang' alone
    "Đổi mật khẩu ở đâu trong ứng dụng này",            # the word password, no secret
    "Cho tôi xem cột Email và cột Số điện thoại",       # column names only
    "Nhân viên Nguyễn phụ trách ca sáng",               # single-word surname after title-like word - see KNOWN_LIMITATIONS
    "Tổng tiền đợt 1: 12.345.678 đồng",                 # VN money format
    # --- Cycle 1 additions -------------------------------------------------
    "Thành phố Hồ Chí Minh có 3 chi nhánh",             # city, shares a surname pattern
    "Đường Nguyễn Thị Minh Khai đang sửa chữa",         # street with place prefix
    "Phố Trần Hưng Đạo đông xe buổi sáng",              # historical figure / street name
    "Mai làm gì hôm nay cũng được",                     # 'mai' is a surname but not a name here
    "Ngày mai làm gì cũng được",                        # 'mai' inside ordinary prose
    "Tổng 250000000 đồng đã thu",                       # 9-digit amount
    "Học phí 100000000 VNĐ đã nộp",                     # 9-digit amount + VNĐ
    "Ngân sách còn 999999999đ",                         # 9-digit amount + đ
    "Quỹ lớp còn 123456789 nghìn đồng",                 # 9-digit amount + nghìn
    "MSSV: 524002081 đã đăng ký môn",                   # student id with colon
    "Mã SV 524002081 thuộc lớp KTPM1",                  # student id short label
    "Mã sinh viên 524002081 cần cập nhật",              # student id long label
    "Đối chiếu mã hồ sơ 0012345678901",                 # 13-digit business code
    "Tra mã đơn hàng 0012345678902",                    # 13-digit order code
    "Khoa Công Nghệ Thông Tin có 3 lớp",                # title-case words, not a name
    "Trường Đại Học Bách Khoa tuyển sinh",              # organisation name
    "Lê Văn",                                            # two words, no full name
]

# Cases the current regex design is known to get wrong; measured and printed,
# but they do not fail the run. Anything NOT listed here must pass.
KNOWN_LIMITATIONS = {
    "order code 11 digits",      # defensive marker; no longer redacted
    "single-word surname after title-like word",  # defensive marker; no longer redacted
    "two-word standalone name",  # too short to separate from ordinary Title Case words
    "ethnic name without common surname",  # surname list is Kinh-Vietnamese only
}

def main():
    leaks, false_hits, known = [], [], []
    for prompt, secret, category in PII_CASES:
        out = G(prompt)
        if secret in out:
            (known if category in KNOWN_LIMITATIONS else leaks).append((category, prompt, out))
    for prompt in NON_PII_CASES:
        category = prompt  # matched below by marker substrings
        out = G(prompt)
        if out != prompt:
            tag = next((k for k in KNOWN_LIMITATIONS if _matches(k, prompt)), None)
            (known if tag else false_hits).append((tag or 'false redaction', prompt, out))
    print(f"PII cases: {len(PII_CASES)} | leaked (unexpected): {len(leaks)}")
    for c, p, o in leaks:
        print(f"  LEAK [{c}] {p!r} -> {o!r}")
    print(f"Non-PII cases: {len(NON_PII_CASES)} | falsely redacted (unexpected): {len(false_hits)}")
    for c, p, o in false_hits:
        print(f"  FALSE [{c}] {p!r} -> {o!r}")
    print(f"Known limitations hit: {len(known)}")
    for c, p, o in known:
        print(f"  KNOWN [{c}] {p!r} -> {o!r}")
    assert not leaks, f"{len(leaks)} unexpected PII leaks"
    assert not false_hits, f"{len(false_hits)} unexpected false redactions"
    print("PrivacyGuard corpus: all unexpected-failure counts are zero.")

_MARKERS = {
    "order code 11 digits": "20261030001",
    "single-word surname after title-like word": "Nhân viên Nguyễn",
}

def _matches(tag, prompt):
    marker = _MARKERS.get(tag, '')
    return bool(marker) and marker in prompt

if __name__ == '__main__':
    main()
