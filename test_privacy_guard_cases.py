"""Measured corpus for PrivacyGuard (muc E): >=50 labelled prompt samples.

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
]

# Cases the current regex design is known to get wrong; measured and printed,
# but they do not fail the run. Anything NOT listed here must pass.
KNOWN_LIMITATIONS = {
    "student id 9 digits",       # indistinguishable from CMND by regex alone
    "order code 11 digits",
    "9-digit amount",
    "13-digit business code",    # indistinguishable from a bank account by regex alone
    "single-word surname after title-like word",
    "standalone full name (no title)",   # regex cannot tell a bare name from ordinary capitalised words
    "standalone name unaccented lowercase",
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
    "student id 9 digits": "524002081",
    "order code 11 digits": "20261030001",
    "9-digit amount": "123456789 đồng",
    "13-digit business code": "0012345678901",
    "single-word surname after title-like word": "Nhân viên Nguyễn",
}

def _matches(tag, prompt):
    marker = _MARKERS.get(tag, '')
    return bool(marker) and marker in prompt

if __name__ == '__main__':
    main()
