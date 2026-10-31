# CYCLE-01 — PrivacyGuard chính xác hơn (giảm che nhầm, bịt lọt tên)

## Pha 1 — Kế hoạch

### Vấn đề người dùng thật
- Gõ "Tính tổng 123456789 đồng cho quỹ lớp" hoặc "Đối chiếu theo mã 0012345678901"
  thì câu gửi AI bị thay bằng `[CMND_AN_DANH]` / `[STK_AN_DANH]`; mô hình mất ngữ
  cảnh số tiền/mã, dễ đề xuất sai cột.
- Gõ tên riêng không kèm từ xưng hô ("Cho tôi hồ sơ của Nguyễn Văn An") thì tên
  KHÔNG bị che — rò rỉ PII ra cloud nếu người dùng bật cloud.
- Đo hiện tại (`test_privacy_guard_cases.py`): 33 mẫu PII, 0 rò rỉ ngoài danh sách
  known; 20 mẫu không phải PII, 0 che nhầm ngoài known; **5 known hit**: 2 lọt tên,
  3 che nhầm (số tiền 9 chữ số, mã 13 chữ số, MSSV 9 chữ số).

### Mục tiêu
1. Bịt 2 ca lọt tên riêng đứng một mình (có dấu và không dấu) khi câu có >= 3 từ
   Title Case bắt đầu bằng họ Việt phổ biến; trừ địa danh/tên cơ quan đã biết.
2. Bỏ che nhầm 3 ca: số tiền 9 chữ số theo sau là đơn vị tiền; mã 13 chữ số đứng
   sau "mã"; MSSV 9 chữ số đứng sau "MSSV"/"mã sinh viên"/"mã SV".
3. Mở rộng bộ đo lên >= 80 mẫu PII và >= 30 mẫu không phải PII, gồm ca biên mới.

### Không làm trong chu kỳ này
- Không đổi luật SĐT/CCCD/STK/MST/email/khóa hiện có (trừ khi test chứng minh lỗi).
- Không đổi giao diện, không đổi luồng thao tác, không đổi schema plan.
- Không dùng thư viện NLP ngoài; chỉ regex + từ điển họ tiếng Việt trong file.

### Tiêu chí nghiệm thu (đo được)
- `python -X utf8 test_privacy_guard_cases.py`:
  - rò rỉ ngoài known = **0/N** với N >= 60 mẫu PII;
  - che nhầm ngoài known = **0/M** với M >= 30 mẫu không phải PII;
  - known hit còn lại **<= 2** (chỉ chấp nhận ca không thể phân biệt bằng ngữ cảnh).
- `test_security_privacy.py` 43/43 và `test_router.py` 7/7 vẫn đạt.
- `test_app.py` 79/79, `test_usecases.py` 37/37 không đổi.
- README mục PrivacyGuard cập nhật đúng số liệu mới và giới hạn còn lại.

### Rủi ro
- Từ điển họ có thể che nhầm tên đường/phố ("Hà Nội", "Lê Lợi"). Giảm thiểu bằng
  danh sách địa danh loại trừ + chỉ nhận >= 3 từ và đo trên mẫu không phải PII.
- Nới ngữ cảnh "mã" có thể bỏ lọt STK thật đứng sau chữ "mã". Giảm thiểu: chỉ bỏ
  qua khi số có đúng 13 chữ số và không có tiền tố ngân hàng (STK/tài khoản).

### Lệnh kiểm tra
```powershell
python -X utf8 test_privacy_guard_cases.py
python -X utf8 test_security_privacy.py
python -X utf8 test_router.py
python -X utf8 test_app.py
python -X utf8 test_usecases.py
```

## Pha 2..7
(chưa làm — ghi kết quả thật vào đây sau mỗi bước)
