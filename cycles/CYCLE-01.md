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

## Pha 2 — Thực hiện (2026-10-31)

Đã sửa `ai_client.py` (chỉ thêm luật + tách hàm, không đổi chữ ký công khai):

- `_protect_context_spans()`: thay tạm các vùng số **được tin** bằng token không
  chứa chữ số trước khi chạy các regex che; sau khi che thì hoàn nguyên. Nhờ vậy
  không phải nới bất kỳ regex SĐT/CCCD/STK/MST/email/khóa nào.
  - `MONEY_9_REGEX`: 9 chữ số + đơn vị tiền (`đồng`, `vnđ`, `vnd`, `nghìn`, `ngàn`,
    `triệu`, `tỷ`, `tỉ`, `đ`).
  - `STUDENT_ID_9_REGEX`: 9 chữ số sau `MSSV`/`MSV`/`mã SV`/`mã sinh viên`.
  - `BUSINESS_CODE_13_REGEX`: 13 chữ số sau `mã` + tối đa 3 từ đệm, **trừ** khi cụm
    đệm chứa `tài khoản`/`ngân hàng`/`STK`/`TK`.
- `_standalone_name_spans()` + `_looks_like_name()`: tên 3–4 từ Title Case bắt đầu
  bằng họ Việt phổ biến; dạng thường/không dấu chỉ nhận khi có `văn/van/thị/thi`
  ở giữa; loại trừ 7 địa danh/danh nhân lịch sử và bỏ qua khi ngay trước là từ chỉ
  địa danh/tổ chức.
- Sửa lỗi phát hiện khi đo: `PHONE_REGEX`/`BANK_REGEX` ăn mất dấu cách sau chuỗi
  có phân cách (`[STK_AN_DANH]của tôi`). Đổi phần đuôi để kết thúc bằng chữ số.

## Pha 3 — QA / đo lại

`python -X utf8 test_privacy_guard_cases.py`:

```
PII cases: 80 | leaked (unexpected): 0
Non-PII cases: 37 | falsely redacted (unexpected): 0
Known limitations hit: 2
  KNOWN [two-word standalone name] 'Hồ sơ của Lê Minh cần xem'
  KNOWN [ethnic name without common surname] 'Hồ sơ của Y Blô Niê cần xem'
```

So với Pha 1: 33→80 mẫu PII, 20→37 mẫu non-PII, known 5→2, không phát sinh
rò lọt/che nhầm ngoài danh sách. Hai known còn lại đúng là giới hạn của thiết kế
regex (tên 2 từ; tên dân tộc không bắt đầu bằng họ Kinh).

## Pha 4 — Mutation 5 chỗ

Chạy script tạm (không đưa vào repo) tắt lần lượt từng luật rồi chạy lại bộ đo:

| Mutation | Kết quả |
|----------|---------|
| Tắt luật tiền 9 chữ số | bắt: 4 che nhầm |
| Tắt luật MSSV 9 chữ số | bắt: 4 che nhầm |
| Tắt luật mã 13 chữ số | bắt: 3 che nhầm |
| Tắt nhận diện tên đứng một mình | bắt: 14 lọt tên |
| Tắt loại trừ địa danh | bắt: 3 che nhầm |

Cả 5 đều bị bắt; mã nguồn được phục hồi sau mỗi lần.

## Pha 5 — Kiểm tra hồi quy (Windows)

| Lệnh | Kết quả |
|------|---------|
| `test_security_privacy.py` | 43/43 |
| `test_router.py` | 7/7 |
| `test_app.py --integration` | 97/97 |
| `test_usecases.py` | 37/37 |
| `test_all_edge_cases.py` | 61/61 (gồm prompt an toàn không bị đổi) |
| `test_gui.py` | 14/14 |
| `test_privacy_guard_cases.py` | 80/80 PII, 37/37 non-PII, 2 known |

## Pha 6 — Phát hành

- Tăng 0.3.2 (hành vi người dùng thấy), cập nhật tiêu đề cửa sổ, README, RELEASES.
- Thêm `release-notes-0.3.2.md` với số đo trước/sau và mục "Chưa kiểm chứng".
- Tag cục bộ `v0.3.2` (không push).

## Pha 7 — Chấm lại & chu kỳ kế

- S-1..S-4 trong `ROADMAP.md` **ĐÓNG** (đo được, có test bảo vệ).
- P0 còn lại: chưa có mục mới đủ nghiêm trọng; chuyển trọng tâm sang P1 (dễ bảo trì).
- Chu kỳ 2: tách `open_ai_settings()` khỏi `gui()` (M-2), giữ nguyên giao diện;
  `test_gui` 14 + `test_ui_acceptance` 48 là lưới an toàn.

## Giới hạn còn lại (ghi trung thực)

- Tên 2 từ (`Lê Minh`) và tên dân tộc không thuộc họ Kinh (`Y Blô Niê`) vẫn lọt.
- Tên đường trùng họ người không nằm trong danh sách loại trừ có thể bị che nhầm;
  đã giảm bằng tiền tố địa danh nhưng chưa đo hết.
- Bộ đo là dữ liệu giả, không phải câu người dùng thật; chưa gọi cloud thật.
- Chưa đo lại benchmark hiệu năng sau thay đổi (prompt ngắn, dự kiến không đáng kể
  nhưng chưa có số).

