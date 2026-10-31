# Bản 0.3.2 · Privacy Guard chính xác hơn: bịt lọt tên, bỏ che nhầm số tiền/mã

Ngày: 2026-10-31 · Nhánh `muse-upgrade` (cục bộ, không push)

## Vì sao có bản này

Bản 0.3.1 che PII bằng regex nhưng đo được hai loại lỗi trái ngược:

- **Lọt PII:** tên riêng đứng một mình, không có từ xưng hô
  (`Cho tôi hồ sơ của Nguyễn Văn An`), không bị che.
- **Che nhầm:** số tiền 9 chữ số (`123456789 đồng`) bị thay bằng `[CMND_AN_DANH]`,
  mã nghiệp vụ 13 chữ số (`mã 0012345678901`) bị thay bằng `[STK_AN_DANH]`,
  MSSV 9 chữ số (`MSSV 524002081`) bị thay bằng `[CMND_AN_DANH]`.

Che nhầm làm mô hình mất ngữ cảnh và dễ đề xuất sai cột; lọt tên là rủi ro riêng tư
khi người dùng bật cloud. Cả hai đều là lỗi hành vi người dùng thấy, nên bản này
tăng số phiên bản thay vì giữ 0.3.1.

## Đã sửa

1. **Ngữ cảnh số 9 chữ số:** giữ nguyên khi theo sau là đơn vị tiền
   (`đồng`, `vnđ`, `vnd`, `nghìn`, `ngàn`, `triệu`, `tỷ`, `tỉ`, `đ`) hoặc khi đứng sau
   nhãn `MSSV` / `MSV` / `mã SV` / `mã sinh viên`.
2. **Ngữ cảnh mã 13 chữ số:** giữ nguyên khi đứng sau `mã` (tối đa 3 từ đệm như
   "mã hồ sơ", "mã đơn hàng"), **trừ** khi cụm đệm nói về `tài khoản`, `ngân hàng`,
   `STK`, `TK` — các trường hợp đó vẫn che thành `[STK_AN_DANH]`.
3. **Tên riêng đứng một mình:** nhận diện 3–4 từ Title Case bắt đầu bằng họ Việt
   phổ biến (danh sách trong `ai_client.py`), và dạng viết thường/không dấu khi có
   `văn`/`van`/`thị`/`thi` ở giữa. Loại trừ địa danh/danh nhân lịch sử đã biết
   ("Hồ Chí Minh", "Trần Hưng Đạo", "Lý Thường Kiệt", "Phạm Ngũ Lão",
   "Đinh Tiên Hoàng", "Võ Văn Tần", "Hai Bà Trưng") và bỏ qua khi ngay trước là từ
   chỉ địa danh/tổ chức (`đường`, `phố`, `phường`, `tỉnh`, `thành phố`, `trường`,
   `công ty`, `ngân hàng`, ...).
4. **Lỗi ăn khoảng trắng (phát hiện khi đo):** SĐT/STK/CCCD có dấu cách hoặc gạch
   nối từng nuốt luôn dấu cách phía sau, làm câu gửi AI bị dính chữ
   (`[STK_AN_DANH]của tôi`). Đã sửa để khớp kết thúc bằng chữ số.

Không đổi: luật SĐT/CCCD/email/MST/mật khẩu/khóa, schema plan, giao diện, luồng
xem trước + token, định dạng file xuất.

## Số đo trước / sau

Bộ đo: `python -X utf8 test_privacy_guard_cases.py` (dữ liệu giả, không gọi mạng).

| Chỉ số | 0.3.1 | 0.3.2 |
|--------|-------|-------|
| Mẫu PII | 33 | 80 |
| Mẫu không phải PII | 20 | 37 |
| Rò lọt ngoài danh sách known | 0 | 0 |
| Che nhầm ngoài danh sách known | 0 | 0 |
| Known hit (đã ghi, không tính fail) | 5 | 2 |

2 known hit còn lại là giới hạn có chủ ý: tên chỉ **2 từ** (`Lê Minh`) và tên dân
tộc **không bắt đầu bằng họ Kinh** (`Y Blô Niê`). Không thể phân biệt với văn bản
thường chỉ bằng regex mà không tăng mạnh che nhầm.

## Kiểm tra đã chạy (Windows, 2026-10-31)

| Lệnh | Kết quả |
|------|---------|
| `python -X utf8 test_privacy_guard_cases.py` | 80 PII / 0 lọt, 37 non-PII / 0 che nhầm, 2 known |
| `python -X utf8 test_security_privacy.py` | 43/43 |
| `python -X utf8 test_router.py` | 7/7 |
| `python -X utf8 test_app.py --integration` | 97/97 (Excel Desktop COM thật) |
| `python -X utf8 test_usecases.py` | 37/37 |
| `python -X utf8 test_all_edge_cases.py` | 61/61 |
| `python -X utf8 test_gui.py` | 14/14 |

Mutation thủ công 5 chỗ cho luật mới (tắt lần lượt luật tiền, MSSV, mã 13 chữ số,
nhận diện tên, loại trừ địa danh): cả 5 đều bị bộ đo bắt (4/4/3 lỗi che nhầm và
14 lỗi lọt tên), rồi phục hồi mã nguồn.

## Chưa kiểm chứng

- Chưa gọi API cloud thật để xem phản hồi của mô hình với câu đã che; bộ đo chỉ kiểm
  tra đầu vào sau khi che (hàm thuần, không mạng).
- Chưa đo trên câu tiếng Việt do người dùng thật nhập ngoài bộ mẫu; các ca hiếm như
  tên 2 từ, tên dân tộc, tên đường trùng họ vẫn có thể sai.
- Chưa chạy lại benchmark hiệu năng; thay đổi chỉ thêm vài phép regex trên prompt
  ngắn, chưa đo tác động.

## Cách kiểm tra nhanh

```powershell
python -X utf8 test_privacy_guard_cases.py
# Tính tổng 123456789 đồng cho quỹ lớp  → giữ nguyên số tiền
# Đối chiếu theo mã 0012345678901       → giữ nguyên mã
# Cho tôi hồ sơ của Nguyễn Văn An       → [TEN_AN_DANH]
```
