# Chu kỳ 6 — Hiệu năng đo thật trên Excel COM (Q3)

## Mục tiêu
- Chạy lại `benchmark.py` và `benchmark_python_stages.py` sau thay đổi PrivacyGuard, so mốc trong `results/`.
- Phân rã: khởi động Excel, mở file, đọc, ghi, tính, đối soát.
- Chỉ tối ưu giai đoạn lớn nhất; mỗi thay đổi ≥5 lần đo, lấy trung vị; không giữ nếu <10% mà phức tạp hơn.
- Đo thêm 20.000 dòng và ghi giới hạn.

## Cổng nghiệm thu
- `run_all_tests.ps1` đạt; `test_prepare_sheets_golden.py` vẫn 27/27.
- Số liệu trước/sau nằm trong `results/`, có trung vị và số lần đo.
- Không đổi công thức/nhãn/đối soát số học.

## Cách đo
- `benchmark_python_stages.py`: 5 lần, lưu `results/benchmark-stages-q3-before-*.json`.
- `benchmark.py --rows 5000`: 5 lần; `--rows 20000`: 5 lần (nhãn `q3-before*`).
- Bổ sung `inspect_phases` (Excel start / mở file / find / đọc / convert) và `write_detail` (build / paste / format / style / wrap-autofit) vào kết quả; chỉ đo, không đổi hành vi.

## Kết quả (2026-11-02, máy Windows thật, Excel Microsoft 365)
### Python thuần — 5 lần, trung vị (`results/benchmark-stages-q3-before-*.json`)
- 20.000 dòng: preflight 0,535s; đọc XLSX nguồn 2,587s; đọc danh mục 1,642s; preview lookup 0,261s; summary 0,181s; clean 0,132s; CSV 0,096s.
- 5.000 dòng: preflight 0,096s; đọc XLSX 0,396/0,292s; preview lookup 0,047s; summary 0,022s; clean 0,023s.
- Kết luận: các tầng Python thuần không phải nút cổ chai; thời gian nằm ở phiên Excel và trao đổi dữ liệu với bridge.

### Excel COM — 5 lần, trung vị (`results/benchmark-q3-before5000-r*.json`, `benchmark-q3-before20k-r*.json`)
- 5.000 dòng: preflight 0,122s; đọc nguồn 5,159s; đọc danh mục 4,745s; lookup 10,714s; summary 9,351s; clean 7,040s.
- 20.000 dòng: preflight 0,333s; đọc nguồn 5,072s; đọc danh mục 4,810s; lookup 19,886s; summary 16,797s; clean 11,987s.
- Giai đoạn ghi (write) 20k: lookup 7,669s; summary 4,917s; clean 3,368s. Giai đoạn mở/khởi động Excel 1,17–1,24s mỗi phiên. Đối soát (verify) 20k: lookup 0,748s; summary 1,244s; clean 0,046s.
- Phân rã lần đọc 20k: khởi động Excel 0,874/0,805s; mở file (phần còn lại của open) ~0,3s; find 0,222/0,203s; đọc Value2 0,050/0,033s; chuyển ma trận 0,236/0,165s; tính lại 0,009s. Phần còn lại (~3,4s/5,1s) là copy/preflight/hash và tuần tự hóa JSON sang PowerShell rồi Python.
- `Write-Grid` đã ghi theo khối từ trước: `paste_seconds` đo được 0,007s cho 20k×4; chi phí lớn nằm ở công thức (lookup 20k, summary 40k) và các bước COM/JSON, không phải vòng lặp ghi từng ô.

### Thay đổi đã cân nhắc nhưng KHÔNG giữ
- Bỏ `WrapText` + `Rows.AutoFit()` cho lưới lớn: đo được 0,29–0,55s/20k (khoảng 2,9% thời gian lookup), dưới 10% và làm đổi hình thức file → không sửa.
- Tái dùng một phiên Excel cho đọc nguồn + danh mục + thực thi: mỗi lần khởi động Excel ~0,8–1,2s; tổng 5 phiên ~4–6s trên ~60s một lượt benchmark (~8%), dưới 10%, nhưng phải giữ COM sống qua nhiều thao tác và xử lý lỗi/timeout phức tạp hơn → không giữ.
- Tối ưu JSON của bridge: chưa có cách đơn giản chứng minh ≥10% mà không đổi giao thức; giữ nguyên.

### So mốc cũ
- `benchmark-final-5000.json` (mốc trước): lookup 15,829s; summary 13,601s; clean 11,114s. Lượt này: 10,714/9,351/7,040s → không tệ hơn.
- Một lần summary 20k mất 261,68s và một lần lookup 5k mất 253,07s do Excel treo tạm thời; các lần còn lại 15–21s. Báo cáo dùng trung vị 5 lần và ghi rõ ngoại lệ, không giấu.
## Giới hạn
- Số đo phụ thuộc máy, Excel đang chạy nền và nhiệt độ CPU; lấy trung vị 5 lần.
- `benchmark.py` đo 3 thao tác riêng, không đo toàn bộ phiên GUI.
- Không đo cloud (không gọi mạng).


## Lượt chạy đầy đủ sau khi thêm đo đạc
- Lượt 1 (`run-20261102-102636.log`): 30 lệnh, **1 lỗi** — `test_app.py --integration` hết 496,1s rồi báo `Excel quá thời gian... đã thử tối đa 2 lần`; đây là kiểu treo Excel tạm thời đã gặp trong benchmark (253s/262s), không phải lỗi logic.
- Chạy riêng ngay sau đó: `test_app.py --integration` **105/105 đạt**.
- Lượt 2 (`run-20261102-104045.log`): **30 lệnh, 0 lỗi, 292,0s, exit 0**. `test_app --integration` 69,9s/105 check; `test_read_boundary` 4; `test_gui` 14; `test_ui_acceptance` 48; `benchmark.py` 30,4s. `examples/` không đổi, không còn Excel do bridge tạo, không có thư mục tạm mới.
- Kết luận Q3: các cổng test đạt; không giữ tối ưu nào vì mọi phương án đơn giản đều dưới 10% hoặc đổi hình thức/giao thức.
