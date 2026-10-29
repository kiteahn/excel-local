# NEXT — việc tiếp theo (cập nhật sau mỗi vòng)

## Vòng tiếp theo: Vòng 6 — Thu hẹp `except Exception` (VĐ-05, mục C của PROMPT)

Hiện trạng: app.py 20 chỗ, ai_client.py 10 chỗ `except Exception`. Làm từng
cụm nhỏ, không sửa lan; mỗi chỗ phải trả lời được: lỗi gì thật sự có thể xảy
ra ở đây, và hành vi đúng là gì (báo lỗi / bỏ qua có chủ đích / rơi xuống
phương án dự phòng).

Làm:
1. Liệt kê từng chỗ kèm ngữ cảnh (`grep -n "except Exception" app.py ai_client.py`
   + đọc xung quanh). Phân loại: (a) nuốt lỗi im lặng — nguy hiểm nhất, xử lý
   trước; (b) bắt rộng nhưng đã có xử lý hợp lý — thu hẹp loại lỗi cụ thể;
   (c) bắt rộng ở ranh giới ngoài (GUI/CLI) — giữ nhưng bảo đảm có log/thông báo.
2. Chỉ sửa cụm (a) và các chỗ (b) rõ ràng trong vòng này (tối đa ~5-8 chỗ),
   mỗi chỗ kèm test hoặc kịch bản tái hiện.
3. Kiểm tra sau mỗi cụm: test_app 79, edge 61, router 7, security 38,
   privacy corpus 53 mẫu, cli 6 — tất cả phải đạt như trước.
4. Commit, cập nhật UPGRADE_LOG (liệt kê chỗ đã sửa + loại lỗi mới), NEXT
   chuyển sang Vòng 7 (tách hàm dài — VĐ-06).

## Hàng đợi sau đó
- Vòng 7: VĐ-06 — tách hàm dài (infer_smart_plan ~478 dòng là ứng viên đầu:
  tách theo nhóm case, giữ nguyên hành vi, test sau mỗi bước).
- Vòng 8: VĐ-07 — đo hiệu năng phân giai đoạn (benchmark.py; tách phần đo
  được trên Linux khỏi phần cần Excel COM).
- Vòng 9: VĐ-09 — dọn repo: gom *.json kết quả/benchmark vào `results/`,
  script phụ vào `tools/` (cập nhật đường dẫn ghi file trong test tương ứng).
- Vòng 10: VĐ-10 giao diện nhẹ + vòng tổng kết 10 vòng (chấm lại điểm Bước 0).

## Câu hỏi chờ người dùng (ghi lại, vẫn làm việc khác — không chờ)
- Chính sách mặc định mong muốn là Smart Router (có cloud) hay Ollama-only?
  Hiện code để người dùng chọn (mặc định config `smart_local`); tài liệu đã
  sửa đúng theo code, chưa đổi chính sách.
