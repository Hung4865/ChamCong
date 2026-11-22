# Chấm công từ RAW và Google Form

Chương trình nhận dữ liệu chấm công RAW (`.csv` hoặc `.xlsx`) và thông báo xin nghỉ từ Google Form (`.xlsx` hoặc `.csv`), rồi xuất hai file Excel: tổng hợp theo mẫu Final và chi tiết theo mẫu ChiTiet.

## Chạy trên Windows

1. Cài Python 3.10 trở lên.
2. Chạy `python -m pip install -r requirements.txt` trong thư mục chương trình.
3. Mở `Mo_chuong_trinh.cmd`, chọn hai file đầu vào và kỳ báo cáo, rồi bấm **Tạo báo cáo**.

File kết quả được lưu trong `ket_qua`. Mẫu tổng hợp, danh sách nhân viên và quy tắc nằm trong `cau_hinh`.

Xem [hướng dẫn chi tiết](HUONG_DAN.md) để tra quy tắc và cách xử lý các mục cần kiểm tra.

## Dữ liệu test tháng 11

Thư mục `ket_qua/test thang 11` chứa hai file đầu vào đã cung cấp:

- `RAW_MayChamCong_2026-10-25_2026-11-24.csv`
- `GG FORM xin nghi.xlsx`

Chọn tháng báo cáo `2026-11`, kỳ `25-24` để kiểm tra khoảng **25/10/2026–24/11/2026**. Đây là kỳ trong tên tệp; cần kiểm tra mức bao phủ dữ liệu thực tế trước khi chốt báo cáo.

## Kiểm tra quy tắc

```powershell
python -B -m unittest discover -s . -p test_monthly_report.py
```
