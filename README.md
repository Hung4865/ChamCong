# Chấm công từ RAW và Google Form

Chương trình nhận dữ liệu chấm công RAW (`.csv` hoặc `.xlsx`) và thông báo xin nghỉ từ Google Form (`.xlsx` hoặc `.csv`), rồi xuất hai file Excel: tổng hợp theo mẫu Final và chi tiết theo mẫu ChiTiet.

## 1. Cài đặt trên Windows

1. Cài Python 3.10 trở lên.
2. Chạy `python -m pip install -r requirements.txt` trong thư mục chương trình.
3. Giữ nguyên cả thư mục tải về. File `Mo_chuong_trinh.cmd` gọi `monthly_report.py`; hai file `attendance_report.py`, `summarize_attendance.py` và thư mục `cau_hinh` cũng phải có để chạy.

Máy cần PowerShell và `curl.exe` để lấy RAW. Kết nối cùng mạng với máy chấm công hoặc qua VPN phù hợp. Nếu đã có RAW thì bỏ qua bước 2.

## 2. Lấy dữ liệu RAW từ máy chấm công

### Thiết lập kết nối lần đầu

1. Sao chép `ket_noi.example.json` thành `ket_noi.local.json` trong cùng thư mục với `lay_cong.ps1`.
2. Sửa `url` thành địa chỉ thật của máy, giữ đường dẫn `/ISAPI/AccessControl/AcsEvent?format=json`. Sửa `username` thành tên đăng nhập. Không ghi mật khẩu trong JSON.
3. Lần tải đầu, PowerShell hỏi mật khẩu. Sau khi tải thành công, tài khoản được lưu mã hóa trong `tai_khoan.local.xml`, chỉ dùng được với tài khoản Windows và máy đã lưu. Khi đổi mật khẩu hoặc tài khoản, xóa file XML này rồi chạy lại.

Hai file `ket_noi.local.json`, `tai_khoan.local.xml` đã được bỏ qua trong `.gitignore`, không đưa lên GitHub. Cấu hình kết nối của bộ lấy RAW cũ không bị thay đổi; bản trong repo cần thiết lập lần đầu như trên.

### Mỗi tháng chỉ sửa kỳ trong dieu_kien.json

Ví dụ kỳ tháng 11/2026, tính cả ngày đầu và ngày cuối:

```json
"startTime": "2026-10-25T00:00:00+07:00",
"endTime": "2026-11-24T23:59:59+07:00"
```

Giữ nguyên các trường còn lại. Script tự phân trang từ vị trí 0 và không ghi đè ngày đã sửa trong `dieu_kien.json`.

Mở PowerShell tại thư mục chương trình. Kiểm tra ngày mà chưa kết nối máy:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\lay_cong.ps1 -KiemTra
```

Tải RAW:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\lay_cong.ps1
```

Kết quả nằm cạnh script: `RAW_MayChamCong_2026-10-25_2026-11-24.csv`. Nếu tệp cùng kỳ đã tồn tại, script thay thế sau khi tải đủ và kiểm tra dữ liệu thành công. Khi tải lỗi, CSV cũ được giữ lại. Chỉ chốt kỳ sau khi đã có dữ liệu hết ngày cuối, kể cả lượt ra chiều.

## 3. Chuẩn bị file xin nghỉ

Tải bảng phản hồi Google Form về Excel `.xlsx` hoặc CSV. Giữ các cột họ tên, loại thông báo, buổi/thời gian, ngày bắt đầu, ngày kết thúc và nội dung. Dùng dữ liệu của toàn kỳ cần báo cáo.

## 4. Xuất báo cáo

1. Mở `Mo_chuong_trinh.cmd`.
2. Chọn RAW vừa tải và file xin nghỉ Google Form.
3. Nhập tháng báo cáo, ví dụ `2026-11`, chọn kỳ `25-24`. Kiểm tra hai ô ngày là `2026-10-25` và `2026-11-24`. Chọn `calendar` nếu cần tháng dương lịch.
4. Chọn thư mục kết quả và bấm **Tạo báo cáo**.
5. Kiểm tra trang `CanKiemTra`, sửa các mục lỗi nguồn rồi chạy lại trước khi chốt.

Hai file kết quả:

- `TongHop_TuDong_2026-10-25_2026-11-24.xlsx`: tổng hợp theo bố cục Final.
- `ChiTiet_TuDong_2026-10-25_2026-11-24.xlsx`: từng nhân viên với ngày nằm ngang, giờ vào/ra, trạng thái và lý do Form theo mẫu ChiTiet.

Nếu tên đã tồn tại khi chạy bằng cửa sổ, chương trình thêm `_v2`, `_v3`... Hai ngày trong tên là ngày đầu và cuối kỳ, tính cả hai ngày.

File kết quả được lưu trong `ket_qua`. Mẫu tổng hợp, danh sách nhân viên và quy tắc nằm trong `cau_hinh`.

Xem [hướng dẫn chi tiết](HUONG_DAN.md) để tra quy tắc và cách xử lý các mục cần kiểm tra.

Quy tắc chính: thứ Bảy/Chủ nhật không tính vắng; chọn lượt sáng sớm nhất trước 13:00 và lượt chiều muộn nhất từ 13:00, có lấy lượt trước 06:30; đúng giờ khi vào ≤08:03:00 và ra ≥17:00:35. Nghỉ sáng, chiều hoặc theo giờ tính 0 ngày nghỉ, nhưng ngày vẫn được tính thực hiện đúng nếu phần còn lại đạt yêu cầu sau thông báo hợp lệ. Chỉ nghỉ cả ngày mới tính vào số ngày vắng có lý do.

## Các file phải giữ cùng nhau

| File/thư mục | Vai trò |
|---|---|
| `lay_cong.ps1`, `dieu_kien.json` | Tải RAW theo kỳ |
| `ket_noi.example.json` | Mẫu cấu hình kết nối, sao chép thành file local để sử dụng |
| `Mo_chuong_trinh.cmd` | Nút mở chương trình báo cáo |
| `monthly_report.py` | Giao diện, ghép dữ liệu và xuất hai báo cáo |
| `attendance_report.py`, `summarize_attendance.py` | Các hàm dùng chung, không được xóa |
| `cau_hinh` | Mẫu Excel, danh sách nhân viên và quy tắc |
| `requirements.txt` | Thư viện Python cần cài |

## Dữ liệu test tháng 11

Thư mục `ket_qua/test thang 11` chứa hai file đầu vào đã cung cấp:

- `RAW_MayChamCong_2026-10-25_2026-11-24.csv`
- `GG FORM xin nghi.xlsx`

Chọn tháng báo cáo `2026-11`, kỳ `25-24` để kiểm tra khoảng **25/10/2026–24/11/2026**. Đây là kỳ trong tên tệp; cần kiểm tra mức bao phủ dữ liệu thực tế trước khi chốt báo cáo.

## Kiểm tra quy tắc

```powershell
python -B -m unittest discover -s . -p test_monthly_report.py
```
