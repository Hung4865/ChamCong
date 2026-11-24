# Chấm công từ RAW và Google Form

Chương trình nhận dữ liệu chấm công RAW (`.csv` hoặc `.xlsx`) và thông báo xin nghỉ từ Google Form (`.xlsx` hoặc `.csv`), rồi xuất hai file Excel: tổng hợp theo mẫu Final và chi tiết theo mẫu ChiTiet.

## 1. Cài đặt trên Windows

1. Tải repo bằng **Code → Download ZIP**, giải nén vào thư mục bất kỳ trên máy. Cài Python 3.10 trở lên.
2. Chạy `python -m pip install -r requirements.txt` trong thư mục chương trình.
3. Giữ nguyên cả thư mục tải về. File `Mo_chuong_trinh.cmd` gọi `monthly_report.py`; hai file `attendance_report.py`, `summarize_attendance.py` và thư mục `cau_hinh` cũng phải có để chạy.

Máy cần PowerShell và `curl.exe` để lấy RAW. Kết nối cùng mạng với máy chấm công hoặc qua VPN phù hợp. Nếu đã có RAW thì bỏ qua bước 2.

## 2. Lấy dữ liệu RAW từ máy chấm công

Giải nén chương trình ở đâu cũng được. Các file `Lay_RAW.cmd`, `lay_cong.ps1` và `dieu_kien.json` cần nằm cùng thư mục. Không cần tạo ổ D hoặc thư mục giống máy của người khác.

### Bản tải mới từ GitHub: điền tài khoản trực tiếp trong script

Mở `lay_cong.ps1` bằng Notepad. Ngay đầu file có ba dòng:

```powershell
$deviceUrl = 'http://DIA_CHI_MAY/ISAPI/AccessControl/AcsEvent?format=json'
$username = 'Dien ten dang nhap'
$password = 'Dien mat khau'
```

Thay địa chỉ API, tên đăng nhập và mật khẩu bằng thông tin bạn đã biết, giữ dấu nháy đơn, rồi lưu file. Nếu mật khẩu có ký tự nháy đơn, viết hai dấu liên tiếp, ví dụ `ab'cd` viết thành `'ab''cd'`. Không cần sao chép thêm file hoặc nhập mật khẩu mỗi lần chạy. Các tháng sau chỉ sửa ngày trong `dieu_kien.json`.

Bản trên GitHub chỉ chứa chữ `Dien mat khau`. Giữ mật khẩu thật trong bản dùng trên máy, không commit/push dòng đã điền mật khẩu lên GitHub.

### Mỗi tháng chỉ sửa kỳ trong dieu_kien.json

Ví dụ kỳ tháng 11/2026, tính cả ngày đầu và ngày cuối:

```json
"startTime": "2026-10-25T00:00:00+07:00",
"endTime": "2026-11-24T23:59:59+07:00"
```

Giữ nguyên các trường còn lại. Script tự phân trang từ vị trí 0 và không ghi đè ngày đã sửa trong `dieu_kien.json`.

**Cách chạy đơn giản:** lưu `dieu_kien.json` rồi bấm đúp **`Lay_RAW.cmd`**. File này tự tìm script trong cùng thư mục, không cần gõ đường dẫn. CSV RAW được lưu ngay tại thư mục đó.

Nếu muốn chạy bằng lệnh: mở thư mục chứa chương trình trong File Explorer, gõ `powershell` vào thanh địa chỉ rồi nhấn Enter. Cửa sổ PowerShell sẽ mở đúng thư mục; không cần lệnh `Set-Location`.

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
| `Lay_RAW.cmd` | Bấm đúp để lấy RAW, tự tìm script trong cùng thư mục |
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

