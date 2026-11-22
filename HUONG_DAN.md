# Tổng hợp chấm công hằng tháng

Mở Mo_chuong_trinh.cmd. Chọn file RAW CSV hoặc Excel XLSX và file xin nghỉ XLSX hoặc CSV. Kiểm tra tháng, ngày bắt đầu và ngày kết thúc rồi bấm Tạo báo cáo. Mỗi lần chạy xuất hai file trong thư mục ket_qua:

- TongHop_TuDong_YYYY-MM-DD_YYYY-MM-DD.xlsx: bảng tổng hợp theo mẫu Final.
- ChiTiet_TuDong_YYYY-MM-DD_YYYY-MM-DD.xlsx: bảng chi tiết theo bố cục ChiTiet, mỗi nhân viên một khối, các ngày nằm ngang.

Hai ngày trong tên là ngày đầu và ngày cuối kỳ, tính cả hai ngày. Hai file được tính từ cùng RAW và Form, cùng các quy tắc. Khi chạy bằng dòng lệnh với --out, tên file chi tiết thay tiền tố TongHop_ bằng ChiTiet_; nếu tên không có tiền tố này thì thêm _ChiTiet.

RAW Excel cần là bảng sự kiện, mỗi dòng một lượt quét, có cột time và employeeNoString. Tool cũng nhận tiêu đề tiếng Việt Ngày giờ và Mã nhân viên, hoặc tách thành Ngày, Giờ, Mã nhân viên. Có thể chọn trang dữ liệu sau trang ghi chú. Ô ngày giờ Excel và chuỗi ngày/tháng/năm đều được đọc về UTC+7.

Mẫu Final và danh sách nhân viên đã được lưu sẵn trong cau_hinh. Mỗi kỳ chỉ cần hai file dữ liệu đầu vào. Nếu tên kết quả đã tồn tại khi chạy qua cửa sổ, tool thêm _v2, _v3 vào tên.

## Kỳ và lịch làm việc

- Chế độ 25-24 lấy từ ngày 25 tháng trước đến ngày 24 tháng báo cáo.
- Chế độ calendar lấy trọn tháng dương lịch.
- Có thể sửa trực tiếp hai ô ngày theo dạng YYYY-MM-DD.
- Mặc định làm thứ Hai đến thứ Sáu. Nhập ngày lễ/ngày nghỉ bổ sung vào holidays trong cau_hinh/quy_tac.json, ví dụ ["2026-09-02"].
- RAW cần bao phủ đúng kỳ. Nếu kỳ là 25/9–24/10 thì chọn RAW của cả khoảng này. Tool đưa các ngày toàn đơn vị không có dữ liệu và kỳ RAW không đủ vào CanKiemTra.

## Cách tính các cột

| Cột | Cách tính |
|---|---|
| D | Ngày làm việc trong kỳ, loại cuối tuần và ngày nghỉ cấu hình |
| E | Ngày có đủ lượt vào sáng và ra chiều thực tế |
| F | Chỉ ngày có thông báo vắng cả ngày. Nghỉ sáng, chiều hoặc theo giờ tính 0 |
| G | Ngày đủ hai lượt và vào ≤08:03:00, ra ≥17:00:35 |
| H | Ngày đúng giờ, vắng cả ngày hoặc đạt phần còn lại sau thông báo. Mỗi ngày tối đa một lần |
| I | Ngày thiếu lượt cần thiết sau áp dụng phạm vi thông báo |
| J | D − H. Có thể gồm ngày chưa xác nhận do lỗi nguồn, cần kiểm tra trước khi chốt |
| K | Các mục Form của nhân viên đang cần kiểm tra |

Giờ vào lấy lượt sáng sớm nhất trước 13:00. Giờ ra lấy lượt chiều muộn nhất từ 13:00. Lượt trước 06:30 vẫn được dùng. Tool bỏ sự kiện invalid/thiếu mã, chuẩn hóa UTC+7 và gộp lượt trùng.

Thông báo nghỉ sáng miễn yêu cầu lượt vào sáng và kiểm tra lượt ra buổi chiều. Nghỉ chiều miễn yêu cầu lượt ra chiều và kiểm tra lượt vào buổi sáng. Thông báo theo giờ chỉ nới mốc vào/ra khi khoảng đã khai bao phủ đầu/cuối ca. Ví dụ thông báo 08:00–09:00 cho phép giờ vào tới 09:00, giờ ra vẫn cần đạt 17:00:35.

Hai thông báo sáng và chiều cùng ngày vẫn tính 0 ở F theo quy tắc đang dùng. H có thể ghi nhận thực hiện đúng sau thông báo.

## File xin nghỉ

Tool nhận trực tiếp file Google Form đã cung cấp. Các cột chính gồm họ tên, loại thông báo, vắng từ ngày, đến ngày, thời gian/buổi và nội dung.

- Họ tên có dấu, không dấu, chữ hoa/chữ thường được chuẩn hóa. Có thể thêm cột Mã nhân viên để ghép chính xác.
- Chọn rõ Cả ngày, Sáng, Chiều. Với nghỉ theo giờ, ghi khoảng rõ ràng hoặc thêm hai cột Từ giờ và Đến giờ dạng 08:00.
- Nếu chọn Sáng/Chiều và nội dung chỉ có một mốc giờ họp, tool sử dụng buổi đã chọn. Nếu nội dung có cả giờ bắt đầu và kết thúc, tool sử dụng khoảng đó.
- Khoảng ngày như "từ 5/10 đến 7/10" hoặc "15-16/10" được trải thành từng ngày.
- Ngày mâu thuẫn, tên chưa khớp hoặc nội dung như "xin muộn một chút" được ghi vào CanKiemTra. Sửa trong file xin nghỉ rồi chạy lại.
- Nếu có cột Trạng thái duyệt, tool áp dụng các dòng đã duyệt. Nếu không có cột này, tool áp dụng các phản hồi ghép được.

## Cấu hình dùng chung

- cau_hinh/NhanVien.csv lưu mã nhân viên, họ tên và phòng. Cập nhật khi thay đổi nhân sự.
- cau_hinh/quy_tac.json lưu mốc giờ, ngày nghỉ và tên thay thế. Mục name_aliases ánh xạ tên trên Form sang mã nhân viên trong danh sách. Chỉ thêm sau khi xác định đúng người.
- cau_hinh/Mau_TongHop.xlsx lưu bố cục bản Final. Các số chấm công cũ đã được xóa khỏi mẫu.
- date_corrections dành cho sửa ngày theo mã SHA256 của đúng file Form, giúp bản sửa ngày không áp nhầm sang file tháng khác. Thông thường chỉ cần sửa trực tiếp ngày trên Form.

## Các trang kết quả

File ChiTiet có trang Sheet1 giống bố cục mẫu: thông tin nhân viên, ngày, giờ vào, giờ ra, tăng ca, muộn, nghỉ sớm, đã tham dự, giờ nghỉ, trạng thái và tóm tắt. Kỳ dài hơn 31 ngày được chia thêm trang Ky_2... Giờ quét thực tế luôn giữ ở hai dòng giờ; lý do Form ghi ở dòng Trạng thái với (S) sáng, (C) chiều, (G) theo giờ. P là phép, CT công tác; # là ngày nghỉ và không tính vắng, A là thiếu lượt cần thiết. L là muộn, E là về sớm, W là đủ lượt cần thiết sau thông báo; W có L/E vẫn có thể chưa đạt giờ. Phút Đã tham dự là chênh lệch giờ vào–ra, chưa trừ nghỉ trưa. Tăng ca và Giờ Nghỉ giữ 0 theo mẫu. Trang ThongBaoForm và CanKiemTra giúp tra lý do và lỗi nguồn. Dòng tóm tắt mỗi người dùng cùng chỉ số với D–J của file tổng hợp.

BaoCaoChamCong có bố cục tổng hợp như Final. ChiTietNgay có toàn bộ ngày, giờ và trạng thái. FormDaGhep chỉ rõ từng dòng thông báo. CanKiemTra liệt kê các mục cần sửa. NguonQuyTac ghi file nguồn, mã kiểm tra và quy tắc đã dùng.

Kết quả có cảnh báo là bản tạm tính cho đến khi xử lý các mục ảnh hưởng. Thông báo chưa khớp được liệt kê để đối soát. Các báo cáo và đối chiếu trước đây được tra trong DanhMuc_TepChamCong.csv.

## Chạy dòng lệnh

python -B -X utf8 "D:/Cuc/vanban/cham-cong/monthly_report.py" --raw "D:/Cuc/cham cong v3/du_lieu_cham_cong_2509_2410_2026.csv" --form "D:/Users/Download/Untitled form (Responses).xlsx" --start 2026-09-25 --end 2026-10-24

Thêm --out để chọn đường dẫn kết quả. Máy cần Python 3.10 trở lên và openpyxl. Cài thư viện bằng python -m pip install openpyxl. Khi chuyển sang máy khác, sao chép cả thư mục tool gồm các file Python và cau_hinh.

Kiểm tra quy tắc bằng python -B -m unittest discover -s cham-cong -p test_monthly_report.py từ thư mục cha của tool.
