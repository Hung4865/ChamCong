"""Tạo báo cáo chuyên cần Excel .xlsx từ máy chấm công và Google Form."""

import argparse
import calendar
import csv
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.exceptions import InvalidFileException


LOCAL_TZ = timezone(timedelta(hours=7))
ROSTER_FILE = "NhanVien_DungTiengViet.xlsx"
DEFAULT_ROSTER = Path("D:/Cuc/cham cong v3") / ROSTER_FILE
DEFAULT_DEPT = "Vien CNS va CDS"
AM, PM = 1, 2
FULL_DAY = AM | PM
LABELS = ("Ngày:", "Đăng ký1", "Ra1", "Tăng ca", "Muộn", "Nghỉ sớm",
          "Đã tham dự", "Giờ Nghỉ", "Trạng thái")
HEADER_LABELS = ("ID Người", "Tên nhân viên", "Bộ phận", "Ngày vào làm", "Vị trí")
HEADER_SPANS = (1, 4, 3, 4, 3, 4, 3, 4, 3, 3)
SUMMARY_LABELS = ("Chuyên cần bình thường", "Cuối tuần", "Thoát", "Thời gian làm thêm",
                  "Tăng ca", "Thời gian làm việc", "Vắng mặt", "Số lần muộn",
                  "Số lần về sớm", "Thời gian nghỉ")
DATE_TEXT = r"\d{1,2}\s*[/.-]\s*\d{1,2}(?:\s*[/.-]\s*(?:\d{4}|\d{2}))?(?!\d)"
RANGE_TEXT = re.compile(
    rf"\btừ\s+(?:ngày\s+)?({DATE_TEXT}|\d{{1,2}})"
    rf"\s*(?:đến|tới|[-–])\s*(?:hết\s+)?(?:ngày\s+)?({DATE_TEXT})",
    re.IGNORECASE,
)
HOUR_TEXT = re.compile(r"\b\d{1,2}\s*(?:h(?:\d{2})?|giờ|:\d{2})(?!\w)|\b\d+\s*(?:p|phút)\b", re.I)


@dataclass(frozen=True)
class Employee:
    id: str
    name: str
    department: str


@dataclass(frozen=True)
class Day:
    check_in: str = "-"
    check_out: str = "-"
    late: int = 0
    early: int = 0
    worked: int = 0
    status: str = "-"


def string(value):
    return "" if value is None else str(value).strip()


def normalize(value):
    value = unicodedata.normalize("NFD", string(value).casefold())
    value = "".join(c for c in value if not unicodedata.combining(c))
    return " ".join(value.replace("đ", "d").split())


def employee_id(value):
    return str(int(value)) if isinstance(value, float) and value.is_integer() else string(value)


def read_sheet(path, required, optional):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in book:
            rows = sheet.values
            header = next(rows, ())
            names = list(map(normalize, header))
            columns = {}
            for field, choices in {**required, **optional}.items():
                for choice in choices:
                    matches = [i for i, name in enumerate(names)
                               if name == choice or name.startswith(choice + " (")]
                    if len(matches) > 1:
                        raise ValueError(f"{path.name} có nhiều cột khớp với «{choice}».")
                    if matches:
                        columns[field] = matches[0]
                        break
            if required.keys() <= columns.keys():
                return [
                    (number, {field: row[i] if i < len(row) else None
                              for field, i in columns.items()})
                    for number, row in enumerate(rows, 2) if any(v is not None for v in row)
                ]
    finally:
        book.close()
    headers = ", ".join(choices[0] for choices in required.values())
    raise ValueError(f"{path.name} thiếu trang tính có các cột cần thiết ({headers}).")


def load_employees(path, default_dept):
    rows = read_sheet(path, {"id": ("person id", "id nguoi"), "name": ("ho va ten",)},
                      {"department": ("don vi", "bo phan")})
    employees = {}
    for number, row in rows:
        identity, name = employee_id(row["id"]), string(row["name"])
        if not identity or not name:
            raise ValueError(f"Danh sách nhân viên dòng {number} thiếu mã hoặc họ tên.")
        if identity in employees:
            raise ValueError(f"Danh sách nhân viên dòng {number} trùng mã {identity}.")
        employees[identity] = Employee(identity, name, string(row.get("department")) or default_dept)
    if not employees:
        raise ValueError("Danh sách nhân viên trống.")
    return employees


def load_clock(path, employees, stats, notes):
    punches = defaultdict(set)
    dates = set()
    unknown = Counter()
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = csv.DictReader(stream)
        if not {"time", "employeeNoString"} <= set(rows.fieldnames or ()):
            raise ValueError("CSV cần có cột time và employeeNoString.")
        for row in rows:
            stats["Sự kiện CSV"] += 1
            raw_time = string(row["time"])
            try:
                stamp = datetime.fromisoformat(raw_time.replace("Z", "+00:00"))
                if len(raw_time) <= 10:
                    raise ValueError
                stamp = stamp.astimezone(LOCAL_TZ) if stamp.tzinfo else stamp
            except ValueError:
                notes.append(f"CSV dòng {rows.line_num} có thời điểm không hợp lệ.")
                continue
            day = stamp.date()
            dates.add(day)
            identity = employee_id(row["employeeNoString"])
            if not identity or normalize(row.get("currentVerifyMode")) == "invalid":
                stats["Sự kiện không hợp lệ hoặc thiếu mã"] += 1
                continue
            if identity not in employees:
                unknown[(day, identity)] += 1
                continue
            times = punches[(identity, day)]
            if stamp.time() in times:
                stats["Lượt quẹt trùng"] += 1
            times.add(stamp.time())
    return dates, punches, unknown


def report_period(dates, month, through):
    if month:
        start = date.fromisoformat(month + "-01")
    else:
        months = {day.replace(day=1) for day in dates}
        if len(months) != 1:
            raise ValueError("Hãy chọn tháng báo cáo (YYYY-MM) khi CSV trống hoặc có nhiều tháng.")
        start = months.pop()
    end = start.replace(day=calendar.monthrange(start.year, start.month)[1])
    if through is None:
        through = max((day for day in dates if start <= day <= end), default=None)
    if through is None or not start <= through <= end:
        raise ValueError("Hãy chọn ngày chốt dữ liệu nằm trong tháng báo cáo.")
    return start, end, through


def form_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if not value:
        return None
    text = string(value)
    try:
        return date.fromisoformat(text)
    except ValueError:
        return datetime.strptime(text, "%d/%m/%Y").date()


def detail_date(text, year, month=None):
    parts = [int(v) for v in re.split(r"\s*[/.-]\s*", text)]
    if len(parts) == 3:
        year = 2000 + parts[2] if parts[2] < 100 else parts[2]
    return date(year, parts[1] if len(parts) > 1 else month, parts[0])


def leave_interval(row, first):
    text = unicodedata.normalize("NFC", string(row.get("detail")))
    end_dates = {d for field in ("end", "other_end") if (d := form_date(row.get(field)))}
    if len(end_dates) > 1:
        raise ValueError("hai cột ngày kết thúc ghi khác nhau")
    last = next(iter(end_dates), None)
    ranges = RANGE_TEXT.findall(text)
    if len(ranges) > 1:
        raise ValueError("nội dung có nhiều khoảng ngày")
    if ranges:
        raw_start, raw_end = ranges[0]
        detail_end = detail_date(raw_end, first.year)
        detail_start = detail_date(raw_start, first.year, detail_end.month)
        if detail_start != first or (last is not None and last != detail_end):
            raise ValueError("ngày trong nội dung khác cột ngày của Form")
        last = detail_end
    elif last is None and re.search(r"\btừ\s+(?:ngày\s+)?\d{1,2}\s*[/.-]\s*\d", text, re.I):
        raise ValueError("nội dung có ngày bắt đầu nhưng thiếu ngày kết thúc rõ ràng")
    if last is not None and last < first:
        raise ValueError("ngày kết thúc nhỏ hơn ngày bắt đầu")
    period = normalize(row.get("period"))
    if period in {"sang", "chieu", "ca ngay"}:
        coverage = {"sang": AM, "chieu": PM, "ca ngay": FULL_DAY}[period]
        if HOUR_TEXT.search(text):
            raise ValueError("nội dung ghi giờ/phút cụ thể, cần xác nhận phạm vi thông báo")
    elif period in {"khac", ""} and (ranges or last is not None and last > first):
        if HOUR_TEXT.search(text) or re.search(r"\b(sáng|chiều|trưa)\b", text, re.I):
            raise ValueError("khoảng ngày có ghi buổi/giờ trong nội dung")
        coverage = FULL_DAY
    else:
        raise ValueError("chưa xác định được buổi vắng")
    return last or first, coverage, bool(ranges and not end_dates)


def load_leaves(path, employees, start, through, stats, notes):
    rows = read_sheet(
        path, {"name": ("ho va ten",), "start": ("vang tu ngay", "tu ngay"),
               "period": ("thoi gian",)},
        {"end": ("den ngay",), "other_end": ("ngay thang ket thuc",),
         "detail": ("noi dung cu the",), "submitted": ("timestamp", "dau thoi gian")},
    )
    names = defaultdict(list)
    for employee in employees.values():
        names[normalize(employee.name)].append(employee)
    leaves = defaultdict(int)
    for number, row in rows:
        try:
            first = form_date(row["start"])
            submitted = form_date(row.get("submitted"))
        except ValueError:
            notes.append(f"Form dòng {number} có ngày không hợp lệ.")
            continue
        if first is None:
            notes.append(f"Form dòng {number} thiếu ngày bắt đầu.")
            continue
        # Phát hiện ngày có khả năng bị đảo ngày/tháng khi gửi Form.
        if submitted and first.day <= 12 and first.day != first.month:
            swapped = date(first.year, first.day, first.month)
            if (abs((first - submitted).days) > 30 and abs((swapped - submitted).days) <= 3
                    and any(start <= day <= through for day in (first, swapped))):
                notes.append(f"Form dòng {number} có thể bị đảo ngày và tháng ({first:%d/%m/%Y}).")
                continue
        if first > through:
            continue
        try:
            last, coverage, parsed_text = leave_interval(row, first)
        except ValueError as exc:
            if first >= start or (submitted and start <= submitted <= through):
                notes.append(f"Form dòng {number} cần kiểm tra ({exc}).")
            continue
        if last < start:
            continue
        matches = names[normalize(row["name"])]
        if len(matches) != 1:
            reason = "trùng họ tên trong danh sách" if matches else "chưa khớp họ tên với danh sách"
            notes.append(f"Form dòng {number} {reason}.")
            continue
        stats["Phản hồi Form đã dùng"] += 1
        if parsed_text:
            stats["Phản hồi có ngày kết thúc đọc từ nội dung"] += 1
        day = max(first, start)
        while day <= min(last, through):
            leaves[(matches[0].id, day)] |= coverage
            day += timedelta(days=1)
    return leaves


def minutes_between(earlier, later):
    zero = date.min
    return (datetime.combine(zero, later) - datetime.combine(zero, earlier)).total_seconds() / 60


def penalty(minutes, grace):
    return int(minutes + 0.5) if minutes > 0 and minutes >= grace else 0


def day_result(times, leave, off, shift_start, shift_end, split, grace):
    check_in = min((stamp for stamp in times if stamp < split), default=None)
    check_out = max((stamp for stamp in times if stamp >= split), default=None)
    worked = int(minutes_between(check_in, check_out) + 0.5) if check_in and check_out else 0
    late = penalty(minutes_between(shift_start, check_in), grace) if check_in and not (leave & AM or off) else 0
    early = penalty(minutes_between(check_out, shift_end), grace) if check_out and not (leave & PM or off) else 0
    if off:
        status = "#"
    elif leave:
        missing = (not leave & AM and check_in is None) or (not leave & PM and check_out is None)
        status = "LV-A" if missing else "LV"
    else:
        status = "W" if check_in and check_out else "A"
    status = ("L-" if late else "") + status + ("-E" if early else "")
    return Day(check_in.isoformat(timespec="seconds") if check_in else "-",
               check_out.isoformat(timespec="seconds") if check_out else "-",
               late, early, worked, status)


def write_row(sheet, row, values, spans, *, bold=False, yellow=False, center=False, wrap=False):
    edge = Side(style="thin", color="FF808080")
    border = Border(top=edge, right=edge, bottom=edge, left=edge)
    font = Font(name="Calibri", size=11, bold=bold)
    fill = PatternFill("solid", fgColor="FFFFFF00")
    column = 1
    for value, span in zip(values, spans):
        target = sheet.cell(row, column, value)
        if isinstance(value, str):
            target.data_type = "s"
        elif isinstance(value, time):
            target.number_format = "hh:mm:ss"
        target.border = border
        target.font = font
        target.alignment = Alignment(
            horizontal="center" if center or (column > 1 and not bold and span == 1) else "left",
            vertical="center", wrap_text=wrap,
        )
        if yellow:
            target.fill = fill
        if span > 1:
            sheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + span - 1)
        column += span


def employee_rows(employee, days, length):
    info = (employee.id, employee.name, employee.department, "-", "-")
    yield [value for pair in zip(HEADER_LABELS, info) for value in pair]
    data = zip(
        LABELS,
        (list(range(1, length + 1)) + ["-"] * (31 - length),
         [time.fromisoformat(d.check_in) if d.check_in != "-" else "-" for d in days],
         [time.fromisoformat(d.check_out) if d.check_out != "-" else "-" for d in days],
         [0] * 31,
         [d.late for d in days], [d.early for d in days], [d.worked for d in days],
         [0] * 31, [d.status for d in days]),
    )
    for label, values in data:
        yield [label, *values]
    status_codes = [day.status.split("-") for day in days]
    totals = (
        sum("W" in status for status in status_codes), sum(d.status == "#" for d in days),
        0, 0, 0, sum(d.worked for d in days), sum("A" in status for status in status_codes),
        sum(d.late > 0 for d in days), sum(d.early > 0 for d in days), 0,
    )
    summary = " ".join(f"{label}:{value};" for label, value in zip(SUMMARY_LABELS, totals))
    if any("LV" in status for status in status_codes):
        summary += f" Thông báo vắng:{sum('LV' in status for status in status_codes)};"
    yield ["Tóm tắt", summary]


def render_report(employees, punches, leaves, start, end, through, args, stats):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.column_dimensions["A"].width = 17
    sheet.column_dimensions.group("B", "AF", outline_level=0, hidden=False)
    sheet.column_dimensions["B"].width = 10.5
    sheet.sheet_format.defaultRowHeight = 16
    titles = ("", "Báo cáo chuyên cần hàng tháng", f"Từ: {start} 00:00:00 Đến: {end} 23:59:59")
    if through != end:
        titles = (*titles[:-1], f"{titles[-1]} (chốt đến {through:%d/%m/%Y})")
    for row, title in enumerate(titles, 2):
        write_row(sheet, row, [title], [32], bold=True, center=True)
    row = 5
    for identity, employee in employees.items():
        days = []
        for day_number in range(1, through.day + 1):
            day = start.replace(day=day_number)
            off = day in args.holiday or args.weekends_off and day.weekday() >= 5
            times = punches.get((identity, day), ())
            leave = leaves.get((identity, day), 0)
            result = day_result(times, leave, off, args.shift_start, args.shift_end, args.split, args.grace)
            stats["Ngày có đúng một đầu quẹt"] += (result.check_in == "-") != (result.check_out == "-")
            stats["Ngày có thông báo vắng được áp dụng"] += "LV" in result.status.split("-")
            days.append(result)
        days.extend([Day()] * (31 - len(days)))
        block = employee_rows(employee, days, end.day)
        write_row(sheet, row, next(block), HEADER_SPANS, bold=True, yellow=True)
        sheet.row_dimensions[row].height = 19
        for offset, values in enumerate(block, 1):
            spans = (1, 31) if offset == 10 else (1,) * 32
            write_row(sheet, row + offset, values, spans)
        row += 11
    legend = (
        "Lưu ý: LV = Có thông báo vắng; LV-A = Có thông báo một buổi và thiếu lượt quẹt "
        "ở buổi còn lại; L = Muộn; E = Nghỉ sớm; W = Đã tham dự; A = Vắng mặt/thiếu lượt quẹt; "
        "# = Ngày nghỉ theo lịch; OT1/OT2/OT3 = Tăng ca. "
        "Đơn vị: phút. Giờ vào/ra hiển thị HH:MM:SS."
    )
    generated = datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S")
    write_row(sheet, row, [legend], [13], wrap=True)
    sheet.row_dimensions[row].height = 40
    write_row(sheet, row + 1, [f"Ngày/Giờ: {generated}"], [5])
    return workbook


@contextmanager
def staged_file(path):
    with NamedTemporaryFile(dir=path.parent, delete=False) as stream:
        temporary = Path(stream.name)
    try:
        yield temporary
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def generate(args):
    if args.out is not None and args.out.suffix.lower() != ".xlsx":
        raise ValueError("Tệp báo cáo cần mang đuôi .xlsx.")
    if not args.shift_start < args.split < args.shift_end or args.grace < 0:
        raise ValueError("Giờ bắt đầu < giờ tách buổi < giờ kết thúc, ngưỡng phút phải không âm.")
    stats, notes = Counter(), []
    employees = load_employees(args.emp, args.department)
    dates, punches, unknown = load_clock(args.raw, employees, stats, notes)
    start, end, through = report_period(dates, args.month, args.through)
    if args.out is None:
        args.out = Path(__file__).resolve().with_name(f"ChiTiet_Tool_{start}_{through}.xlsx")
    review = args.out.with_name(f"DoiSoat_Tool_{start}_{through}.txt")
    if {args.out.resolve(), review.resolve()} & {p.resolve() for p in (args.raw, args.form, args.emp)}:
        raise ValueError("Hãy chọn đường dẫn đầu ra khác đường dẫn tệp nguồn.")
    for (day, identity), count in sorted(unknown.items()):
        if start <= day <= through:
            notes.append(f"CSV ngày {day:%d/%m/%Y} có {count} lượt của mã {identity} ngoài danh sách.")
    leaves = load_leaves(args.form, employees, start, through, stats, notes)
    stats["Nhân viên trong danh sách"] = len(employees)
    stats["Lượt quẹt duy nhất trong kỳ sau khi lọc"] = sum(
        len(times) for (_, day), times in punches.items() if start <= day <= through
    )
    report = render_report(employees, punches, leaves, start, end, through, args, stats)
    review_text = [
        f"Đối soát kỳ {start:%d/%m/%Y} đến {through:%d/%m/%Y}",
        f"CSV: {args.raw.name}", f"Form: {args.form.name}", f"Danh sách: {args.emp.name}",
        "", *(f"{label}: {count}" for label, count in stats.items()),
        "", f"Số mục cần kiểm tra: {len(notes)}", *notes, "",
    ]
    with staged_file(args.out) as temporary:
        report.save(temporary)
    with staged_file(review) as temporary:
        temporary.write_text("\n".join(review_text), encoding="utf-8-sig")
    return review, len(notes)


def parser():
    cli = argparse.ArgumentParser(description=__doc__)
    local_roster = Path(__file__).resolve().with_name(ROSTER_FILE)
    cli.add_argument("--raw", required=True, type=Path, help="CSV chấm công")
    cli.add_argument("--form", required=True, type=Path, help="XLSX phản hồi Google Form")
    cli.add_argument("--emp", type=Path, default=local_roster if local_roster.is_file() else DEFAULT_ROSTER)
    cli.add_argument("--out", type=Path, help="Mặc định ChiTiet_Tool_YYYY-MM-DD_YYYY-MM-DD.xlsx")
    cli.add_argument("--month", help="Tháng báo cáo YYYY-MM")
    cli.add_argument("--through", type=date.fromisoformat, help="Ngày chốt YYYY-MM-DD")
    cli.add_argument("--shift-start", type=time.fromisoformat, default=time(8))
    cli.add_argument("--shift-end", type=time.fromisoformat, default=time(17))
    cli.add_argument("--split", type=time.fromisoformat, default=time(13), help="Mốc tách lượt vào và ra")
    cli.add_argument("--grace", type=int, default=5, help="Ngưỡng tính muộn/về sớm, đơn vị phút")
    cli.add_argument("--weekends-off", action="store_true", help="Đánh dấu thứ Bảy và Chủ Nhật là ngày nghỉ")
    cli.add_argument("--holiday", type=date.fromisoformat, action="append", default=[])
    cli.add_argument("--department", default=DEFAULT_DEPT, help="Bộ phận khi danh sách để trống")
    return cli


def pick_files(cli):
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    try:
        raw = filedialog.askopenfilename(title="Chọn CSV chấm công", filetypes=[("CSV", "*.csv")])
        if not raw:
            return
        form = filedialog.askopenfilename(title="Chọn XLSX phản hồi Google Form",
                                          filetypes=[("Excel", "*.xlsx")])
        if not form:
            return
        roster = cli.get_default("emp")
        if not roster.is_file():
            roster = filedialog.askopenfilename(title="Chọn XLSX danh sách nhân viên",
                                                filetypes=[("Excel", "*.xlsx")])
            if not roster:
                return
        try:
            employees = load_employees(Path(roster), DEFAULT_DEPT)
            dates, _, _ = load_clock(Path(raw), employees, Counter(), [])
            start, _, through = report_period(dates, None, None)
            output = filedialog.asksaveasfilename(
                title="Lưu báo cáo chuyên cần", defaultextension=".xlsx",
                initialdir=Path(raw).parent,
                initialfile=f"ChiTiet_Tool_{start}_{through}.xlsx",
                filetypes=[("Báo cáo Excel", "*.xlsx")],
            )
            if not output:
                return
            args = cli.parse_args(["--raw", raw, "--form", form, "--emp", str(roster), "--out", output])
            review, issues = generate(args)
        except (OSError, ValueError, BadZipFile, InvalidFileException) as exc:
            messagebox.showerror("Lỗi tạo báo cáo", str(exc))
            return
        messagebox.showinfo("Đã tạo báo cáo", f"{args.out}\n\nTệp đối soát ({issues} mục cần kiểm tra)\n{review}")
    finally:
        root.destroy()


def main():
    cli = parser()
    if len(sys.argv) == 1:
        pick_files(cli)
        return
    args = cli.parse_args()
    try:
        review, issues = generate(args)
    except (OSError, ValueError, BadZipFile, InvalidFileException) as exc:
        sys.exit(f"Lỗi: {exc}")
    print(f"Báo cáo: {args.out.resolve()}")
    print(f"Đối soát: {review.resolve()} ({issues} mục cần kiểm tra)")


if __name__ == "__main__":
    main()
