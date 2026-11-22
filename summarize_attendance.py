"""Tổng hợp kỳ chấm công và đối chiếu với bảng kết quả đã chốt."""

import argparse
from collections import Counter, defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path
import re
import sys

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import attendance_report as base


METRICS = (
    "Tổng số ngày làm việc", "Ngày điểm danh đủ dữ liệu", "Ngày vắng có lý do",
    "Ngày thực hiện đúng", "Ngày đúng gồm vắng có lý do", "Ngày thiếu dữ liệu",
    "Ngày thực hiện chưa đúng",
)
KINDS = {"nghi phep", "di cong tac", "di hoc", "du hop", "du hoi thao", "khac"}
COMPACT_RANGE = re.compile(
    r"\b(?:ngay\s+)?(\d{1,2})\s*[-–]\s*(\d{1,2})\s*[/]\s*(\d{1,2})(?:[/](\d{4}))?"
)


def date_range(start, end):
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def notification_dates(row, first):
    """Chỉ xác định ngày thông báo. Chỉ tiêu chuẩn đếm cả ngày khi có thông báo một buổi/giờ."""
    detail = base.string(row.get("detail"))
    ends = {d for field in ("end", "other_end") if (d := base.form_date(row.get(field)))}
    if len(ends) > 1:
        raise ValueError("Hai cột kết thúc khác nhau")
    end = next(iter(ends), None)
    ranges = base.RANGE_TEXT.findall(detail)
    compact = COMPACT_RANGE.findall(base.normalize(detail))
    if len(ranges) + len(compact) > 1:
        raise ValueError("Nội dung có nhiều khoảng ngày")
    if ranges:
        start_text, end_text = ranges[0]
        detail_end = base.detail_date(end_text, first.year)
        detail_start = base.detail_date(start_text, first.year, first.month)
        if len(re.split(r"\s*[/.-]\s*", end_text)) < 3 and detail_end.month < first.month:
            detail_end = detail_end.replace(year=first.year + 1)
    elif compact:
        day1, day2, month, year = compact[0]
        detail_start = date(int(year or first.year), int(month), int(day1))
        detail_end = date(int(year or first.year), int(month), int(day2))
    else:
        detail_start = detail_end = None
    if detail_start is not None:
        if detail_start != first or end is not None and end != detail_end:
            raise ValueError("Ngày trong nội dung khác các cột ngày")
        end = detail_end
    if end is None and re.search(r"\btu\s+(?:ngay\s+)?\d{1,2}[/.-]\d", base.normalize(detail)):
        raise ValueError("Thiếu ngày kết thúc của khoảng vắng")
    end = end or first
    if end < first:
        raise ValueError("Ngày kết thúc trước ngày bắt đầu")
    return end


def load_notifications(path, employees, start, end):
    rows = base.read_sheet(
        path,
        {"name": ("ho va ten",), "start": ("vang tu ngay", "tu ngay"),
         "kind": ("loai thong bao",)},
        {"end": ("den ngay",), "other_end": ("ngay thang ket thuc",),
         "period": ("thoi gian",), "detail": ("noi dung cu the",),
         "submitted": ("timestamp", "dau thoi gian")},
    )
    names = defaultdict(list)
    for employee in employees.values():
        names[base.normalize(employee.name)].append(employee.id)
    notifications, issues = defaultdict(list), []
    for number, row in rows:
        issue = [number, base.string(row["name"]), base.string(row["start"]),
                 base.string(row.get("kind")), base.string(row.get("period")),
                 base.string(row.get("detail"))]
        try:
            first = base.form_date(row["start"])
            submitted = base.form_date(row.get("submitted"))
            if first is None:
                raise ValueError("Thiếu ngày bắt đầu")
            if submitted and first.day <= 12 and first.day != first.month:
                swapped = date(first.year, first.day, first.month)
                if abs((first - submitted).days) > 30 and abs((swapped - submitted).days) <= 3:
                    if any(start <= day <= end for day in (first, swapped)):
                        raise ValueError("Có thể đảo ngày/tháng, cần xác nhận")
            last = notification_dates(row, first)
        except (ValueError, TypeError) as exc:
            if any(isinstance(row.get(field), (date, datetime))
                   and start <= base.form_date(row[field]) <= end
                   for field in ("start", "submitted")):
                issues.append([*issue, str(exc)])
            continue
        if first > end or last < start:
            continue
        matches = names[base.normalize(row["name"])]
        if len(matches) != 1:
            issues.append([*issue, "Họ tên không khớp duy nhất với danh sách nhân viên"])
            continue
        if base.normalize(row["kind"]) not in KINDS:
            issues.append([*issue, "Loại thông báo chưa xác định"])
            continue
        for day in date_range(max(first, start), min(last, end)):
            notifications[(matches[0], day)].append(
                (number, base.string(row["kind"]), base.string(row.get("period")),
                 base.string(row.get("detail"))))
    return notifications, issues


def load_reference(path, employees):
    book = load_workbook(path, data_only=True)
    try:
        sheet = book["BaoCaoChamCong"]
        names = {base.normalize(employee.name): employee.id for employee in employees.values()}
        records = []
        department = ""
        for row in range(5, sheet.max_row + 1):
            name = sheet.cell(row, 3).value
            if not name:
                continue
            identity = names.get(base.normalize(name))
            if not identity:
                raise ValueError(f"Tên trong báo cáo chuẩn không khớp danh sách tại dòng {row}")
            department = sheet.cell(row, 2).value or department
            values = [sheet.cell(row, column).value for column in range(4, 11)]
            if any(not isinstance(value, (int, float)) for value in values):
                raise ValueError(f"Báo cáo chuẩn chưa có giá trị số đầy đủ tại dòng {row}")
            records.append((row, identity, department, values))
        if len(records) != len(employees) or len({record[1] for record in records}) != len(records):
            raise ValueError("Danh sách trong báo cáo chuẩn thiếu hoặc trùng nhân viên")
        return records
    finally:
        book.close()


def calculate(employees, records, punches, notifications, days, latest_in, earliest_out):
    results, detail = {}, []
    for _, identity, _, _ in records:
        counts = Counter()
        for day in days:
            times = sorted(punches.get((identity, day), ()))
            ins = [stamp for stamp in times if stamp < time(13)]
            outs = [stamp for stamp in times if stamp >= time(13)]
            check_in, check_out = (ins[0] if ins else None), (outs[-1] if outs else None)
            paired = bool(check_in and check_out)
            punctual = bool(paired and check_in <= latest_in and check_out >= earliest_out)
            notices = notifications.get((identity, day), ())
            justified = bool(notices)
            missing = not paired and not justified
            counts["paired"] += paired
            counts["punctual"] += punctual
            counts["justified"] += justified
            counts["missing"] += missing
            detail.append([
                employees[identity].name, day, check_in, check_out, int(paired),
                int(punctual), int(justified), int(missing),
                "; ".join(str(n[0]) for n in notices),
                "; ".join(f"{n[1]} / {n[2]} / {n[3]}" for n in notices),
                "; ".join(stamp.isoformat(timespec="seconds") for stamp in times),
            ])
        included = counts["punctual"] + counts["justified"]
        results[identity] = [
            len(days), counts["paired"], counts["justified"], counts["punctual"],
            included, counts["missing"], len(days) - included,
        ]
    return results, detail


def add_table(book, name, headers, rows, widths):
    sheet = book.create_sheet(name)
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    for cell in sheet[1]:
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="213C57")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[1].height = 40
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial", size=11)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if isinstance(cell.value, str):
                cell.data_type = "s"
            elif isinstance(cell.value, (datetime, date)):
                cell.number_format = "dd/mm/yyyy"
            elif isinstance(cell.value, time):
                cell.number_format = "hh:mm:ss"
    return sheet


def export(args, employees, records, results, detail, issues, stats):
    book = load_workbook(args.reference)
    if book.sheetnames != ["BaoCaoChamCong"]:
        raise ValueError("Báo cáo chuẩn phải có duy nhất trang BaoCaoChamCong")
    sheet = book["BaoCaoChamCong"]
    sheet["B2"] = (
        f"BÁO CÁO TỔNG HỢP CHẤM CÔNG TỪ {args.start:%d/%m/%Y} ĐẾN {args.end:%d/%m/%Y}"
    )
    sheet["A3"] = (
        f"Tính từ {args.raw.name} và {args.form.name}. "
        f"Chỉ tính thứ Hai–thứ Sáu, loại ngày nghỉ cấu hình. "
        f"Đúng giờ khi vào ≤ {args.latest_in} và ra ≥ {args.earliest_out}. "
        "Một ngày có thông báo sáng/chiều/theo giờ được tính 1 lần ở cột F. "
        "E = đủ cặp quẹt. G = đủ cặp và đúng giờ. H = G + F. I = thiếu cặp và không có thông báo. J = D − H. "
        "H dùng phép cộng theo ghi chú báo cáo chuẩn, có thể đếm trùng ngày vừa đúng giờ vừa có thông báo. "
        "Xem Chi tiết ngày và Đối chiếu để kiểm tra."
    )
    sheet["A3"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    sheet.row_dimensions[3].height = 65
    differences, checks = [], []
    for row, identity, _, expected in records:
        values = results[identity]
        for column, value in enumerate(values, 4):
            sheet.cell(row, column, value)
        sheet.cell(row, 8, f"=G{row}+F{row}")
        sheet.cell(row, 10, f"=D{row}-H{row}")
        for offset, (actual, target) in enumerate(zip(values, expected)):
            differences.append([
                employees[identity].name, METRICS[offset], target, actual, actual - target,
                "Khớp" if actual == target else "Chênh lệch", f"BaoCaoChamCong!{get_column_letter(offset+4)}{row}",
            ])
        if expected[3] > expected[1]:
            checks.append([employees[identity].name, "Báo cáo chuẩn", "G > E",
                           f"G={expected[3]}, E={expected[1]}", f"G{row}/E{row}"])
        if expected[4] != expected[3] + expected[2]:
            checks.append([employees[identity].name, "Báo cáo chuẩn", "H khác G + F",
                           f"H={expected[4]}, G+F={expected[3]+expected[2]}", f"H{row}"])
        overlaps = sum(record[0] == employees[identity].name and record[5] and record[6] for record in detail)
        if overlaps:
            checks.append([employees[identity].name, "Tính từ nguồn", "Ngày vừa đúng giờ vừa có thông báo",
                           f"{overlaps} ngày. Cột H cộng cả G và F theo ghi chú báo cáo chuẩn.",
                           "Chi tiết ngày"])
        if values[4] > values[0]:
            checks.append([employees[identity].name, "Tính từ nguồn", "H lớn hơn tổng ngày làm việc",
                           f"H={values[4]}, D={values[0]}. Cần xác nhận cách cộng G và F.", f"H{row}"])
    add_table(book, "Đối chiếu", ["Họ tên", "Chỉ tiêu", "Bảng chuẩn", "Tính từ nguồn",
                                 "Chênh lệch", "Kết quả", "Ô trong bảng chuẩn"],
              differences, [30, 36, 16, 17, 15, 18, 32])
    add_table(book, "Chi tiết ngày", [
        "Họ tên", "Ngày", "Giờ vào", "Giờ ra", "Đủ dữ liệu", "Đúng giờ",
        "Vắng có lý do", "Thiếu dữ liệu sau loại vắng có lý do", "Dòng Form",
        "Thông báo", "Tất cả lượt quẹt đã lọc",
    ], detail, [30, 14, 14, 14, 14, 14, 16, 22, 18, 65, 65])
    add_table(book, "Form cần kiểm tra", [
        "Dòng Form", "Họ tên", "Ngày khai", "Loại", "Buổi", "Nội dung", "Vấn đề",
    ], issues, [12, 30, 22, 20, 15, 70, 55])
    add_table(book, "Kiểm tra quy tắc", ["Họ tên", "Nguồn", "Vấn đề", "Chi tiết", "Vị trí"],
              checks, [30, 22, 45, 75, 22])
    audit_rows = [[key, value] for key, value in stats.items()]
    audit_rows += [
        ["File RAW", str(args.raw.resolve())], ["File Form", str(args.form.resolve())],
        ["File danh sách", str(args.emp.resolve())], ["File chuẩn để đối chiếu", str(args.reference.resolve())],
        ["Vai trò file chuẩn", "Bố cục, thứ tự nhân viên và số liệu so sánh. Kết quả tính từ RAW và Form."],
        ["Kỳ tính", f"{args.start} đến {args.end}"],
        ["Điều kiện vào", f"≤ {args.latest_in}"], ["Điều kiện ra", f"≥ {args.earliest_out}"],
        ["Mốc chọn lượt quẹt", "Giờ vào sớm nhất trước 13:00, giờ ra muộn nhất từ 13:00"],
        ["Thông báo theo giờ hoặc nửa ngày", "Cột F tính 1 ngày nếu có thông báo trong ngày làm việc, theo ghi chú bảng chuẩn"],
        ["Trạng thái thông báo", "Ghi nhận phản hồi trong Form, dữ liệu Form không có cột xác nhận đã duyệt"],
        ["Ngày nghỉ", ", ".join(str(day) for day in args.holiday) or "Chỉ thứ Bảy và Chủ Nhật"],
    ]
    add_table(book, "Nguồn và quy tắc", ["Nội dung", "Giá trị"], audit_rows, [45, 110])
    if args.out.resolve() in {p.resolve() for p in (args.reference, args.raw, args.form, args.emp)}:
        raise ValueError("Đường dẫn kết quả trùng tệp nguồn")
    if args.out.suffix.lower() != ".xlsx":
        raise ValueError("Kết quả phải có đuôi .xlsx")
    with base.staged_file(args.out) as temporary:
        book.save(temporary)
    book.close()
    exact = sum(all(v == target for v, target in zip(results[eid], expected))
                for _, eid, _, expected in records)
    matches = [sum(results[eid][i] == expected[i] for _, eid, _, expected in records)
               for i in range(7)]
    print(f"Kết quả: {args.out.resolve()}")
    print(f"Nhân viên khớp toàn bộ: {exact}/{len(records)}")
    print("Khớp theo cột D–J: " + ", ".join(map(str, matches)))
    print(f"Phản hồi Form cần kiểm tra: {len(issues)}")
    print("Tổng tính từ nguồn D–J: " + ", ".join(str(sum(v[i] for v in results.values())) for i in range(7)))


def pick_arguments(cli):
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    try:
        selected = {}
        for field, title, filetypes in (
            ("raw", "Chọn CSV RAW chấm công", [("CSV", "*.csv")]),
            ("form", "Chọn phản hồi Google Form", [("Excel", "*.xlsx")]),
            ("reference", "Chọn báo cáo tổng hợp chuẩn để đối chiếu", [("Excel", "*.xlsx")]),
        ):
            value = filedialog.askopenfilename(title=title, filetypes=filetypes)
            if not value:
                return None
            selected[field] = value
        for field, label, default in (
            ("start", "Ngày bắt đầu (YYYY-MM-DD)", "2026-09-25"),
            ("end", "Ngày kết thúc (YYYY-MM-DD)", "2026-10-24"),
        ):
            value = simpledialog.askstring("Kỳ tổng hợp", label, initialvalue=default)
            if value is None:
                return None
            try:
                date.fromisoformat(value)
            except ValueError:
                messagebox.showerror("Ngày không hợp lệ", "Nhập ngày theo dạng YYYY-MM-DD.")
                return None
            selected[field] = value
        if not base.DEFAULT_ROSTER.is_file():
            selected["emp"] = filedialog.askopenfilename(
                title="Chọn danh sách nhân viên", filetypes=[("Excel", "*.xlsx")])
            if not selected["emp"]:
                return None
        selected["out"] = filedialog.asksaveasfilename(
            title="Lưu báo cáo tổng hợp và đối chiếu", defaultextension=".xlsx",
            initialfile=f"TongHop_Tool_{selected['start']}_{selected['end']}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not selected["out"]:
            return None
        return cli.parse_args([part for key, value in selected.items() for part in (f"--{key}", value)])
    finally:
        root.destroy()


def main():
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument("--raw", required=True, type=Path)
    cli.add_argument("--form", required=True, type=Path)
    cli.add_argument("--emp", type=Path, default=base.DEFAULT_ROSTER)
    cli.add_argument("--reference", required=True, type=Path)
    cli.add_argument("--out", type=Path, help="Mặc định TongHop_Tool_YYYY-MM-DD_YYYY-MM-DD.xlsx")
    cli.add_argument("--start", required=True, type=date.fromisoformat)
    cli.add_argument("--end", required=True, type=date.fromisoformat)
    cli.add_argument("--latest-in", type=time.fromisoformat, default=time(8, 3))
    cli.add_argument("--earliest-out", type=time.fromisoformat, default=time(17, 0, 35))
    cli.add_argument("--holiday", action="append", type=date.fromisoformat, default=[])
    args = cli.parse_args() if len(sys.argv) > 1 else pick_arguments(cli)
    if args is None:
        return
    if args.end < args.start:
        cli.error("Ngày kết thúc phải từ ngày bắt đầu trở đi")
    if args.out is None:
        args.out = Path(__file__).resolve().with_name(f"TongHop_Tool_{args.start}_{args.end}.xlsx")
    employees = base.load_employees(args.emp, "")
    records = load_reference(args.reference, employees)
    stats, clock_notes = Counter(), []
    dates, punches, unknown = base.load_clock(args.raw, employees, stats, clock_notes)
    issues = [["", "", "", "", "", "", note] for note in clock_notes]
    days = [day for day in date_range(args.start, args.end)
            if day.weekday() < 5 and day not in args.holiday]
    if not any(args.start <= day <= args.end for day in dates):
        raise ValueError("RAW không có dữ liệu trong kỳ")
    for (day, identity), count in unknown.items():
        if args.start <= day <= args.end:
            issues.append(["", identity, str(day), "", "", "", f"{count} sự kiện có mã ngoài danh sách"])
    notifications, form_issues = load_notifications(args.form, employees, args.start, args.end)
    issues.extend(form_issues)
    stats["Nhân viên"] = len(employees)
    stats["Ngày làm việc trong kỳ"] = len(days)
    stats["Lượt quẹt duy nhất trong kỳ"] = sum(
        len(times) for (_, day), times in punches.items() if args.start <= day <= args.end)
    results, detail = calculate(employees, records, punches, notifications, days,
                                args.latest_in, args.earliest_out)
    export(args, employees, records, results, detail, issues, stats)


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError) as error:
        sys.exit(f"Lỗi: {error}")
