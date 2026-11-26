"""Hai tệp RAW và xin nghỉ tạo báo cáo tổng hợp theo mẫu đã cấu hình."""

import argparse
import calendar
import csv
import hashlib
import json
import os
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
import queue
import re
import shutil
import sys
import threading
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

import attendance_report as base
import summarize_attendance as legacy


ROOT = Path(__file__).resolve().parent
CONFIG = ROOT / "cau_hinh" / "quy_tac.json"
COMPACT = re.compile(r"\b(?:ngay\s+)?(\d{1,2})\s*[-–]\s*(\d{1,2})/(\d{1,2})(?:/(\d{4}))?")
CLOCK = re.compile(r"(?<!\d)([0-2]?\d)\s*(?:h|giờ|:)\s*([0-5]\d)?", re.I)
KINDS = legacy.KINDS
FORM_REQUIRED = {
    "name": ("ho va ten", "ho ten"), "start": ("vang tu ngay", "tu ngay"),
    "kind": ("loai thong bao", "loai"),
}
FORM_OPTIONAL = {
    "id": ("ma nhan vien", "person id"), "end": ("den ngay",),
    "other_end": ("ngay thang ket thuc",), "period": ("thoi gian", "buoi"),
    "detail": ("noi dung cu the", "noi dung"), "submitted": ("timestamp", "dau thoi gian"),
    "from_hour": ("tu gio",), "to_hour": ("den gio",), "approval": ("trang thai duyet",),
}
RAW_FIELDS = ("time", "employeeNoString", "name", "cardReaderNo", "doorNo", "currentVerifyMode")
RAW_ALIASES = {
    "time": ("time", "timestamp", "thoi gian", "ngay gio", "thoi gian cham cong"),
    "employeeNoString": ("employeenostring", "employee no string", "employee id", "employee no",
                         "person id", "personid", "ma nhan vien", "manhanvien", "ma nv", "id nguoi"),
    "name": ("name", "ho va ten", "ho ten"),
    "cardReaderNo": ("cardreaderno", "card reader no"),
    "doorNo": ("doorno", "door no"),
    "currentVerifyMode": ("currentverifymode", "current verify mode"),
    "date": ("date", "ngay", "ngay cham cong"),
    "hour": ("hour", "gio", "gio cham cong"),
}


@dataclass(frozen=True)
class Notice:
    row: int
    kind: str
    period: str
    detail: str
    scope: str
    begin: time | None = None
    finish: time | None = None


def load_setup(path):
    settings = json.loads(path.read_text(encoding="utf-8-sig"))
    for field in ("morning_cutoff", "latest_in", "earliest_out", "shift_start", "shift_end"):
        settings[field] = time.fromisoformat(settings[field])
    if not (settings["shift_start"] <= settings["latest_in"] < settings["morning_cutoff"]
            < settings["shift_end"] <= settings["earliest_out"]):
        raise ValueError("Các mốc giờ trong cấu hình chưa đúng thứ tự.")
    if settings["period_mode"] not in ("25-24", "calendar"):
        raise ValueError("Kỳ mặc định trong cấu hình phải là 25-24 hoặc calendar.")
    settings["holidays"] = {date.fromisoformat(value) for value in settings["holidays"]}
    if not set(settings["workdays"]) <= set(range(7)):
        raise ValueError("Ngày làm việc phải là số từ 0 (thứ Hai) đến 6 (Chủ nhật).")
    employees = {}
    with (path.parent / settings["employees"]).open(encoding="utf-8-sig", newline="") as stream:
        for row in csv.DictReader(stream):
            identity, name = base.string(row["MaNhanVien"]), base.string(row["HoTen"])
            if not identity or not name or identity in employees:
                raise ValueError("Danh sách nhân viên thiếu mã/tên hoặc trùng mã.")
            employees[identity] = base.Employee(identity, name, base.string(row["Phong"]))
    if not employees:
        raise ValueError("Danh sách nhân viên trống.")
    aliases = {base.normalize(name): str(identity) for name, identity in settings["name_aliases"].items()}
    if not set(aliases.values()) <= set(employees):
        raise ValueError("Tên thay thế trong cấu hình trỏ tới mã nhân viên chưa có.")
    settings["name_aliases"] = aliases
    settings["template_path"] = path.parent / settings["template"]
    return settings, employees


def period_dates(month, mode):
    first = date.fromisoformat(month + "-01")
    if mode == "calendar":
        return first, first.replace(day=calendar.monthrange(first.year, first.month)[1])
    return (first - timedelta(days=1)).replace(day=25), first.replace(day=24)


def raw_columns(header):
    normalized = [base.normalize(value) for value in header]
    columns = {}
    for field, aliases in RAW_ALIASES.items():
        matches = [index for index, name in enumerate(normalized) if name in aliases]
        if len(matches) > 1:
            raise ValueError(f"RAW Excel có nhiều cột cùng khớp {field}. Hãy giữ một cột cho trường này.")
        if matches:
            columns[field] = matches[0]
    return columns


def raw_stamp(value, epoch, separate_date=None):
    if separate_date is not None:
        if isinstance(separate_date, (int, float)):
            separate_date = from_excel(separate_date, epoch)
        day = form_date(separate_date)
        if isinstance(value, (int, float)):
            value = from_excel(value, epoch)
        moment = value.time() if isinstance(value, datetime) else hour(value)
        if day is None or moment is None:
            raise ValueError("Thiếu ngày hoặc giờ.")
        stamp = datetime.combine(day, moment)
    else:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            value = from_excel(value, epoch)
        if isinstance(value, datetime):
            stamp = value
        else:
            text = base.string(value)
            if len(text) <= 10:
                raise ValueError("Thiếu giờ.")
            try:
                stamp = datetime.fromisoformat(text.replace("Z", "+00:00"))
            except ValueError:
                stamp = None
                for pattern in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M:%S"):
                    try:
                        stamp = datetime.strptime(text, pattern)
                        break
                    except ValueError:
                        pass
                if stamp is None:
                    raise ValueError("Ngày giờ chưa hợp lệ.")
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=base.LOCAL_TZ)
    return stamp.isoformat()


def load_raw(path, employees, stats, issues):
    if path.suffix.lower() == ".csv":
        return base.load_clock(path, employees, stats, issues)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("File RAW cần là CSV hoặc Excel XLSX.")
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in book:
            rows = iter(sheet.values)
            for header_row in range(1, 21):
                header = next(rows, None)
                if header is None:
                    break
                columns = raw_columns(header)
                if "employeeNoString" not in columns:
                    continue
                if "time" not in columns and not {"date", "hour"} <= columns.keys():
                    continue
                with TemporaryDirectory(prefix="cham_cong_raw_") as directory:
                    temporary = Path(directory) / "raw.csv"
                    with temporary.open("w", encoding="utf-8-sig", newline="") as stream:
                        writer = csv.DictWriter(stream, fieldnames=RAW_FIELDS)
                        writer.writeheader()
                        for row in rows:
                            if not any(value is not None for value in row):
                                stream.write("\n")
                                continue
                            record = {field: row[column] if column < len(row) else None
                                      for field, column in columns.items()}
                            identity = base.employee_id(record.get("employeeNoString"))
                            try:
                                if "time" in columns:
                                    raw_time = record.get("time")
                                    only_hour = isinstance(raw_time, time) or (
                                        isinstance(raw_time, (int, float)) and 0 <= raw_time < 1) or bool(
                                            re.fullmatch(r"\d{1,2}:\d{2}(?::\d{2})?", base.string(raw_time)))
                                    value = raw_stamp(raw_time, book.epoch,
                                                      record.get("date") if only_hour and "date" in columns else None)
                                else:
                                    value = raw_stamp(record.get("hour"), book.epoch, record.get("date"))
                            except (ValueError, TypeError, OverflowError):
                                value = "Ngày giờ không hợp lệ"
                            writer.writerow({
                                **{field: record.get(field, "") for field in RAW_FIELDS},
                                "time": value, "employeeNoString": identity,
                            })
                    converted_issues = []
                    result = base.load_clock(temporary, employees, stats, converted_issues)
                    for issue in converted_issues:
                        match = re.match(r"CSV dòng (\d+)(.*)", issue)
                        issues.append(
                            f"RAW Excel trang {sheet.title}, dòng {int(match[1]) + header_row - 1}{match[2]}"
                            if match else issue)
                    return result
        raise ValueError(
            "RAW Excel cần có bảng mỗi dòng một lượt quét, với cột time và employeeNoString "
            "(hoặc Ngày, Giờ và Mã nhân viên).")
    finally:
        book.close()


def read_form(path):
    if path.suffix.lower() == ".xlsx":
        return base.read_sheet(path, FORM_REQUIRED, FORM_OPTIONAL)
    if path.suffix.lower() != ".csv":
        raise ValueError("File xin nghỉ cần là XLSX hoặc CSV.")
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        columns = {}
        for field, aliases in {**FORM_REQUIRED, **FORM_OPTIONAL}.items():
            for alias in aliases:
                matches = [name for name in reader.fieldnames or ()
                           if base.normalize(name) == alias or base.normalize(name).startswith(alias + " (")]
                if len(matches) > 1:
                    raise ValueError(f"File Form có nhiều cột cùng khớp {alias}.")
                if matches:
                    columns[field] = matches[0]
                    break
        if not FORM_REQUIRED.keys() <= columns.keys():
            raise ValueError("File Form thiếu họ tên, ngày bắt đầu hoặc loại thông báo.")
        return [(number, {field: row.get(column) for field, column in columns.items()})
                for number, row in enumerate(reader, 2) if any(row.values())]


def hour(value):
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    return time.fromisoformat(base.string(value)) if value else None


def form_date(value):
    if not isinstance(value, str):
        return base.form_date(value)
    text = value.strip()
    if not text:
        return None
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        for pattern in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
            try:
                return datetime.strptime(text, pattern).date()
            except ValueError:
                pass
    raise ValueError("Ngày chưa đúng dạng ngày/tháng/năm hoặc YYYY-MM-DD.")


def scope_of(row, rules):
    period, text = base.normalize(row.get("period")), base.normalize(row.get("detail"))
    first, last = hour(row.get("from_hour")), hour(row.get("to_hour"))
    if first or last:
        if not first or not last or first >= last:
            raise ValueError("Cần đủ Từ giờ và Đến giờ, theo thứ tự trong cùng ngày.")
        return "gio", first, last
    if period == "ca ngay":
        return "ca_ngay", None, None
    matches = list(CLOCK.finditer(base.string(row.get("detail"))))
    if len(matches) == 2:
        separator = base.string(row.get("detail"))[matches[0].end():matches[1].start()]
        if re.search(r"đến|tới|[-–]", separator, re.I):
            first, last = [time(int(m[1]), int(m[2] or 0)) for m in matches]
            if first >= last:
                raise ValueError("Khoảng giờ vắng chưa hợp lệ.")
            return "gio", first, last
    relative = re.search(r"(\d+)\s*(?:p\b|phut\b|tieng\b)", text)
    if relative and ("muon" in text or "som" in text):
        if "muon" in text and "som" in text:
            parts = re.findall(r"\b(muon|som)\b[^\d]*?(\d+)\s*(p\b|phut\b|tieng\b)", text)
            minutes_by_type = {kind: int(amount) * (60 if unit == "tieng" else 1)
                               for kind, amount, unit in parts}
            if set(minutes_by_type) != {"muon", "som"} or any(
                    value <= 0 or value >= 9 * 60 for value in minutes_by_type.values()):
                raise ValueError("Chưa rõ hai khoảng xin muộn và về sớm. Điền thành hai dòng Từ giờ/Đến giờ.")
            latest = (datetime.combine(date.min, rules["shift_start"])
                      + timedelta(minutes=minutes_by_type["muon"])).time()
            earliest = (datetime.combine(date.min, rules["shift_end"])
                        - timedelta(minutes=minutes_by_type["som"])).time()
            return "gio_kep", latest, earliest
        minutes = int(relative[1]) * (60 if "tieng" in relative[0] else 1)
        if minutes <= 0 or minutes >= 9 * 60:
            raise ValueError("Thời lượng xin muộn/về sớm chưa hợp lệ.")
        anchor = rules["shift_start"] if "muon" in text else rules["shift_end"]
        stamp = datetime.combine(date.min, anchor)
        other = (stamp + timedelta(minutes=minutes if "muon" in text else -minutes)).time()
        return ("gio", anchor, other) if "muon" in text else ("gio", other, anchor)
    if len(matches) == 1 and re.search(r"(xin\s+)?\d{1,2}\s*(?:h|gio).*\bden\b", text):
        deadline = time(int(matches[0][1]), int(matches[0][2] or 0))
        if deadline > rules["shift_start"]:
            return "gio", rules["shift_start"], deadline
    if period in ("sang", "chieu"):
        return period, None, None
    if matches:
        raise ValueError("Chưa xác định đủ khoảng vắng. Điền buổi hoặc Từ giờ/Đến giờ.")
    if re.search(r"\bnghi\s+(?:buoi\s+)?sang\b", text):
        return "sang", None, None
    if re.search(r"\bnghi\s+(?:buoi\s+)?chieu\b", text):
        return "chieu", None, None
    if "ca ngay" in text or base.RANGE_TEXT.search(base.string(row.get("detail"))) or COMPACT.search(text):
        return "ca_ngay", None, None
    raise ValueError("Chưa rõ cả ngày, buổi sáng, buổi chiều hoặc khoảng giờ.")


def match_identity(row, employees, names, aliases):
    identity = base.employee_id(row.get("id"))
    if identity:
        if identity not in employees:
            raise ValueError("Mã nhân viên trên Form chưa có trong danh sách.")
        return identity
    key = base.normalize(row["name"])
    if key in aliases:
        return aliases[key]
    candidates = names.get(key, [])
    if len(candidates) != 1:
        raise ValueError("Họ tên chưa khớp duy nhất. Bổ sung mã nhân viên hoặc tên thay thế trong cấu hình.")
    return candidates[0]


def load_notices(path, employees, rules, start, end):
    names = defaultdict(list)
    for employee in employees.values():
        names[base.normalize(employee.name)].append(employee.id)
    notices, issues, used, seen = defaultdict(list), [], [], set()
    corrections = rules["date_corrections"].get(hashlib.sha256(path.read_bytes()).hexdigest(), {})
    for number, row in read_form(path):
        known_date = None
        correction = corrections.get(str(number), {})
        row = {**row, **correction}
        try:
            first = form_date(row["start"])
            submitted = form_date(row.get("submitted"))
            known_date = first
            if first is None:
                raise ValueError("Thiếu ngày bắt đầu.")
            if not correction and submitted and first.day <= 12 and first.day != first.month:
                swapped = date(first.year, first.day, first.month)
                if (abs((first - submitted).days) > 30 and abs((swapped - submitted).days) <= 3
                        and any(start <= day <= end for day in (first, swapped))):
                    raise ValueError("Ngày có thể bị đảo ngày/tháng.")
            row = {**row, "end": form_date(row.get("end")), "other_end": form_date(row.get("other_end"))}
            last = legacy.notification_dates(row, first)
            if last < start or first > end:
                continue
            identity = match_identity(row, employees, names, rules["name_aliases"])
            if base.normalize(row["kind"]) not in KINDS:
                raise ValueError("Loại thông báo chưa nhận diện.")
            approval = base.normalize(row.get("approval"))
            if approval in ("tu choi", "khong duyet", "huy", "rejected", "cancelled"):
                continue
            if approval and approval not in ("da duyet", "dong y", "chap nhan", "approved"):
                raise ValueError("Trạng thái duyệt chưa xác nhận.")
            scope, begin, finish = scope_of(row, rules)
            scopes = [("gio", rules["shift_start"], begin), ("gio", finish, rules["shift_end"])] if scope == "gio_kep" else [(scope, begin, finish)]
            for day in legacy.date_range(max(start, first), min(end, last)):
                for scope, begin, finish in scopes:
                    notice = Notice(number, base.string(row["kind"]), base.string(row.get("period")),
                                    base.string(row.get("detail")), scope, begin, finish)
                    key = (identity, day, notice.kind, scope, begin, finish, notice.detail)
                    used.append([identity, employees[identity].name, day, number, notice.kind,
                                 notice.period, notice.detail, scope, begin, finish, key in seen])
                    if key not in seen:
                        notices[(identity, day)].append(notice)
                        seen.add(key)
        except (ValueError, TypeError) as error:
            # Giữ các lỗi chưa xác định ngày để người dùng có thể sửa và chạy lại.
            relevant = known_date is not None and start <= known_date <= end
            try:
                submitted = form_date(row.get("submitted"))
                relevant = relevant or submitted is not None and start <= submitted <= end
                relevant = relevant or known_date is None and submitted is None
            except ValueError:
                relevant = True
            if relevant:
                issues.append([number, base.string(row["name"]), base.string(row.get("start")),
                               base.string(row.get("kind")), base.string(row.get("period")),
                               base.string(row.get("detail")), str(error)])
    return notices, issues, used


def attendance_limits(notices, rules):
    morning_excused = any(n.scope == "sang" for n in notices)
    afternoon_excused = any(n.scope == "chieu" for n in notices)
    latest, earliest = rules["latest_in"], rules["earliest_out"]
    for notice in notices:
        if notice.scope == "gio":
            morning_excused = morning_excused or (
                notice.begin <= rules["shift_start"] and notice.finish >= rules["morning_cutoff"])
            afternoon_excused = afternoon_excused or (
                notice.begin <= rules["morning_cutoff"] and notice.finish >= rules["shift_end"])
            if notice.begin <= rules["shift_start"] < notice.finish:
                latest = max(latest, notice.finish)
            if notice.begin < rules["shift_end"] <= notice.finish:
                earliest = min(earliest, notice.begin)
    return morning_excused, afternoon_excused, latest, earliest


def day_metrics(times, notices, rules):
    incoming = min((t for t in times if t < rules["morning_cutoff"]), default=None)
    outgoing = max((t for t in times if t >= rules["morning_cutoff"]), default=None)
    paired = incoming is not None and outgoing is not None
    punctual = bool(paired and incoming <= rules["latest_in"] and outgoing >= rules["earliest_out"])
    full = any(n.scope == "ca_ngay" for n in notices)
    morning_excused, afternoon_excused, latest, earliest = attendance_limits(notices, rules)
    missing = not full and ((incoming is None and not morning_excused)
                            or (outgoing is None and not afternoon_excused))
    remaining_ok = ((morning_excused or incoming is not None and incoming <= latest)
                    and (afternoon_excused or outgoing is not None and outgoing >= earliest))
    credited = bool(full or punctual or rules["credit_partial_notices"] and notices and remaining_ok)
    if full:
        status = "Vắng có lý do cả ngày"
    elif missing:
        status = "Thiếu lượt điểm danh cần thiết"
    elif credited:
        status = "Thực hiện đúng" if punctual else "Thực hiện đúng sau áp dụng thông báo"
    else:
        status = "Chưa đạt giờ làm việc còn lại"
    return incoming, outgoing, int(paired), int(full), int(punctual), int(credited), int(missing), status


def compute(employees, punches, notices, rules, start, end):
    results, detail = {}, []
    for identity, employee in employees.items():
        total = Counter()
        for day in legacy.date_range(start, end):
            workday = day.weekday() in rules["workdays"] and day not in rules["holidays"]
            events = sorted(punches.get((identity, day), ()))
            entries = notices.get((identity, day), ())
            values = day_metrics(events, entries, rules)
            incoming, outgoing, paired, full, good, credited, missing, status = values
            if workday:
                total.update(D=1, E=paired, F=full, G=good, H=credited, I=missing)
            else:
                status = "Ngày nghỉ lễ theo lịch đơn vị" if day in rules["holidays"] else "Nghỉ cuối tuần"
            detail.append([
                identity, employee.name, employee.department, day, int(workday), incoming, outgoing,
                paired if workday else 0, full if workday else 0, good if workday else 0,
                credited if workday else 0, missing if workday else 0, status,
                ", ".join(str(n.row) for n in entries),
                " | ".join(f"{n.kind} / {n.period} / {n.detail}" for n in entries),
                "; ".join(t.isoformat(timespec="seconds") for t in events),
            ])
        total["J"] = total["D"] - total["H"]
        results[identity] = [total[key] for key in "DEFGHIJ"]
    return results, detail


def detail_output_path(summary):
    name = summary.name
    if name.startswith("TongHop_"):
        name = "ChiTiet_" + name[len("TongHop_"):]
    else:
        name = summary.stem + "_ChiTiet.xlsx"
    return summary.with_name(name)


def detailed_day(row, notices, rules):
    incoming, outgoing = row[5], row[6]
    if not row[4]:
        return incoming or "-", outgoing or "-", 0, 0, 0, "Nghỉ lễ" if row[3] in rules["holidays"] else "#"
    morning_excused, afternoon_excused, latest, earliest = attendance_limits(notices, rules)
    late_seconds = base.minutes_between(latest, incoming) * 60 if incoming and not morning_excused and not row[8] else 0
    early_seconds = base.minutes_between(outgoing, earliest) * 60 if outgoing and not afternoon_excused and not row[8] else 0
    late = max(0, int((late_seconds + 59) // 60))
    early = max(0, int((early_seconds + 59) // 60))
    worked = int(base.minutes_between(incoming, outgoing) + 0.5) if incoming and outgoing else 0
    kind_codes = {
        "nghi phep": "P", "di cong tac": "CT", "di hoc": "Đh",
        "du hop": "Họp", "du hoi thao": "DHt", "khac": "Ld",
    }
    labels = []
    for notice in notices:
        code = kind_codes[base.normalize(notice.kind)]
        suffix = {"sang": "(S)", "chieu": "(C)", "gio": "(G)"}.get(notice.scope, "")
        if code + suffix not in labels:
            labels.append(code + suffix)
    if row[8]:
        status = "/".join(labels)
    else:
        status = "A" if row[11] else "W"
        status = ("L-" if late else "") + status + ("-E" if early else "")
        if labels:
            status += "/" + "/".join(labels)
    return incoming or "-", outgoing or "-", late, early, worked, status


def export_detail(out, raw, form, start, end, employees, detail, notices, issues, rules):
    book = Workbook()
    book.remove(book.active)
    lookup = {(row[0], row[3]): row for row in detail}
    all_days = legacy.date_range(start, end)
    for page, offset in enumerate(range(0, len(all_days), 31), 1):
        days = all_days[offset:offset + 31]
        sheet = book.create_sheet("Sheet1" if page == 1 else f"Ky_{page}")
        sheet.sheet_view.showGridLines = False
        sheet.column_dimensions["A"].width = 18
        for column in range(2, 33):
            sheet.column_dimensions[get_column_letter(column)].width = 11
        base.write_row(sheet, 3, ["Báo cáo chuyên cần hàng tháng"], [32], bold=True, center=True)
        base.write_row(sheet, 4, [f"Từ: {days[0]} 00:00:00 Đến: {days[-1]} 23:59:59"], [32], bold=True, center=True)
        base.write_row(sheet, 5, [
            "P = nghỉ phép, CT = công tác, Đh = đi học, Họp = dự họp, DHt = hội thảo, Ld = lý do khác. "
            "(S) = sáng, (C) = chiều, (G) = theo giờ. # = nghỉ cuối tuần. Nghỉ lễ = ngày nghỉ theo lịch đơn vị. A = thiếu lượt cần thiết. "
            f"Có {len(issues)} mục cần kiểm tra nguồn."
        ], [32], wrap=True)
        sheet.row_dimensions[5].height = 30
        header = 6
        for identity, employee in employees.items():
            info = (employee.id, employee.name, employee.department, "-", "-")
            values = [value for pair in zip(base.HEADER_LABELS, info) for value in pair]
            base.write_row(sheet, header, values, base.HEADER_SPANS, bold=True, yellow=True, wrap=True)
            sheet.row_dimensions[header].height = 42
            records = [lookup[(identity, day)] for day in days]
            computed = [detailed_day(row, notices.get((identity, row[3]), ()), rules) for row in records]
            padding = 31 - len(days)
            rows = [
                ["Ngày:", *[day.day for day in days], *(["-"] * padding)],
                ["Đăng ký1", *[r[0] for r in computed], *(["-"] * padding)],
                ["Ra1", *[r[1] for r in computed], *(["-"] * padding)],
                ["Tăng ca", *([0] * 31)],
                ["Muộn", *[r[2] for r in computed], *([0] * padding)],
                ["Nghỉ sớm", *[r[3] for r in computed], *([0] * padding)],
                ["Đã tham dự", *[r[4] for r in computed], *([0] * padding)],
                ["Giờ Nghỉ", *([0] * 31)],
                ["Trạng thái", *[r[5] for r in computed], *(["-"] * padding)],
            ]
            for row_offset, values in enumerate(rows, 1):
                base.write_row(sheet, header + row_offset, values, (1,) * 32, wrap=row_offset == 9)
                sheet.row_dimensions[header + row_offset].height = 36 if row_offset == 9 else 18
            for index, row in enumerate(records, 2):
                if not row[4]:
                    for row_offset in range(1, 10):
                        sheet.cell(header + row_offset, index).fill = PatternFill("solid", fgColor="FFE9EDF2")
            totals = [
                sum(row[4] for row in records), sum(row[7] for row in records),
                sum(row[8] for row in records), sum(row[9] for row in records),
                sum(row[10] for row in records), sum(row[11] for row in records),
            ]
            totals.append(totals[0] - totals[4])
            names = ("Ngày làm việc", "Đủ dữ liệu", "Vắng cả ngày", "Đúng giờ thực tế",
                     "Đúng sau thông báo", "Thiếu lượt", "Chưa đạt")
            summary = ". ".join(f"{label} {value}" for label, value in zip(names, totals))
            base.write_row(sheet, header + 10, ["Tóm tắt", summary], [1, 31], wrap=True)
            sheet.row_dimensions[header + 10].height = 27
            header += 11
        note = (
            "Giờ vào/ra lấy từ RAW và giữ trong hai dòng giờ. Thông báo Form được ghi tại dòng Trạng thái. "
            "Muộn/về sớm là số phút vượt mốc cho phép sau áp dụng thông báo, làm tròn lên phút. "
            "Đã tham dự là chênh lệch phút giữa hai giờ đã chọn. Tăng ca và Giờ Nghỉ để 0 theo mẫu. "
            "Nghỉ một buổi hoặc theo giờ tính 0 ngày nghỉ."
        )
        base.write_row(sheet, header, [note], [32], wrap=True)
        sheet.row_dimensions[header].height = 40
        sheet.freeze_panes = "B8"
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.print_area = f"A3:AF{header}"
    legacy.add_table(book, "CanKiemTra", [
        "Dòng Form/CSV", "Họ tên hoặc mã", "Ngày khai", "Loại", "Buổi", "Nội dung", "Vấn đề",
    ], issues, [16, 30, 24, 22, 18, 80, 80])
    legacy.add_table(book, "ThongBaoForm", [
        "Mã nhân viên", "Họ tên", "Ngày", "Giờ vào RAW", "Giờ ra RAW", "Dòng Form", "Thông báo",
    ], [[row[0], row[1], row[3], row[5], row[6], row[13], row[14]] for row in detail if row[13]],
                     [14, 30, 14, 15, 15, 20, 100])
    legacy.add_table(book, "Nguon", ["Nội dung", "Giá trị"], [
        ["RAW", str(raw.resolve())], ["Form", str(form.resolve())],
        ["Kỳ", f"{start} đến {end}"],
        ["Ngày nghỉ lễ trong kỳ", ", ".join(f"{day:%d/%m/%Y}" for day in sorted(rules["holidays"]) if start <= day <= end) or "Không có"],
        ["Quy tắc", "Cùng dữ liệu và kết quả ngày với báo cáo tổng hợp xuất kèm."],
    ], [35, 120])
    with base.staged_file(out) as temporary:
        book.save(temporary)
    book.close()


def export_report(out, raw, form, start, end, employees, results, detail, issues, used, stats, rules):
    book = load_workbook(rules["template_path"])
    sheet = book["BaoCaoChamCong"]
    styles = [copy(sheet.cell(5, col)._style) for col in range(1, 12)]
    for merged in list(sheet.merged_cells.ranges):
        if merged.min_row >= 5:
            sheet.unmerge_cells(str(merged))
    sheet.delete_rows(5, max(0, sheet.max_row - 4))
    sheet["B2"] = f"BÁO CÁO TỔNG HỢP CHẤM CÔNG TỪ {start:%d/%m/%Y} ĐẾN {end:%d/%m/%Y}"
    sheet["A3"] = (
        "Ngày nghỉ lễ trong kỳ: " + (", ".join(f"{day:%d/%m/%Y}" for day in sorted(rules["holidays"]) if start <= day <= end) or "Không có") + ". Không tính vào D–J. "
        f"Chọn lượt sáng sớm nhất trước {rules['morning_cutoff']} và lượt chiều muộn nhất. "
        f"Đúng giờ thực tế khi vào ≤{rules['latest_in']}, ra ≥{rules['earliest_out']}. "
        "F chỉ tính vắng cả ngày. Vắng một buổi/theo giờ tính 0 ở F. "
        "H đếm mỗi ngày một lần khi đúng giờ, vắng cả ngày hoặc đạt phần giờ còn lại sau thông báo. "
        "I chỉ tính thiếu lượt cần thiết sau áp dụng thông báo. J = D − H. "
        f"Có {len(issues)} mục cần kiểm tra nguồn. Kết quả cần chốt sau khi xử lý các mục này."
    )
    sheet["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.row_dimensions[3].height = 95
    sheet["F4"] = "Số ngày vắng có lý do cả ngày"
    sheet["H4"] = "Số ngày thực hiện đúng (gồm cả ngày và phần ngày có thông báo hợp lệ)"
    sheet["J4"] = "Số ngày chưa đạt hoặc chưa đủ căn cứ xác nhận"
    groups = []
    personal_issues = Counter()
    names = {base.normalize(employee.name): identity for identity, employee in employees.items()}
    for issue in issues:
        identity = names.get(base.normalize(issue[1]), rules["name_aliases"].get(base.normalize(issue[1])))
        if identity:
            personal_issues[identity] += 1
    for index, (identity, employee) in enumerate(employees.items(), 5):
        values = [index - 4, employee.department, employee.name, *results[identity],
                  f"Tạm tính, có {personal_issues[identity]} mục Form cần kiểm tra"
                  if personal_issues[identity] else ""]
        for column, value in enumerate(values, 1):
            cell = sheet.cell(index, column, value)
            cell._style = copy(styles[column - 1])
            if isinstance(value, str):
                cell.data_type = "s"
            cell.alignment = Alignment(vertical="center", horizontal="left" if column in (2, 3, 11) else "center", wrap_text=True)
        sheet.row_dimensions[index].height = 34
        if not groups or groups[-1][0] != employee.department:
            groups.append([employee.department, index, index])
        else:
            groups[-1][2] = index
    for department, first, last in groups:
        if last > first:
            sheet.merge_cells(start_row=first, start_column=2, end_row=last, end_column=2)
    sheet.freeze_panes = "D5"
    sheet.sheet_view.showGridLines = False
    sheet.print_area = f"A1:K{4 + len(employees)}"
    legacy.add_table(book, "ChiTietNgay", [
        "Mã nhân viên", "Họ tên", "Phòng", "Ngày", "Ngày làm việc", "Giờ vào RAW", "Giờ ra RAW",
        "Đủ hai giờ", "Vắng cả ngày", "Đúng giờ thực tế", "Thực hiện đúng sau thông báo",
        "Thiếu lượt cần thiết", "Trạng thái", "Dòng Form", "Thông báo", "Các lượt quẹt RAW",
    ], detail, [14, 30, 35, 14, 15, 15, 15, 15, 15, 18, 22, 20, 45, 20, 80, 75])
    scope_names = {"ca_ngay": "Cả ngày", "sang": "Buổi sáng", "chieu": "Buổi chiều", "gio": "Theo giờ"}
    form_rows = [[*row[:7], scope_names[row[7]], *row[8:]] for row in used]
    legacy.add_table(book, "FormDaGhep", [
        "Mã nhân viên", "Họ tên", "Ngày", "Dòng Form", "Loại", "Buổi", "Nội dung",
        "Phạm vi áp dụng", "Từ giờ", "Đến giờ", "Thông báo trùng đã gộp",
    ], form_rows, [14, 30, 14, 12, 22, 18, 80, 20, 15, 15, 22])
    legacy.add_table(book, "CanKiemTra", [
        "Dòng Form/CSV", "Họ tên hoặc mã", "Ngày khai", "Loại", "Buổi", "Nội dung", "Vấn đề",
    ], issues, [16, 30, 24, 22, 18, 80, 80])
    info = [
        ["RAW", str(raw.resolve())], ["SHA256 RAW", hashlib.sha256(raw.read_bytes()).hexdigest()],
        ["Form", str(form.resolve())], ["SHA256 Form", hashlib.sha256(form.read_bytes()).hexdigest()],
        ["Kỳ", f"{start} đến {end}"], ["Thông báo một buổi", "Nghỉ sáng kiểm tra giờ ra còn lại. Nghỉ chiều kiểm tra giờ vào còn lại. F tính 0."],
        ["Thông báo theo giờ", "Chỉ nới mốc vào/ra khi khoảng giờ thông báo bao phủ đầu/cuối ca. Giờ bắt đầu/kết thúc không rõ được đưa vào CanKiemTra."],
        ["Ngày vắng", "Chỉ thông báo cả ngày tính 1. Hai thông báo sáng và chiều không tự cộng thành 1 ở F."],
        ["Duyệt Form", "Nếu có cột Trạng thái duyệt, chỉ áp dụng mục đã duyệt. Nếu không có cột này, áp dụng phản hồi ghép được và ghi rõ nguồn."],
        ["Cấu hình", json.dumps({k: str(v) for k, v in rules.items() if k not in ("date_corrections",)}, ensure_ascii=False)],
        ["Điều chỉnh Form đã xác nhận", json.dumps(rules.get("date_corrections", {}).get(
            hashlib.sha256(form.read_bytes()).hexdigest(), {}), ensure_ascii=False)],
        *[[key, value] for key, value in stats.items()],
    ]
    legacy.add_table(book, "NguonQuyTac", ["Nội dung", "Giá trị"], info, [40, 120])
    if out.resolve() in {raw.resolve(), form.resolve(), rules["template_path"].resolve()}:
        raise ValueError("Đường dẫn kết quả trùng file nguồn.")
    if out.suffix.lower() != ".xlsx":
        raise ValueError("File kết quả cần có đuôi .xlsx.")
    out.parent.mkdir(parents=True, exist_ok=True)
    with base.staged_file(out) as temporary:
        book.save(temporary)
    book.close()


def generate(raw, form, start, end, out, config=CONFIG, holidays=None):
    if start > end:
        raise ValueError("Ngày bắt đầu phải trước hoặc bằng ngày kết thúc.")
    rules, employees = load_setup(config)
    if holidays is not None:
        rules["holidays"] = set(holidays)
        if any(type(day) is not date or not start <= day <= end for day in rules["holidays"]):
            raise ValueError("Ngày nghỉ lễ phải là ngày hợp lệ nằm trong kỳ báo cáo.")
    stats, clock_issues = Counter(), []
    dates, punches, unknown = load_raw(raw, employees, stats, clock_issues)
    if not any(start <= day <= end for day in dates):
        raise ValueError("RAW không có sự kiện nào trong kỳ đã chọn.")
    notices, issues, used = load_notices(form, employees, rules, start, end)
    issues.extend([["", "", "", "", "", "", text] for text in clock_issues])
    for (day, identity), count in unknown.items():
        if start <= day <= end:
            issues.append(["", identity, str(day), "", "", "", f"{count} sự kiện có mã ngoài danh sách nhân viên."])
    ranges = re.findall(r"\d{4}-\d{2}-\d{2}", raw.name)
    if len(ranges) >= 2 and (date.fromisoformat(ranges[0]) > start or date.fromisoformat(ranges[1]) < end):
        issues.append(["", "", "", "", "", raw.name, "Kỳ trong tên RAW chưa bao phủ kỳ báo cáo. Cần bổ sung RAW đủ kỳ."])
    results, detail = compute(employees, punches, notices, rules, start, end)
    for day in legacy.date_range(start, end):
        if (day.weekday() in rules["workdays"] and day not in rules["holidays"]
                and not any(punches.get((identity, day)) for identity in employees)):
            issues.append(["", "Toàn đơn vị", str(day), "", "", "", "Không có lượt hợp lệ toàn đơn vị. Kiểm tra ngày nghỉ hoặc dữ liệu nguồn."])
    stats["Nhân viên"] = len(employees)
    stats["Ngày làm việc mỗi người"] = next(iter(results.values()))[0]
    stats["Phản hồi đã ghép"] = len({r[3] for r in used})
    companion = detail_output_path(out)
    sources = {raw.resolve(), form.resolve(), rules["template_path"].resolve(), config.resolve()}
    if out.resolve() in sources or companion.resolve() in sources:
        raise ValueError("Đường dẫn kết quả trùng file nguồn hoặc cấu hình.")
    if out.suffix.lower() != ".xlsx":
        raise ValueError("File kết quả cần có đuôi .xlsx.")
    out.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="cham_cong_xuat_", dir=out.parent) as directory:
        stage = Path(directory)
        summary_stage, detail_stage = stage / out.name, stage / companion.name
        export_report(summary_stage, raw, form, start, end, employees, results, detail, issues, used, stats, rules)
        export_detail(detail_stage, raw, form, start, end, employees, detail, notices, issues, rules)
        committed = []
        backups = {}
        for target in (out, companion):
            if target.exists():
                backup = stage / (target.name + ".backup")
                shutil.copy2(target, backup)
                backups[target] = backup
        try:
            for staged, target in ((summary_stage, out), (detail_stage, companion)):
                os.replace(staged, target)
                committed.append(target)
        except OSError:
            for target in reversed(committed):
                if target in backups:
                    os.replace(backups[target], target)
                else:
                    target.unlink(missing_ok=True)
            raise
    return results, issues


def suggest_month(mode):
    today = datetime.now(base.LOCAL_TZ).date()
    current = today.replace(day=1)
    if mode == "calendar" or today.day < 24:
        current = (current - timedelta(days=1)).replace(day=1)
    return current.strftime("%Y-%m")


def choose_holidays(parent, start, end, selected, on_apply):
    """Chọn nhiều ngày trực tiếp trên lịch, chỉ thay đổi khi bấm Áp dụng."""
    import tkinter as tk
    from tkinter import ttk

    dialog = tk.Toplevel(parent)
    dialog.title("Chọn ngày nghỉ lễ")
    dialog.transient(parent)
    dialog.resizable(False, False)
    frame = ttk.Frame(dialog, padding=16)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text=f"Kỳ báo cáo: {start:%d/%m/%Y} – {end:%d/%m/%Y}").pack(anchor="w")
    ttk.Label(frame, text="Bấm ngày để chọn hoặc bỏ chọn. Ngày được chọn có dấu ✓.").pack(anchor="w", pady=(4, 12))
    pending = {day for day in selected if start <= day <= end}
    month = [start.replace(day=1)]
    navigation = ttk.Frame(frame)
    navigation.pack(fill="x")
    title = tk.StringVar()
    grid = ttk.Frame(frame)
    grid.pack(pady=10)
    summary = tk.StringVar()
    ttk.Label(frame, textvariable=summary, wraplength=450).pack(anchor="w", pady=(0, 10))

    def shift(amount):
        year, zero_month = divmod(month[0].year * 12 + month[0].month - 1 + amount, 12)
        month[0] = date(year, zero_month + 1, 1)
        draw()

    previous = ttk.Button(navigation, text="◀", width=4, command=lambda: shift(-1))
    previous.pack(side="left")
    ttk.Label(navigation, textvariable=title, anchor="center").pack(side="left", expand=True)
    following = ttk.Button(navigation, text="▶", width=4, command=lambda: shift(1))
    following.pack(side="right")

    def toggle(day):
        if day in pending:
            pending.remove(day)
        else:
            pending.add(day)
        draw()

    def draw():
        for widget in grid.winfo_children():
            widget.destroy()
        title.set(f"Tháng {month[0].month:02d}/{month[0].year}")
        previous.configure(state="normal" if month[0] > start.replace(day=1) else "disabled")
        following.configure(state="normal" if month[0] < end.replace(day=1) else "disabled")
        for column, label in enumerate(("T2", "T3", "T4", "T5", "T6", "T7", "CN")):
            ttk.Label(grid, text=label, anchor="center", width=7).grid(row=0, column=column, pady=3)
        for row, week in enumerate(calendar.monthcalendar(month[0].year, month[0].month), 1):
            for column, number in enumerate(week):
                if not number:
                    continue
                day = month[0].replace(day=number)
                button = tk.Button(grid, text=f"✓ {number}" if day in pending else str(number),
                                   width=6, relief="sunken" if day in pending else "raised",
                                   background="#DCE6F1" if day in pending else "SystemButtonFace",
                                   command=lambda value=day: toggle(value))
                button.grid(row=row, column=column, padx=2, pady=2)
                if not start <= day <= end:
                    button.configure(state="disabled")
        summary.set("Đã chọn: " + (", ".join(f"{day:%d/%m/%Y}" for day in sorted(pending)) or "Không có ngày nghỉ lễ"))

    def apply():
        on_apply(set(pending))
        dialog.destroy()

    def clear():
        pending.clear()
        draw()

    actions = ttk.Frame(frame)
    actions.pack(fill="x")
    ttk.Button(actions, text="Bỏ chọn tất cả", command=clear).pack(side="left")
    ttk.Button(actions, text="Hủy", command=dialog.destroy).pack(side="right")
    ttk.Button(actions, text="Áp dụng", command=apply).pack(side="right", padx=8)
    draw()
    dialog.grab_set()
    dialog.focus_set()
    return dialog


def launch_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    rules, _ = load_setup(CONFIG)
    root = tk.Tk()
    root.title("Tổng hợp chấm công từ RAW và Form")
    root.geometry("900x540")
    root.minsize(820, 520)
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    raw, form = tk.StringVar(), tk.StringVar()
    mode = tk.StringVar(value=rules["period_mode"])
    month = tk.StringVar(value=suggest_month(mode.get()))
    first, last = tk.StringVar(), tk.StringVar()
    destination = tk.StringVar(value=str(ROOT / "ket_qua"))
    status = tk.StringVar(value="Chọn RAW và file xin nghỉ, kiểm tra kỳ rồi bấm Tạo báo cáo.")
    output = {"path": None}
    messages = queue.Queue()
    selected_holidays = set()
    holiday_text = tk.StringVar(value="Không có ngày nghỉ lễ")

    def set_holidays(days):
        selected_holidays.clear()
        selected_holidays.update(days)
        holiday_text.set(", ".join(f"{day:%d/%m/%Y}" for day in sorted(days)) or "Không có ngày nghỉ lễ")

    def reset_holidays(*_):
        try:
            begin, finish = date.fromisoformat(first.get()), date.fromisoformat(last.get())
        except ValueError:
            return
        set_holidays({day for day in rules["holidays"] if begin <= day <= finish})

    first.trace_add("write", reset_holidays)
    last.trace_add("write", reset_holidays)

    def browse(variable, filetypes):
        value = filedialog.askopenfilename(filetypes=filetypes)
        if value:
            variable.set(value)

    for row, (label, variable, types) in enumerate([
        ("File RAW", raw, [("RAW Excel / CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv")]),
        ("File xin nghỉ", form, [("Excel / CSV", "*.xlsx *.csv")]),
    ]):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=10)
        ttk.Button(frame, text="Chọn file", command=lambda v=variable, t=types: browse(v, t)).grid(row=row, column=2)
    ttk.Label(frame, text="Tháng báo cáo").grid(row=2, column=0, sticky="w", pady=8)
    ttk.Entry(frame, textvariable=month, width=12).grid(row=2, column=1, sticky="w", padx=10)
    period_box = ttk.Combobox(frame, textvariable=mode, values=["25-24", "calendar"], state="readonly", width=14)
    period_box.grid(row=2, column=2)

    def update_period(*_):
        try:
            begin, finish = period_dates(month.get(), mode.get())
            first.set(begin.isoformat())
            last.set(finish.isoformat())
        except ValueError:
            pass

    mode.trace_add("write", update_period)
    month.trace_add("write", update_period)
    update_period()
    ttk.Label(frame, text="Từ ngày / đến ngày").grid(row=3, column=0, sticky="w", pady=8)
    dates_frame = ttk.Frame(frame)
    dates_frame.grid(row=3, column=1, sticky="w", padx=10)
    ttk.Entry(dates_frame, textvariable=first, width=15).pack(side="left")
    ttk.Label(dates_frame, text="  đến  ").pack(side="left")
    ttk.Entry(dates_frame, textvariable=last, width=15).pack(side="left")
    ttk.Label(frame, text="Ngày dạng YYYY-MM-DD. calendar là tháng dương lịch.").grid(row=4, column=1, sticky="w", padx=10)
    ttk.Label(frame, text="Ngày nghỉ lễ").grid(row=5, column=0, sticky="w", pady=8)
    ttk.Label(frame, textvariable=holiday_text, wraplength=560).grid(row=5, column=1, sticky="w", padx=10)

    def pick_holidays():
        try:
            begin, finish = date.fromisoformat(first.get()), date.fromisoformat(last.get())
            if begin > finish:
                raise ValueError("Ngày bắt đầu phải trước ngày kết thúc.")
        except ValueError as error:
            messagebox.showerror("Kiểm tra kỳ báo cáo", str(error))
            return
        choose_holidays(root, begin, finish, selected_holidays, set_holidays)

    ttk.Button(frame, text="Chọn ngày…", command=pick_holidays).grid(row=5, column=2)
    ttk.Label(frame, text="Ngày nghỉ lễ áp dụng cho lần xuất này; không tính ngày làm việc, thiếu lượt hoặc vắng.").grid(row=6, column=1, columnspan=2, sticky="w", padx=10)
    ttk.Label(frame, text="Thư mục kết quả").grid(row=7, column=0, sticky="w", pady=8)
    ttk.Entry(frame, textvariable=destination).grid(row=7, column=1, sticky="ew", padx=10)

    def choose_directory():
        value = filedialog.askdirectory()
        if value:
            destination.set(value)

    ttk.Button(frame, text="Chọn thư mục", command=choose_directory).grid(row=7, column=2)
    progress = ttk.Progressbar(frame, mode="indeterminate")
    progress.grid(row=9, column=0, columnspan=3, sticky="ew", pady=12)
    ttk.Label(frame, textvariable=status, wraplength=800).grid(row=10, column=0, columnspan=3, sticky="w")

    def complete():
        try:
            success, text = messages.get_nowait()
        except queue.Empty:
            root.after(150, complete)
            return
        progress.stop()
        create.configure(state="normal")
        status.set(text)
        (messagebox.showinfo if success else messagebox.showerror)("Tổng hợp chấm công", text)

    def run():
        try:
            begin, finish = date.fromisoformat(first.get()), date.fromisoformat(last.get())
            input_raw, input_form = Path(raw.get()), Path(form.get())
            if not raw.get() or not form.get() or not input_raw.is_file() or not input_form.is_file():
                raise ValueError("Hãy chọn đủ hai file đang tồn tại.")
            if begin > finish:
                raise ValueError("Ngày bắt đầu phải trước ngày kết thúc.")
            run_holidays = set(selected_holidays)
            target = Path(destination.get()) / f"TongHop_TuDong_{begin}_{finish}.xlsx"
            version = 2
            while target.exists() or detail_output_path(target).exists():
                target = Path(destination.get()) / f"TongHop_TuDong_{begin}_{finish}_v{version}.xlsx"
                version += 1
        except ValueError as error:
            messagebox.showerror("Kiểm tra đầu vào", str(error))
            return
        create.configure(state="disabled")
        progress.start()
        status.set("Đang ghép dữ liệu và tạo báo cáo…")

        def worker():
            try:
                _, issues = generate(input_raw, input_form, begin, finish, target, holidays=run_holidays)
                output["path"] = target
                messages.put((True, f"Đã tạo hai file:\n{target}\n{detail_output_path(target)}\n"
                                   f"Có {len(issues)} mục trong trang CanKiemTra."))
            except Exception as error:
                messages.put((False, str(error)))

        threading.Thread(target=worker, daemon=True).start()
        root.after(150, complete)

    create = ttk.Button(frame, text="Tạo báo cáo", command=run)
    create.grid(row=8, column=1, sticky="w", padx=10, pady=10)
    def open_output():
        folder = output["path"].parent if output["path"] else Path(destination.get())
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))

    ttk.Button(frame, text="Mở thư mục kết quả", command=open_output).grid(row=8, column=2)
    root.mainloop()


def main():
    if len(sys.argv) == 1:
        launch_gui()
        return
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument("--raw", type=Path, required=True)
    cli.add_argument("--form", type=Path, required=True)
    cli.add_argument("--start", type=date.fromisoformat, required=True)
    cli.add_argument("--end", type=date.fromisoformat, required=True)
    cli.add_argument("--out", type=Path)
    cli.add_argument("--config", type=Path, default=CONFIG)
    args = cli.parse_args()
    out = args.out or ROOT / "ket_qua" / f"TongHop_TuDong_{args.start}_{args.end}.xlsx"
    results, issues = generate(args.raw, args.form, args.start, args.end, out, args.config)
    print(f"Kết quả: {out.resolve()}")
    print(f"Chi tiết: {detail_output_path(out).resolve()}")
    print(f"{len(results)} nhân viên. {len(issues)} mục cần kiểm tra.")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError) as error:
        sys.exit(f"Lỗi: {error}")
