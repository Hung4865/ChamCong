"""Kiểm tra các quy tắc chấm công đã thống nhất."""

from collections import Counter
from datetime import date, datetime, time
from pathlib import Path
import csv
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

import monthly_report as report


class AttendanceRulesTest(unittest.TestCase):
    def setUp(self):
        self.rules, _ = report.load_setup(report.CONFIG)
        self.rules.update(
            shift_start=time(8), shift_end=time(17), latest_in=time(8, 3),
            earliest_out=time(17, 0, 35), morning_cutoff=time(13),
            credit_partial_notices=True, workdays=[0, 1, 2, 3, 4], holidays=set(),
        )

    def notice(self, scope, begin=None, finish=None):
        return report.Notice(2, "Nghỉ phép", "", "", scope, begin, finish)

    def test_early_face_scan_and_latest_exit(self):
        values = report.day_metrics(
            [time(6, 22, 4), time(8, 3, 6), time(16, 59), time(17, 3, 15)], [], self.rules)
        self.assertEqual(values[:7], (time(6, 22, 4), time(17, 3, 15), 1, 0, 1, 1, 0))

    def test_morning_leave_credits_remaining_afternoon_without_counting_leave(self):
        values = report.day_metrics([time(17, 10)], [self.notice("sang")], self.rules)
        self.assertEqual(values[2:7], (0, 0, 0, 1, 0))
        early = report.day_metrics([time(16, 55)], [self.notice("sang")], self.rules)
        self.assertEqual(early[5], 0)

    def test_afternoon_leave_credits_remaining_morning(self):
        values = report.day_metrics([time(7, 50)], [self.notice("chieu")], self.rules)
        self.assertEqual(values[2:7], (0, 0, 0, 1, 0))

    def test_hours_only_cover_the_requested_range(self):
        notice = self.notice("gio", time(8), time(9))
        good = report.day_metrics([time(8, 40), time(17, 10)], [notice], self.rules)
        late = report.day_metrics([time(9, 10), time(17, 10)], [notice], self.rules)
        self.assertEqual(good[3:7], (0, 0, 1, 0))
        self.assertEqual(late[5], 0)

    def test_full_leave_and_actual_attendance_are_counted_once(self):
        values = report.day_metrics([time(7, 50), time(17, 10)],
                                    [self.notice("ca_ngay"), self.notice("ca_ngay")], self.rules)
        self.assertEqual(values[3:7], (1, 1, 1, 0))

    def test_two_partial_notices_do_not_become_one_day_leave(self):
        values = report.day_metrics([], [self.notice("sang"), self.notice("chieu")], self.rules)
        self.assertEqual(values[3:7], (0, 0, 1, 0))

    def test_weekend_is_excluded_even_with_a_full_leave_notice(self):
        employee = report.base.Employee("1", "A", "Phòng")
        day = date(2026, 10, 3)
        results, detail = report.compute({"1": employee}, {}, {("1", day): [self.notice("ca_ngay")]},
                                         self.rules, day, day)
        self.assertEqual(results["1"], [0] * 7)
        self.assertEqual(detail[0][12], "Nghỉ cuối tuần")

    def test_detail_preserves_early_scan_partial_leave_and_excludes_weekend(self):
        employees = {"1": report.base.Employee("1", "Nhân viên A", "Phòng")}
        start, end = date(2026, 10, 1), date(2026, 10, 3)
        punches = {
            ("1", start): [time(6, 15), time(8, 10), time(17, 1)],
            ("1", date(2026, 10, 2)): [time(17, 2)],
        }
        notices = {("1", date(2026, 10, 2)): [self.notice("sang")]}
        results, detail = report.compute(employees, punches, notices, self.rules, start, end)
        self.assertEqual(results["1"], [2, 1, 0, 1, 2, 0, 0])
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "ChiTiet.xlsx"
            report.export_detail(target, Path("raw.csv"), Path("form.xlsx"), start, end,
                                 employees, detail, notices, [], self.rules)
            book = load_workbook(target)
            sheet = book["Sheet1"]
            self.assertEqual(sheet["B8"].value, time(6, 15))
            self.assertEqual(sheet["B9"].value, time(17, 1))
            self.assertEqual(sheet["B15"].value, "W")
            self.assertEqual(sheet["C15"].value, "W/P(S)")
            self.assertEqual(sheet["D15"].value, "#")
            self.assertIn("Vắng cả ngày 0", sheet["B16"].value)
            self.assertIn("Đúng sau thông báo 2", sheet["B16"].value)
            self.assertIn("Chưa đạt 0", sheet["B16"].value)
            self.assertEqual(sheet["C8"].value, "-")
            self.assertEqual(sheet["C9"].value, time(17, 2))
            self.assertEqual(sheet.max_column, 32)
            book.close()

    def test_year_boundary_and_leap_month(self):
        self.assertEqual(report.period_dates("2027-01", "25-24"),
                         (date(2026, 12, 25), date(2027, 1, 24)))
        self.assertEqual(report.period_dates("2028-02", "calendar"),
                         (date(2028, 2, 1), date(2028, 2, 29)))
        self.assertEqual(report.legacy.notification_dates(
            {"detail": "Nghỉ phép từ 30/12 đến 2/1"}, date(2026, 12, 30)), date(2027, 1, 2))

    def test_csv_form_duplicate_ranges_and_two_hour_requests(self):
        text = (
            "Họ tên,Vắng từ ngày,Loại thông báo,Thời gian,Nội dung\n"
            "A,2026-10-01,Nghỉ phép,Cả ngày,Nghỉ phép\n"
            "A,2026-10-01,Nghỉ phép,Cả ngày,Nghỉ phép\n"
            "A,2026-10-02,Khác,Khác,Xin đến muộn sáng 1 tiếng và chiều về sớm 1 tiếng\n"
        )
        employee = report.base.Employee("1", "A", "Phòng")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "form.csv"
            path.write_text(text, encoding="utf-8-sig")
            notices, issues, used = report.load_notices(
                path, {"1": employee}, self.rules, date(2026, 10, 1), date(2026, 10, 2))
        self.assertEqual(issues, [])
        self.assertEqual(len(notices["1", date(2026, 10, 1)]), 1)
        self.assertEqual(len(notices["1", date(2026, 10, 2)]), 2)
        good = report.day_metrics([time(8, 50), time(16, 10)],
                                 notices["1", date(2026, 10, 2)], self.rules)
        self.assertEqual(good[3:7], (0, 0, 1, 0))
        self.assertEqual(sum(bool(row[-1]) for row in used), 1)

    def test_short_name_is_not_silently_assigned(self):
        employee = report.base.Employee("1", "Nguyễn Ngọc Thành", "Phòng")
        with self.assertRaises(ValueError):
            report.match_identity({"name": "Thành"}, {"1": employee},
                                  {"nguyen ngoc thanh": ["1"]}, {})
        self.assertEqual(report.match_identity({"name": "Thành"}, {"1": employee},
                                               {}, {"thanh": "1"}), "1")

    def test_raw_excel_matches_csv_with_typed_dates_and_numeric_ids(self):
        employees = {"1": report.base.Employee("1", "A", "Phòng")}
        events = [
            ["2026-10-23T06:22:04+07:00", "1", "A", "1", "1", "faceOrFpOrCardOrPw"],
            ["2026-10-23T08:03:06+07:00", "1", "A", "1", "1", "faceOrFpOrCardOrPw"],
            ["2026-10-23T17:03:15+07:00", "1", "A", "1", "1", "faceOrFpOrCardOrPw"],
            ["2026-10-23T17:03:15+07:00", "1", "A", "1", "1", "faceOrFpOrCardOrPw"],
            ["2026-10-23T07:00:00+07:00", "", "", "", "1", "invalid"],
        ]
        with tempfile.TemporaryDirectory() as directory:
            csv_path, xlsx_path = Path(directory) / "raw.csv", Path(directory) / "raw.xlsx"
            with csv_path.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(report.RAW_FIELDS)
                writer.writerows(events)
            book = Workbook()
            book.active.title = "Ghi chú"
            book.active.append(["Dữ liệu RAW ở trang tiếp theo"])
            sheet = book.create_sheet("RAW")
            sheet.append(["Xuất từ máy"])
            sheet.append(report.RAW_FIELDS)
            for index, event in enumerate(events):
                row = list(event)
                row[1] = int(row[1]) if row[1] else None
                if index == 0:
                    row[0] = datetime(2026, 10, 23, 6, 22, 4)
                elif index == 1:
                    row[0] = "23/10/2026 08:03:06"
                sheet.append(row)
            book.save(xlsx_path)
            counters = [Counter(), Counter()]
            issues = [[], []]
            csv_result = report.load_raw(csv_path, employees, counters[0], issues[0])
            xlsx_result = report.load_raw(xlsx_path, employees, counters[1], issues[1])
        self.assertEqual(xlsx_result, csv_result)
        self.assertEqual(counters[0], counters[1])
        self.assertEqual(issues, [[], []])

    def test_raw_excel_accepts_separate_date_and_time(self):
        employees = {"1": report.base.Employee("1", "A", "Phòng")}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "raw.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.append(["Ngày", "Thời gian", "Mã nhân viên"])
            sheet.append([datetime(2026, 10, 23), time(6, 22, 4), 1])
            sheet.append([datetime(2026, 10, 23), time(17, 3, 15), 1])
            book.save(path)
            issues = []
            _, punches, _ = report.load_raw(path, employees, Counter(), issues)
        self.assertEqual(issues, [])
        self.assertEqual(punches["1", date(2026, 10, 23)], {time(6, 22, 4), time(17, 3, 15)})

    def test_raw_excel_reports_original_sheet_row_for_bad_time(self):
        employees = {"1": report.base.Employee("1", "A", "Phòng")}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "raw.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "SuKien"
            sheet.append(["Xuất từ máy"])
            sheet.append(["time", "employeeNoString"])
            sheet.append([])
            sheet.append(["sai ngày giờ", 1])
            book.save(path)
            issues = []
            report.load_raw(path, employees, Counter(), issues)
        self.assertEqual(len(issues), 1)
        self.assertIn("trang SuKien, dòng 4", issues[0])


if __name__ == "__main__":
    unittest.main()
