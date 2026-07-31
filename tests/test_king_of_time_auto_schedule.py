import csv
import io
import unittest
from datetime import date

import pandas as pd

from king_of_time_auto_schedule import (
    CSV_COLUMNS,
    build_auto_schedule_export,
    calendar_week_number,
    csv_bytes,
    normalize_employee_code,
    weekday_number,
)


class KingOfTimeAutoScheduleTests(unittest.TestCase):
    def setUp(self):
        self.staff = [
            {"id": 1, "staff_name": "池田", "staff_code": "００１"},
            {"id": 2, "staff_name": "藤野", "staff_code": "A002"},
        ]
        self.patterns = {
            "日勤": {"pattern_code": "001", "pattern_name": "日勤"},
            "夜勤": {"pattern_code": "002", "pattern_name": "夜勤"},
            "管": {"pattern_code": "004", "pattern_name": "管理業務"},
        }
        self.settings = {}

    def build(self, rows, selected=(1, 2), confirmed=True, **kwargs):
        frame = pd.DataFrame(rows, columns=["id", "shift_date", "staff_name", "shift_kind"])
        return build_auto_schedule_export(
            2026, 7, frame, self.staff, selected, self.patterns, self.settings,
            is_confirmed=confirmed, **kwargs,
        )

    @staticmethod
    def parsed(output):
        return list(csv.reader(io.StringIO(output.decode("cp932"))))

    def test_csv_has_exactly_three_columns_in_required_order_and_no_header(self):
        self.assertEqual(CSV_COLUMNS, ["勤務日", "従業員コード", "パターンコード"])
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(self.parsed(output), [["20260701", "001", "001"]])

    def test_date_is_yyyymmdd_with_zero_padded_month_and_day(self):
        preview, _, _, _ = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertEqual(preview.iloc[0]["勤務日"], "20260701")

    def test_pattern_and_employee_leading_zeroes_are_preserved(self):
        preview, _, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertEqual(preview.iloc[0]["従業員コード"], "001")
        self.assertEqual(preview.iloc[0]["パターンコード"], "001")
        self.assertEqual(output.decode("cp932"), "20260701,001,001\r\n")

    def test_all_rows_have_three_columns_without_trailing_empty_column(self):
        _, errors, _, output = self.build([
            (1, "2026-07-01", "池田", "日勤"),
            (2, "2026-07-02", "藤野", "夜勤"),
        ])
        self.assertTrue(errors.empty)
        rows = self.parsed(output)
        self.assertTrue(rows)
        self.assertTrue(all(len(row) == 3 for row in rows))
        self.assertTrue(all(not line.endswith(",") for line in output.decode("cp932").splitlines()))

    def test_work_patterns_use_existing_mapping_including_management_aliases(self):
        preview, errors, _, output = self.build([
            (1, "2026-07-01", "池田", "日勤"),
            (2, "2026-07-02", "池田", "夜勤"),
            (3, "2026-07-03", "池田", "管理業務"),
        ], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(preview["パターンコード"].tolist(), ["001", "002", "004"])
        self.assertEqual(self.parsed(output), [
            ["20260701", "001", "001"],
            ["20260702", "001", "002"],
            ["20260703", "001", "004"],
        ])

    def test_non_work_shift_kinds_are_excluded(self):
        excluded = ["休", "休み", "有休", "希望休", "明", "明け", "夜勤明け"]
        rows = [(index, f"2026-07-{index:02d}", "池田", kind) for index, kind in enumerate(excluded, 1)]
        rows.append((20, "2026-07-20", "池田", "日勤"))
        preview, errors, warnings, output = self.build(rows, selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(preview["元のシフト区分"].tolist(), ["日勤"])
        self.assertEqual(len(warnings), len(excluded))
        self.assertEqual(self.parsed(output), [["20260720", "001", "001"]])

    def test_cp932_crlf_no_bom_and_no_blank_lines(self):
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(output.decode("cp932"), "20260701,001,001\r\n")
        self.assertNotIn(b"\n", output.replace(b"\r\n", b""))
        self.assertFalse(output.startswith(b"\xef\xbb\xbf"))
        self.assertNotIn("\r\n\r\n", output.decode("cp932"))

    def test_missing_employee_code_is_error_without_internal_id_fallback(self):
        self.staff[0]["staff_code"] = None
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertIn("登録されていません", errors.iloc[0]["内容"])
        self.assertIsNone(output)

    def test_invalid_employee_code_is_error(self):
        self.staff[0]["staff_code"] = "ABC-001"
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertIn("3～10文字", errors.iloc[0]["内容"])
        self.assertIsNone(output)

    def test_numeric_employee_code_stays_string(self):
        self.staff[0]["staff_code"] = 123
        preview, errors, _, _ = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(preview.iloc[0]["従業員コード"], "123")
        self.assertIsInstance(preview.iloc[0]["従業員コード"], str)

    def test_full_width_employee_code_is_normalized(self):
        self.assertEqual(normalize_employee_code("ＡＢ１２"), "AB12")

    def test_missing_pattern_code_is_error(self):
        self.patterns["日勤"]["pattern_code"] = ""
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertIn("パターンコード", errors.iloc[0]["内容"])
        self.assertIsNone(output)

    def test_unmapped_shift_is_warning_not_guessed(self):
        preview, errors, warnings, output = self.build([(1, "2026-07-01", "池田", "その他")], selected=(1,))
        self.assertTrue(preview.empty)
        self.assertTrue(errors.empty)
        self.assertIn("未設定", warnings.iloc[0]["内容"])
        self.assertEqual(output, b"")

    def test_unconfirmed_month_is_not_output(self):
        preview, errors, warnings, output = self.build(
            [(1, "2026-07-01", "池田", "日勤")], selected=(1,), confirmed=False,
        )
        self.assertTrue(preview.empty)
        self.assertFalse(errors.empty)
        self.assertFalse(warnings.empty)
        self.assertIsNone(output)

    def test_duplicate_employee_and_work_date_is_error(self):
        _, errors, _, output = self.build([
            (1, "2026-07-01", "池田", "日勤"),
            (2, "2026-07-01", "池田", "夜勤"),
        ], selected=(1,))
        self.assertIn("重複", " ".join(errors["内容"].tolist()))
        self.assertIsNone(output)

    def test_same_input_always_produces_date_order(self):
        rows = [
            (2, "2026-07-02", "藤野", "日勤"),
            (1, "2026-07-01", "池田", "日勤"),
        ]
        first = self.build(rows)[3]
        second = self.build(list(reversed(rows)))[3]
        self.assertEqual(first, second)
        self.assertEqual(self.parsed(first)[0][0], "20260701")

    def test_calendar_helpers_remain_unchanged(self):
        self.assertEqual(calendar_week_number(date(2026, 8, 1)), 1)
        self.assertEqual(calendar_week_number(date(2026, 8, 31)), 6)
        self.assertEqual([weekday_number(date(2026, 7, day)) for day in range(5, 12)], list(range(1, 8)))

    def test_empty_target_month_returns_existing_error(self):
        preview, errors, warnings, output = self.build([], selected=(1,))
        self.assertTrue(preview.empty)
        self.assertFalse(errors.empty)
        self.assertTrue(warnings.empty)
        self.assertIsNone(output)

    def test_only_selected_staff_is_output(self):
        preview, _, _, _ = self.build([
            (1, "2026-07-01", "池田", "日勤"),
            (2, "2026-07-01", "藤野", "日勤"),
        ], selected=(2,))
        self.assertEqual(preview["職員名"].tolist(), ["藤野"])

    def test_no_selected_staff_cannot_generate(self):
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=())
        self.assertFalse(errors.empty)
        self.assertIsNone(output)


if __name__ == "__main__":
    unittest.main()
