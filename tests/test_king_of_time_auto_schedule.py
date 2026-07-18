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
            "日勤": {"pattern_code": "DAY001", "pattern_name": "日勤", "day_type_code": "1", "day_type_name": "平日"},
            "夜勤": {"pattern_code": "NGT001", "pattern_name": "夜勤", "day_type_code": "1", "day_type_name": "平日"},
            "管": {"pattern_code": "MGR001", "pattern_name": "管理勤務", "day_type_code": "1", "day_type_name": "平日"},
        }
        self.settings = {
            "rest_day_type_code": "3", "rest_day_type_name": "法定外休日", "rest_leave_name": "公休",
            "paid_day_type_code": "1", "paid_day_type_name": "平日", "paid_leave_name": "有休",
            "statutory_weekday": "", "holiday_day_type_code": "1",
        }

    def build(self, rows, selected=(1, 2), confirmed=True, **kwargs):
        frame = pd.DataFrame(rows, columns=["id", "shift_date", "staff_name", "shift_kind"])
        return build_auto_schedule_export(
            2026, 7, frame, self.staff, selected, self.patterns, self.settings,
            is_confirmed=confirmed, **kwargs,
        )

    def test_csv_columns_are_exactly_the_ten_requested_columns(self):
        self.assertEqual(len(CSV_COLUMNS), 10)
        content = csv_bytes([]).decode("cp932")
        self.assertEqual(next(csv.reader(io.StringIO(content))), CSV_COLUMNS)

    def test_employee_code_is_output_and_full_width_is_normalized(self):
        preview, errors, _, _ = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(preview.iloc[0]["従業員コード"], "001")
        self.assertEqual(normalize_employee_code("ＡＢ１２"), "AB12")

    def test_missing_employee_code_is_error(self):
        self.staff[0]["staff_code"] = None
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertIn("登録されていません", errors.iloc[0]["内容"])
        self.assertIsNone(output)

    def test_invalid_employee_code_is_error(self):
        self.staff[0]["staff_code"] = "ABC-001"
        _, errors, _, _ = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertIn("3～10文字", errors.iloc[0]["内容"])

    def test_calendar_week_number_uses_sunday_calendar_rows(self):
        self.assertEqual(calendar_week_number(date(2026, 8, 1)), 1)
        self.assertEqual(calendar_week_number(date(2026, 8, 2)), 2)
        self.assertEqual(calendar_week_number(date(2026, 8, 31)), 6)

    def test_weekday_number_is_sunday_one_through_saturday_seven(self):
        expected = [1, 2, 3, 4, 5, 6, 7]
        actual = [weekday_number(date(2026, 7, day)) for day in range(5, 12)]
        self.assertEqual(actual, expected)

    def test_holiday_can_be_number_eight(self):
        preview, errors, _, _ = self.build(
            [(1, "2026-07-20", "池田", "日勤")], selected=(1,), holiday_dates={date(2026, 7, 20)},
        )
        self.assertTrue(errors.empty)
        self.assertEqual((preview.iloc[0]["曜日の番号"], preview.iloc[0]["曜日"]), (8, "祝"))

    def test_holiday_can_keep_normal_weekday_number(self):
        preview, _, _, _ = self.build(
            [(1, "2026-07-20", "池田", "日勤")], selected=(1,),
            holiday_dates={date(2026, 7, 20)}, holiday_as_eight=False,
        )
        self.assertEqual(preview.iloc[0]["曜日の番号"], 2)

    def test_day_shift_pattern_is_output(self):
        preview, _, _, _ = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertEqual(preview.iloc[0]["パターンコード"], "DAY001")

    def test_night_entry_is_output_but_night_after_is_not(self):
        preview, errors, warnings, _ = self.build([
            (1, "2026-07-01", "池田", "夜勤"),
            (2, "2026-07-02", "池田", "夜勤明け"),
        ], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(preview["元のシフト区分"].tolist(), ["夜勤"])
        self.assertIn("明", warnings.iloc[0]["内容"])

    def test_rest_is_non_statutory_holiday_and_public_rest(self):
        preview, _, _, _ = self.build([(1, "2026-07-01", "池田", "休み")], selected=(1,))
        row = preview.iloc[0]
        self.assertEqual((row["勤務日種別コード"], row["勤務日種別名"], row["休暇区分名"]), ("3", "法定外休日", "公休"))

    def test_paid_leave_has_configured_leave_name(self):
        preview, _, _, _ = self.build([(1, "2026-07-01", "池田", "有休")], selected=(1,))
        self.assertEqual(preview.iloc[0]["休暇区分名"], "有休")

    def test_unconfirmed_month_is_not_output(self):
        preview, errors, warnings, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,), confirmed=False)
        self.assertTrue(preview.empty)
        self.assertFalse(errors.empty)
        self.assertFalse(warnings.empty)
        self.assertIsNone(output)

    def test_missing_pattern_code_is_error(self):
        self.patterns["日勤"]["pattern_code"] = ""
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertIn("パターンコード", errors.iloc[0]["内容"])
        self.assertIsNone(output)

    def test_duplicate_employee_week_weekday_is_error(self):
        _, errors, _, output = self.build([
            (1, "2026-07-01", "池田", "日勤"),
            (2, "2026-07-01", "池田", "休み"),
        ], selected=(1,))
        self.assertIn("重複", " ".join(errors["内容"].tolist()))
        self.assertIsNone(output)

    def test_only_selected_staff_is_output(self):
        preview, _, _, _ = self.build([
            (1, "2026-07-01", "池田", "日勤"),
            (2, "2026-07-01", "藤野", "日勤"),
        ], selected=(2,))
        self.assertEqual(preview["名前"].tolist(), ["藤野"])

    def test_no_selected_staff_cannot_generate(self):
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=())
        self.assertFalse(errors.empty)
        self.assertIsNone(output)

    def test_preview_only_columns_are_not_in_final_csv(self):
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertTrue(errors.empty)
        header = output.decode("cp932").splitlines()[0]
        self.assertNotIn("元の日付", header)
        self.assertNotIn("元のシフト区分", header)

    def test_cp932_crlf_and_no_bom(self):
        _, _, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertEqual(output[:2], "従".encode("cp932")[:2])
        self.assertIn(b"\r\n", output)
        self.assertNotIn(b"\n", output.replace(b"\r\n", b""))
        self.assertFalse(output.startswith(b"\xef\xbb\xbf"))
        self.assertIn("池田", output.decode("cp932"))

    def test_blank_values_never_become_nan(self):
        preview, _, _, output = self.build([(1, "2026-07-01", "池田", "休み")], selected=(1,))
        self.assertEqual(preview.iloc[0]["パターン名"], "")
        self.assertNotIn("nan", output.decode("cp932").lower())

    def test_same_input_always_produces_same_order(self):
        rows = [
            (2, "2026-07-02", "藤野", "日勤"),
            (1, "2026-07-01", "池田", "日勤"),
        ]
        first = self.build(rows)[3]
        second = self.build(list(reversed(rows)))[3]
        self.assertEqual(first, second)

    def test_unmapped_shift_is_warning_not_silently_converted(self):
        preview, errors, warnings, _ = self.build([(1, "2026-07-01", "池田", "その他")], selected=(1,))
        self.assertTrue(preview.empty)
        self.assertTrue(errors.empty)
        self.assertIn("未設定", warnings.iloc[0]["内容"])

    def test_statutory_weekday_setting_changes_rest_day_type(self):
        self.settings["statutory_weekday"] = "4"  # 2026-07-01 is Wednesday.
        preview, _, _, _ = self.build([(1, "2026-07-01", "池田", "休み")], selected=(1,))
        self.assertEqual((preview.iloc[0]["勤務日種別コード"], preview.iloc[0]["勤務日種別名"]), ("2", "法定休日"))

    def test_leading_zero_employee_codes_are_preserved_as_strings(self):
        for code in ("001", "0001", "123", "A001", "０１２３", "1234567890"):
            with self.subTest(code=code):
                self.staff[0]["staff_code"] = code
                preview, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
                expected = normalize_employee_code(code)
                self.assertTrue(errors.empty)
                self.assertIsInstance(preview.iloc[0]["従業員コード"], str)
                self.assertEqual(preview.iloc[0]["従業員コード"], expected)
                self.assertIn(f"{expected},池田,".encode("cp932"), output)

    def test_ten_character_employee_code_is_allowed(self):
        self.staff[0]["staff_code"] = "1234567890"
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertIsNotNone(output)

    def test_eleven_character_employee_code_is_error(self):
        self.staff[0]["staff_code"] = "12345678901"
        _, errors, _, output = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertIn("3～10文字", errors.iloc[0]["内容"])
        self.assertIsNone(output)

    def test_cp932_unencodable_name_is_specific_error(self):
        self.staff[0]["staff_name"] = "𠮷田"
        _, errors, _, output = self.build([(1, "2026-07-01", "𠮷田", "日勤")], selected=(1,))
        self.assertIn("名前「𠮷田」", " ".join(errors["内容"].tolist()))
        self.assertIn("cp932", " ".join(errors["内容"].tolist()))
        self.assertIsNone(output)

    def test_requested_japanese_values_are_cp932_encodable(self):
        for value in ("池田", "藤野", "髙橋", "﨑田", "管理勤務", "法定外休日", "有休", "公休"):
            with self.subTest(value=value):
                self.assertEqual(value.encode("cp932").decode("cp932"), value)

    def test_week_five_and_six_are_calculated(self):
        self.assertEqual(calendar_week_number(date(2026, 8, 23)), 5)
        self.assertEqual(calendar_week_number(date(2026, 8, 30)), 6)

    def test_month_starting_saturday_keeps_day_one_in_week_one(self):
        self.assertEqual(date(2026, 8, 1).weekday(), 5)
        self.assertEqual(calendar_week_number(date(2026, 8, 1)), 1)

    def test_month_starting_sunday_keeps_first_row_as_week_one(self):
        self.assertEqual(date(2026, 2, 1).weekday(), 6)
        self.assertEqual(calendar_week_number(date(2026, 2, 1)), 1)
        self.assertEqual(calendar_week_number(date(2026, 2, 7)), 1)

    def test_empty_target_month_returns_error_without_crash(self):
        preview, errors, warnings, output = self.build([], selected=(1,))
        self.assertTrue(preview.empty)
        self.assertFalse(errors.empty)
        self.assertTrue(warnings.empty)
        self.assertIsNone(output)

    def test_every_csv_row_has_exactly_ten_columns_and_trailing_newline(self):
        _, errors, _, output = self.build([
            (1, "2026-07-01", "池田", "日勤"),
            (2, "2026-07-02", "藤野", "休み"),
        ])
        self.assertTrue(errors.empty)
        parsed = list(csv.reader(io.StringIO(output.decode("cp932"))))
        self.assertTrue(all(len(row) == 10 for row in parsed))
        self.assertTrue(output.endswith(b"\r\n"))

    def test_numeric_employee_code_stays_a_string(self):
        self.staff[0]["staff_code"] = 123
        preview, errors, _, _ = self.build([(1, "2026-07-01", "池田", "日勤")], selected=(1,))
        self.assertTrue(errors.empty)
        self.assertEqual(preview.iloc[0]["従業員コード"], "123")
        self.assertIsInstance(preview.iloc[0]["従業員コード"], str)


if __name__ == "__main__":
    unittest.main()
