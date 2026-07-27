import functools
import inspect
import sys
import types
import unittest
from datetime import date, datetime
from unittest.mock import Mock, patch

import pandas as pd


class _CacheData:
    def __call__(self, func=None, **_kwargs):
        if func is None:
            return lambda target: self(target)
        cache = {}

        @functools.wraps(func)
        def wrapped(*args, **kwargs):
            bound = inspect.signature(func).bind(*args, **kwargs)
            key = repr(tuple(
                (name, value)
                for name, value in bound.arguments.items()
                if not name.startswith("_")
            ))
            if key not in cache:
                cache[key] = func(*args, **kwargs)
            return cache[key]

        wrapped.clear = cache.clear
        return wrapped


fake_streamlit = types.ModuleType("streamlit")
fake_streamlit.cache_data = _CacheData()
fake_streamlit.session_state = {}
fake_streamlit.secrets = {}
sys.modules.setdefault("streamlit", fake_streamlit)

import app  # noqa: E402
import db  # noqa: E402


def _status(confirmed=False, error=""):
    return {
        "is_confirmed": 1 if confirmed else 0,
        "confirmed_at": "2026-07-20" if confirmed else "",
        "confirmed_by": "tester" if confirmed else "",
        "status_error": error,
    }


def _shift_row(row_id, shift_date, staff, kind, start="", end="", memo=""):
    return {
        "id": row_id,
        "shift_date": shift_date,
        "staff_name": staff,
        "shift_kind": kind,
        "start_time": start,
        "end_time": end,
        "next_day": 0,
        "memo": memo,
        "created_at": "2026-06-01 00:00:00",
        "updated_at": "2026-06-01 00:00:00",
    }


class ShiftSafetyTests(unittest.TestCase):
    def test_confirmed_month_rejects_every_staff_shift_write_path(self):
        matrix = pd.DataFrame([{"職員名": "田中", "1": "日"}])
        draft = pd.DataFrame([{
            "日付": "2026-07-01", "勤務": "日勤", "候補職員": "田中",
            "理由": "test", "保存対象": True,
        }])
        current = pd.DataFrame([_shift_row(1, "2026-07-01", "田中", "日勤")])

        routes = [
            lambda: app.save_single_shift("2026-07-01", "田中", "日勤"),
            lambda: app.save_day_shift_assignments("2026-07-01", {"田中": "日"}),
            lambda: app.save_basic_day_shift("2026-07-01", "田中", "", ""),
            lambda: app.save_shift_matrix_from_editor(2026, 7, matrix),
            lambda: app.save_ai_shift_draft_rows(draft),
            lambda: app.clear_month_staff_shifts(2026, 7),
            lambda: app.update_staff_shift(1, "2026-07-01", "田中", "夜勤"),
            lambda: app.delete_staff_shift(1),
        ]

        with (
            patch.object(app, "get_shift_month_status", return_value=_status(confirmed=True)),
            patch.object(app, "fetch_df", return_value=current),
            patch.object(app, "execute") as execute,
            patch.object(app, "execute_many") as execute_many,
            patch.object(app, "execute_transaction") as execute_transaction,
        ):
            for route in routes:
                with self.subTest(route=inspect.getsource(route).strip()):
                    with self.assertRaises(app.ShiftUpdateBlockedError):
                        route()

        execute.assert_not_called()
        execute_many.assert_not_called()
        execute_transaction.assert_not_called()

    def test_unconfirmed_month_allows_save(self):
        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "fetch_staff_shifts_raw", return_value=pd.DataFrame()),
            patch.object(app, "execute", return_value=123) as execute,
        ):
            result = app.save_single_shift("2026-07-01", "田中", "日勤")

        self.assertEqual((result.status, result.count), (app.SHIFT_SAVE_SAVED, 1))
        execute.assert_called_once()

    def test_status_lookup_failure_rejects_update(self):
        with (
            patch.object(app, "get_shift_month_status", return_value=_status(error="db unavailable")),
            patch.object(app, "execute") as execute,
        ):
            with self.assertRaises(app.ShiftUpdateBlockedError):
                app.save_single_shift("2026-07-01", "田中", "日勤")
        execute.assert_not_called()

    def test_status_reader_marks_database_failure_instead_of_assuming_draft(self):
        app.get_shift_month_status.clear()
        with patch.object(app, "fetch_df", side_effect=RuntimeError("database offline")):
            status = app.get_shift_month_status(2026, 7)
        app.get_shift_month_status.clear()

        self.assertFalse(status["is_confirmed"])
        self.assertIn("database offline", status["status_error"])

    def test_fatal_errors_block_confirmation_and_clean_month_can_confirm(self):
        errors = pd.DataFrame([{
            "区分": "必要人数不足", "重要度": "高", "日付": "2026-07-01",
            "職員名": "", "内容": "日勤あと1",
        }])
        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "set_shift_month_status") as set_status,
        ):
            blocked = app.confirm_shift_month_if_valid(2026, 7, errors, "tester")
            self.assertEqual(blocked.status, app.SHIFT_SAVE_BLOCKED)
            set_status.assert_not_called()

            confirmed = app.confirm_shift_month_if_valid(2026, 7, pd.DataFrame(), "tester")
            self.assertEqual(confirmed.status, app.SHIFT_SAVE_SAVED)
            set_status.assert_called_once_with(2026, 7, True, "tester")

    def test_confirmation_rules_separate_staffing_shortages_and_quality_warnings(self):
        shortage = pd.DataFrame([{"日付": "2026-07-01", "不足": "夜勤あと1"}])
        quality = pd.DataFrame([
            {"重要度": "中", "種類": "5連勤", "日付": "2026-07-05", "職員名": "A", "内容": "注意"},
            {"重要度": "高", "種類": "6連勤", "日付": "2026-07-06", "職員名": "B", "内容": "要調整"},
        ])
        limits = pd.DataFrame([{"重要度": "中", "種類": "合計上限超過", "職員名": "C", "内容": "21/20"}])
        duplicates = pd.DataFrame([{"種類": "同一職員・同一日の重複", "日付": "2026-07-07", "職員名": "D", "内容": "2件"}])
        kot = pd.DataFrame([{"職員名": "E", "勤務日": "2026-07-08", "項目": "従業員コード", "内容": "未登録"}])

        errors = app.build_shift_confirmation_errors(shortage, quality, limits, duplicates, kot)
        warnings = app.build_shift_quality_warnings(quality, limits, duplicates, kot)

        self.assertEqual(errors["区分"].tolist(), ["必要人数不足"])
        self.assertEqual(len(warnings), 5)
        self.assertIn("5連勤", warnings["区分"].tolist())
        self.assertIn("6連勤", warnings["区分"].tolist())
        self.assertIn("合計上限超過", warnings["区分"].tolist())
        self.assertIn("KING OF TIME確認事項", warnings["区分"].tolist())

    def test_day_night_and_combined_shortages_are_detailed_and_block_confirmation(self):
        cases = (
            ("日勤不足", 1, 1, ["日勤"]),
            ("夜勤不足", 2, 0, ["夜勤"]),
            ("日勤夜勤不足", 0, 0, ["日勤", "夜勤"]),
        )
        for label, day_count, night_count, expected_kinds in cases:
            with self.subTest(label=label):
                shortage = pd.DataFrame([{
                    "日付": "2026-07-01",
                    "日勤人数": day_count,
                    "夜勤人数": night_count,
                    "状態": "要確認",
                }])
                staffing_shortages = app.build_shift_confirmation_errors(
                    shortage, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
                )
                self.assertEqual(staffing_shortages["勤務区分"].tolist(), expected_kinds)
                self.assertFalse(app.can_finalize_shift(staffing_shortages))
                with (
                    patch.object(app, "get_shift_month_status", return_value=_status()),
                    patch.object(app, "set_shift_month_status") as set_status,
                ):
                    result = app.confirm_shift_month_if_valid(
                        2026, 7, staffing_shortages, "tester"
                    )
                self.assertEqual(result.status, app.SHIFT_SAVE_BLOCKED)
                self.assertIn("必要人数", result.message)
                set_status.assert_not_called()

    def test_quality_warnings_do_not_block_confirmation(self):
        quality = pd.DataFrame([{
            "重要度": "高", "種類": "6連勤", "日付": "2026-07-06",
            "職員名": "A", "内容": "要調整",
        }])
        limits = pd.DataFrame([{
            "重要度": "高", "種類": "日勤上限超過",
            "職員名": "A", "内容": "11/10",
        }])
        warnings = app.build_shift_quality_warnings(
            quality, limits, pd.DataFrame(), pd.DataFrame()
        )
        staffing_shortages = app.build_shift_confirmation_errors(
            pd.DataFrame(), quality, limits, pd.DataFrame()
        )

        self.assertEqual(len(warnings), 2)
        self.assertTrue(staffing_shortages.empty)
        self.assertTrue(app.can_finalize_shift(staffing_shortages))
        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "set_shift_month_status") as set_status,
        ):
            result = app.confirm_shift_month_if_valid(
                2026, 7, staffing_shortages, "tester"
            )
        self.assertEqual(result.status, app.SHIFT_SAVE_SAVED)
        set_status.assert_called_once_with(2026, 7, True, "tester")

    def test_monthly_editor_updates_only_changed_cell_and_preserves_other_records(self):
        month_rows = pd.DataFrame([
            _shift_row(1, "2026-07-01", "田中", "休み", memo="通常休"),
            _shift_row(2, "2026-07-02", "田中", "希望休", memo="本人希望"),
            _shift_row(3, "2026-07-03", "田中", "有休", memo="有休申請"),
            _shift_row(4, "2026-07-04", "田中", "日勤", "09:00", "18:00", "個別時刻"),
            _shift_row(5, "2026-07-05", "田中", "日勤", "08:30", "17:30", "変更対象"),
            _shift_row(6, "2026-07-05", "佐藤", "夜勤", "16:30", "09:30", "別職員"),
        ])
        row = {"職員名": "田中", **{str(day): "" for day in range(1, 32)}}
        row.update({"1": "休", "2": "希", "3": "有", "4": "日", "5": "夜"})
        edited = pd.DataFrame([row])
        transaction = Mock()

        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "fetch_staff_shifts_raw", side_effect=[month_rows, month_rows]),
            patch.object(app, "execute_transaction", transaction),
        ):
            result = app.save_shift_matrix_from_editor(2026, 7, edited)

        self.assertEqual((result.status, result.changed_cells), (app.SHIFT_SAVE_SAVED, 1))
        operations = transaction.call_args.args[0]
        flattened_params = [operation[1] for operation in operations]
        self.assertIn(("2026-07-05", "田中"), flattened_params)
        for protected_date in ("2026-07-01", "2026-07-02", "2026-07-03", "2026-07-04"):
            self.assertFalse(any(protected_date in str(params) for params in flattened_params))
        self.assertFalse(any("佐藤" in str(params) for params in flattened_params))
        insert_rows = operations[-1][1]
        self.assertEqual(insert_rows[0][0:3], ("2026-07-05", "田中", "夜勤"))

    def test_unchanged_monthly_editor_performs_no_write(self):
        month_rows = pd.DataFrame([
            _shift_row(1, "2026-07-01", "田中", "休み", memo="保持"),
            _shift_row(2, "2026-07-02", "田中", "希望休", memo="保持"),
            _shift_row(3, "2026-07-03", "田中", "有休", memo="保持"),
            _shift_row(4, "2026-07-04", "田中", "日勤", "09:00", "18:00", "保持"),
        ])
        row = {"職員名": "田中", **{str(day): "" for day in range(1, 32)}}
        row.update({"1": "休", "2": "希", "3": "有", "4": "日"})
        edited = pd.DataFrame([row])

        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "fetch_staff_shifts_raw", side_effect=[month_rows, month_rows]),
            patch.object(app, "execute_transaction") as transaction,
        ):
            result = app.save_shift_matrix_from_editor(2026, 7, edited)

        self.assertEqual(result.status, app.SHIFT_SAVE_NO_CHANGE)
        transaction.assert_not_called()

    def test_blank_change_deletes_only_the_selected_staff_and_date(self):
        month_rows = pd.DataFrame([
            _shift_row(1, "2026-07-10", "田中", "休み", "", "", "削除対象"),
            _shift_row(2, "2026-07-10", "佐藤", "夜勤", "16:30", "09:30", "保持"),
            _shift_row(3, "2026-08-10", "田中", "日勤", "09:00", "18:00", "別月"),
        ])
        row = {"職員名": "田中", **{str(day): "" for day in range(1, 32)}}
        edited = pd.DataFrame([row])
        transaction = Mock()

        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "fetch_staff_shifts_raw", side_effect=[month_rows, month_rows]),
            patch.object(app, "execute_transaction", transaction),
        ):
            result = app.save_shift_matrix_from_editor(2026, 7, edited)

        self.assertEqual((result.status, result.changed_cells, result.count), (app.SHIFT_SAVE_SAVED, 1, 0))
        operations = transaction.call_args.args[0]
        self.assertEqual(len(operations), 1)
        self.assertEqual(operations[0][1], ("2026-07-10", "田中"))

    def test_duplicate_save_is_not_reported_as_success(self):
        existing = pd.DataFrame([_shift_row(1, "2026-07-01", "田中", "日勤")])
        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "fetch_staff_shifts_raw", return_value=existing),
            patch.object(app, "execute") as execute,
        ):
            result = app.save_single_shift("2026-07-01", "田中", "日勤")

        self.assertEqual(result.status, app.SHIFT_SAVE_DUPLICATE)
        execute.assert_not_called()


class ShiftNightRuleTests(unittest.TestCase):
    @staticmethod
    def quality(rows, year=2026, month=7):
        return app.create_shift_quality_check_table(pd.DataFrame(rows), year, month)

    @staticmethod
    def row(shift_date, staff_name, shift_kind, **extra):
        return {
            "shift_date": shift_date,
            "staff_name": staff_name,
            "shift_kind": shift_kind,
            **extra,
        }

    def test_shift_kind_classification_is_explicit(self):
        expected = {
            "日": app.SHIFT_CATEGORY_ACTUAL_WORK,
            "管": app.SHIFT_CATEGORY_ACTUAL_WORK,
            "夜": app.SHIFT_CATEGORY_NIGHT,
            "明": app.SHIFT_CATEGORY_NIGHT_AFTER,
            "休": app.SHIFT_CATEGORY_OFF,
            "希": app.SHIFT_CATEGORY_OFF,
            "有": app.SHIFT_CATEGORY_OFF,
            "他": app.SHIFT_CATEGORY_ACTUAL_WORK,
            "": app.SHIFT_CATEGORY_BLANK,
        }
        for symbol, category in expected.items():
            with self.subTest(symbol=symbol):
                self.assertEqual(app.classify_shift_kind_for_quality_check(symbol), category)

    def test_night_to_night_after_is_normal_even_with_next_day_end_flag(self):
        checks = self.quality([
            self.row("2026-07-09", "田中", "夜勤", start_time="16:30", end_time="09:30", next_day=1),
            self.row("2026-07-10", "田中", "夜勤明け"),
        ])

        night_checks = checks[checks["種類"].astype(str).str.contains("夜勤|明け翌日", regex=True)]
        self.assertTrue(night_checks.empty)
        errors = app.build_shift_confirmation_errors(
            pd.DataFrame(), checks, pd.DataFrame(), pd.DataFrame()
        )
        self.assertTrue(errors.empty)

    def test_night_to_actual_work_is_high_error_with_specific_dates_and_symbol(self):
        for kind, symbol in (("日勤", "日"), ("管", "管"), ("夜勤", "夜"), ("その他", "他")):
            with self.subTest(kind=kind):
                checks = self.quality([
                    self.row("2026-07-09", "田中", "夜勤"),
                    self.row("2026-07-10", "田中", kind),
                ])
                errors = checks[
                    (checks["種類"] == "夜勤翌日勤務")
                    & (checks["重要度"] == "高")
                ]
                self.assertEqual(len(errors), 1)
                detail = errors.iloc[0]["内容"]
                self.assertIn("7月9日の夜勤", detail)
                self.assertIn("7月10日", detail)
                self.assertIn(f"「{symbol}」", detail)

    def test_missing_night_after_is_medium_warning_for_blank_and_leave(self):
        for kind, symbol in ((None, "空欄"), ("休み", "休"), ("希望休", "希"), ("有休", "有")):
            with self.subTest(kind=kind):
                rows = [self.row("2026-07-09", "田中", "夜勤")]
                if kind:
                    rows.append(self.row("2026-07-10", "田中", kind))
                checks = self.quality(rows)
                warning = checks[checks["種類"] == "夜勤明け未登録"]
                self.assertEqual(len(warning), 1)
                self.assertEqual(warning.iloc[0]["重要度"], "中")
                self.assertIn(f"翌日の登録：{symbol}", warning.iloc[0]["内容"])
                confirmation_errors = app.build_shift_confirmation_errors(
                    pd.DataFrame(), checks, pd.DataFrame(), pd.DataFrame()
                )
                self.assertNotIn("夜勤明け未登録", confirmation_errors.get("区分", pd.Series(dtype=str)).tolist())

    def test_night_after_then_day_is_attention_only_and_does_not_block_confirmation(self):
        checks = self.quality([
            self.row("2026-07-09", "田中", "夜勤"),
            self.row("2026-07-10", "田中", "夜勤明け"),
            self.row("2026-07-11", "田中", "日勤"),
        ])
        attention = checks[checks["種類"] == "明け翌日勤務（注意）"]
        self.assertEqual(len(attention), 1)
        self.assertEqual(attention.iloc[0]["重要度"], "中")
        confirmation_errors = app.build_shift_confirmation_errors(
            pd.DataFrame(), checks, pd.DataFrame(), pd.DataFrame()
        )
        self.assertTrue(confirmation_errors.empty)

        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "set_shift_month_status") as set_status,
        ):
            result = app.confirm_shift_month_if_valid(2026, 7, confirmation_errors, "tester")
        self.assertEqual(result.status, app.SHIFT_SAVE_SAVED)
        set_status.assert_called_once()

    def test_month_year_leap_boundaries_and_date_types_are_normalized(self):
        cases = (
            (2026, 8, date(2026, 7, 31), datetime(2026, 8, 1, 0, 0)),
            (2027, 1, datetime(2026, 12, 31, 16, 30), "2027-01-01 00:00:00"),
            (2024, 3, "2024-02-29 16:30:00", date(2024, 3, 1)),
        )
        for year, month, night_date, after_date in cases:
            with self.subTest(year=year, month=month):
                checks = self.quality([
                    self.row(night_date, "田中　太郎", "夜勤"),
                    self.row(after_date, "田中 太郎", "夜勤明け"),
                ], year, month)
                self.assertNotIn("夜勤翌日勤務", checks["種類"].tolist())
                self.assertNotIn("夜勤明け未登録", checks["種類"].tolist())

    def test_other_staff_next_day_work_is_not_mistaken_for_night_staff(self):
        checks = self.quality([
            self.row("2026-07-09", "田中", "夜勤"),
            self.row("2026-07-10", "佐藤", "日勤"),
        ])
        self.assertTrue(checks[checks["種類"] == "夜勤翌日勤務"].empty)

    def test_real_night_next_day_error_is_warning_and_does_not_block_confirmation(self):
        checks = self.quality([
            self.row("2026-07-09", "田中", "夜勤"),
            self.row("2026-07-10", "田中", "日勤"),
        ])
        staffing_shortages = app.build_shift_confirmation_errors(
            pd.DataFrame(), checks, pd.DataFrame(), pd.DataFrame()
        )
        warnings = app.build_shift_quality_warnings(
            checks, pd.DataFrame(), pd.DataFrame()
        )
        self.assertTrue(staffing_shortages.empty)
        self.assertIn("夜勤翌日勤務", warnings["区分"].tolist())
        with (
            patch.object(app, "get_shift_month_status", return_value=_status()),
            patch.object(app, "set_shift_month_status") as set_status,
        ):
            result = app.confirm_shift_month_if_valid(2026, 7, staffing_shortages, "tester")
        self.assertEqual(result.status, app.SHIFT_SAVE_SAVED)
        set_status.assert_called_once_with(2026, 7, True, "tester")


class ShiftManagerUiHelperTests(unittest.TestCase):
    def test_month_navigation_handles_year_boundaries(self):
        self.assertEqual(app.calculate_shift_month(2026, 7, -1), (2026, 6))
        self.assertEqual(app.calculate_shift_month(2026, 7, 1), (2026, 8))
        self.assertEqual(app.calculate_shift_month(2026, 12, 1), (2027, 1))
        self.assertEqual(app.calculate_shift_month(2026, 1, -1), (2025, 12))

        state = {"shift_year": 2025, "shift_month": 3}
        with (
            patch.object(app.st, "session_state", state),
            patch.object(app, "today_jst", return_value=date(2026, 7, 20)),
        ):
            app.move_shift_month(use_current=True)
        self.assertEqual((state["shift_year"], state["shift_month"]), (2026, 7))

    def test_month_move_clears_only_transient_state_and_keeps_month_save_times(self):
        state = {
            "shift_active_month": (2026, 7),
            "selected_shift_date": "2026-07-10",
            "ai_shift_draft": pd.DataFrame([{"value": 1}]),
            "shift_matrix_editor_2026_7_全職員_0": {"edited_rows": {}},
            "shift_staff_filter_2026_7": "田中",
            "clear_month_shift_confirm_2026_7": True,
            "king_of_time_selected_staff_2026_7": [1],
            "shift_last_saved_at_2026_7": "2026年7月20日 21:35",
            "shift_last_saved_at_2026_8": "2026年8月20日 20:00",
            "unrelated": "keep",
        }

        app.clear_shift_month_transient_state(state, 2026, 7)

        self.assertNotIn("selected_shift_date", state)
        self.assertNotIn("ai_shift_draft", state)
        self.assertFalse(any(key.startswith("shift_matrix_editor_2026_7_") for key in state))
        self.assertNotIn("shift_staff_filter_2026_7", state)
        self.assertNotIn("clear_month_shift_confirm_2026_7", state)
        self.assertNotIn("king_of_time_selected_staff_2026_7", state)
        self.assertEqual(state["shift_last_saved_at_2026_7"], "2026年7月20日 21:35")
        self.assertEqual(state["shift_last_saved_at_2026_8"], "2026年8月20日 20:00")
        self.assertEqual(state["unrelated"], "keep")
        self.assertEqual(state["shift_editor_reset_counter"], 1)

        state["selected_shift_date"] = "2026-07-11"
        self.assertTrue(app.sync_shift_month_state(state, 2026, 8))
        self.assertEqual(state["shift_active_month"], (2026, 8))
        self.assertNotIn("selected_shift_date", state)
        self.assertFalse(app.sync_shift_month_state(state, 2026, 8))

    def test_status_display_distinguishes_draft_confirmed_and_lookup_failure(self):
        self.assertEqual(app.shift_month_status_display(_status()), ("info", "📝", "作成中"))
        self.assertEqual(
            app.shift_month_status_display(_status(confirmed=True)),
            ("success", "🔒", "確定済み"),
        )
        self.assertEqual(
            app.shift_month_status_display(_status(error="db unavailable")),
            ("error", "⚠️", "状態を確認できません"),
        )

        state = {"shift_last_saved_at_2026_7": "2026年7月20日 21:35"}
        with (
            patch.object(app.st, "session_state", state),
            patch.object(app.st, "success", create=True) as success,
            patch.object(app.st, "caption", create=True) as caption,
        ):
            app.render_shift_month_status_header(2026, 7, _status(confirmed=True))
        header_text = success.call_args.args[0]
        self.assertIn("対象年月：2026年7月", header_text)
        self.assertIn("状態：確定済み", header_text)
        self.assertIn("確定日時：2026-07-20", header_text)
        caption.assert_called_once_with("最終保存：2026年7月20日 21:35")

    def test_shift_legend_matches_editor_symbols_and_configured_shift_kinds(self):
        legend = dict(app.SHIFT_LEGEND_ITEMS)
        self.assertEqual(set(legend), set(app.SHIFT_EDITOR_OPTIONS) - {""})
        self.assertEqual(legend["休"], "通常の休み")
        self.assertEqual(legend["希"], "希望休")
        self.assertEqual(legend["有"], "有給休暇")
        for kind in ("日勤", "管", "夜勤", "夜勤明け", "休み", "希望休", "有休", "その他"):
            self.assertIn(kind, app.SHIFT_KINDS)

    def test_last_saved_time_updates_only_for_success_and_is_separated_by_month(self):
        state = {}
        saved_at = datetime(2026, 7, 20, 21, 35)
        saved = app.ShiftSaveResult(app.SHIFT_SAVE_SAVED, count=1)
        self.assertTrue(app.update_shift_last_saved_state(state, saved, 2026, 7, saved_at))
        self.assertEqual(state["shift_last_saved_at_2026_7"], "2026年7月20日 21:35")

        for status in (
            app.SHIFT_SAVE_NO_CHANGE,
            app.SHIFT_SAVE_DUPLICATE,
            app.SHIFT_SAVE_BLOCKED,
        ):
            result = app.ShiftSaveResult(status)
            self.assertFalse(
                app.update_shift_last_saved_state(
                    state, result, 2026, 7, datetime(2026, 7, 20, 22, 0)
                )
            )
        self.assertEqual(state["shift_last_saved_at_2026_7"], "2026年7月20日 21:35")

        self.assertTrue(
            app.update_shift_last_saved_state(
                state, saved, 2026, 8, datetime(2026, 8, 1, 9, 5)
            )
        )
        self.assertEqual(state["shift_last_saved_at_2026_8"], "2026年8月1日 09:05")
        self.assertEqual(state["shift_last_saved_at_2026_7"], "2026年7月20日 21:35")


class _TransactionCursor:
    def __init__(self):
        self.calls = 0
        self.closed = False

    def execute(self, _query, _params=()):
        self.calls += 1
        if self.calls == 2:
            raise RuntimeError("second operation failed")

    def executemany(self, _query, _params):
        self.execute(_query, _params)

    def close(self):
        self.closed = True


class _TransactionConnection:
    def __init__(self):
        self.cursor_value = _TransactionCursor()
        self.commits = 0
        self.rollbacks = 0
        self.closed = False

    def cursor(self):
        return self.cursor_value

    def commit(self):
        self.commits += 1

    def rollback(self):
        self.rollbacks += 1

    def close(self):
        self.closed = True


class TransactionTests(unittest.TestCase):
    def test_transaction_rolls_back_when_later_operation_fails(self):
        connection = _TransactionConnection()
        with patch.object(db, "get_conn", return_value=connection):
            with self.assertRaisesRegex(RuntimeError, "second operation failed"):
                db.execute_transaction([
                    ("DELETE FROM staff_shifts WHERE id=?", (1,)),
                    ("INSERT INTO staff_shifts(id) VALUES (?)", (2,)),
                ])

        self.assertEqual(connection.commits, 0)
        self.assertEqual(connection.rollbacks, 1)
        self.assertTrue(connection.cursor_value.closed)
        self.assertTrue(connection.closed)


if __name__ == "__main__":
    unittest.main()
