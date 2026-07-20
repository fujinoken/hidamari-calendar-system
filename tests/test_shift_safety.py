import functools
import inspect
import sys
import types
import unittest
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

    def test_confirmation_error_rules_separate_high_and_medium_checks(self):
        shortage = pd.DataFrame([{"日付": "2026-07-01", "不足": "夜勤あと1"}])
        quality = pd.DataFrame([
            {"重要度": "中", "種類": "5連勤", "日付": "2026-07-05", "職員名": "A", "内容": "注意"},
            {"重要度": "高", "種類": "6連勤", "日付": "2026-07-06", "職員名": "B", "内容": "要調整"},
        ])
        limits = pd.DataFrame([{"重要度": "中", "種類": "合計上限超過", "職員名": "C", "内容": "21/20"}])
        duplicates = pd.DataFrame([{"種類": "同一職員・同一日の重複", "日付": "2026-07-07", "職員名": "D", "内容": "2件"}])
        kot = pd.DataFrame([{"職員名": "E", "勤務日": "2026-07-08", "項目": "従業員コード", "内容": "未登録"}])

        errors = app.build_shift_confirmation_errors(shortage, quality, limits, duplicates, kot)

        self.assertEqual(len(errors), 5)
        self.assertNotIn("5連勤", errors["区分"].tolist())
        self.assertIn("6連勤", errors["区分"].tolist())
        self.assertIn("合計上限超過", errors["区分"].tolist())
        self.assertIn("KING OF TIME重大エラー", errors["区分"].tolist())

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
