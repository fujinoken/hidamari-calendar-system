import functools
import inspect
import sys
import types
import unittest
from datetime import date
from unittest.mock import patch

import pandas as pd


class _CacheData:
    def __call__(self, func=None, **_kwargs):
        if func is None:
            return lambda target: self(target)
        cache = {}

        @functools.wraps(func)
        def wrapped(*args, **kwargs):
            bound = inspect.signature(func).bind(*args, **kwargs)
            key = repr(tuple((name, value) for name, value in bound.arguments.items() if not name.startswith("_")))
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


def _legacy_shift_day_labels(df, staff_name, target_date):
    labels = []
    if df is None or df.empty:
        return labels
    target_staff = app.normalize_staff_name(staff_name)
    staff_df = df.copy()
    staff_df["_staff_norm"] = staff_df["staff_name"].apply(app.normalize_staff_name)
    staff_df = staff_df[staff_df["_staff_norm"] == target_staff]
    day_rows = staff_df[staff_df["shift_date"].astype(str) == str(target_date)]
    day_shift_kinds = []
    for _, row in day_rows.iterrows():
        kind = str(row["shift_kind"])
        day_shift_kinds.append(kind)
        label = app.shift_short_label(kind)
        if label and label not in labels:
            labels.append(label)
    has_off_label = (
        any(value in labels for value in ["希", "有", "他"])
        or any(kind in app.OFF_OR_BLOCKING_SHIFT_KINDS for kind in day_shift_kinds)
    )
    try:
        previous_date = (
            pd.to_datetime(str(target_date)).date() - pd.Timedelta(days=1)
        ).strftime("%Y-%m-%d")
        previous_rows = staff_df[
            (staff_df["shift_date"].astype(str) == previous_date)
            & (staff_df["shift_kind"].astype(str) == "夜勤")
        ]
        if not previous_rows.empty and "明" not in labels and not has_off_label:
            labels.insert(0, "明")
    except Exception:
        pass
    if has_off_label and "明" in labels:
        labels = [value for value in labels if value != "明"]
    return labels


def _legacy_shift_day_actual_label(df, staff_name, target_date):
    labels = []
    staff_df = df.copy()
    staff_df["_staff_norm"] = staff_df["staff_name"].apply(app.normalize_staff_name)
    day_rows = staff_df[
        (staff_df["_staff_norm"] == app.normalize_staff_name(staff_name))
        & (staff_df["shift_date"].astype(str) == str(target_date))
    ]
    for _, row in day_rows.iterrows():
        label = app.shift_short_label(str(row.get("shift_kind") or ""))
        if label and label not in labels:
            labels.append(label)
    return app.labels_to_cell_value(labels)


class SpeedImprovementTests(unittest.TestCase):
    def test_page_builds_shift_snapshot_once(self):
        source = inspect.getsource(app.page_shift_manager)
        self.assertEqual(source.count("build_shift_read_index("), 1)
        self.assertNotIn("_build_shift_report_snapshot(", source)
        self.assertIn('shift_snapshot = shift_read_index["report_snapshot"]', source)

    def test_shift_read_index_handles_empty_and_normalizes_each_row_once(self):
        empty = app.build_shift_read_index(pd.DataFrame())
        self.assertEqual(empty["normalized_staff_names"], ())
        self.assertEqual(empty["kinds_by_staff_date"], {})

        shifts = pd.DataFrame([
            {"shift_date": "2026-06-30", "staff_name": " 藤野 ", "shift_kind": "夜勤"},
            {"shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤"},
            {"shift_date": "2026-07-01", "staff_name": "池田", "shift_kind": "日勤"},
            {"shift_date": pd.NaT, "staff_name": None, "shift_kind": float("nan")},
        ])
        with patch.object(app, "normalize_staff_name", wraps=app.normalize_staff_name) as normalize:
            index = app.build_shift_read_index(shifts)
        self.assertEqual(normalize.call_count, len(shifts))
        self.assertEqual(index["normalized_staff_names"], ("池田", "藤野"))
        self.assertEqual(index["kinds_by_staff_date"]["藤野"]["2026-06-30"], ("夜勤",))
        self.assertNotIn("nan", index["normalized_staff_names"])

    def test_shift_read_index_is_stable_for_row_order_and_multiple_rows(self):
        shifts = pd.DataFrame([
            {"id": 3, "shift_date": "2026-07-02", "staff_name": "藤野", "shift_kind": "希望休"},
            {"id": 2, "shift_date": "2026-07-01", "staff_name": "池田", "shift_kind": "夜勤"},
            {"id": 1, "shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤"},
            {"id": 4, "shift_date": "2026-07-02", "staff_name": "藤野", "shift_kind": "その他"},
        ])
        first = app.build_shift_read_index(shifts, 2026, 7)
        second = app.build_shift_read_index(shifts.iloc[::-1].reset_index(drop=True), 2026, 7)
        self.assertEqual(first, second)
        self.assertEqual(first["kinds_by_staff_date"]["藤野"]["2026-07-02"], ("その他", "希望休"))

    def test_indexed_day_labels_and_matrices_match_legacy_results(self):
        shifts = pd.DataFrame([
            {"shift_date": "2026-06-30", "staff_name": "藤野", "shift_kind": "夜勤"},
            {"shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤"},
            {"shift_date": "2026-07-01", "staff_name": "池田", "shift_kind": "日勤"},
            {"shift_date": "2026-07-02", "staff_name": "池田", "shift_kind": "夜勤"},
            {"shift_date": "2026-07-03", "staff_name": "池田", "shift_kind": "希望休"},
        ])
        index = app.build_shift_read_index(shifts)
        for staff_name in ("藤野", "池田"):
            for target_date in ("2026-07-01", "2026-07-02", "2026-07-03"):
                self.assertEqual(
                    app.shift_day_labels_for_staff(shifts, staff_name, target_date, index),
                    _legacy_shift_day_labels(shifts, staff_name, target_date),
                )
                self.assertEqual(
                    app.shift_day_actual_label_for_staff(shifts, staff_name, target_date, index),
                    _legacy_shift_day_actual_label(shifts, staff_name, target_date),
                )
        pd.testing.assert_frame_equal(
            app.create_shift_matrix(shifts, 2026, 7, index),
            app.create_shift_matrix(shifts, 2026, 7),
        )
        pd.testing.assert_frame_equal(
            app.create_editable_shift_matrix(["藤野", "池田"], shifts, 2026, 7, index),
            app.create_editable_shift_matrix(["藤野", "池田"], shifts, 2026, 7),
        )
        shortage = app.create_shift_shortage_table(shifts, 2026, 7, index)
        self.assertEqual(shortage.iloc[0]["日勤人数"], 2)
        self.assertEqual(shortage.iloc[0]["夜勤人数"], 0)
        self.assertEqual(shortage.iloc[0]["状態"], "要確認")

    def test_quality_check_reuses_supplied_matrix_and_index(self):
        shifts = pd.DataFrame([
            {"shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤"},
            {"shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "希望休"},
        ])
        index = app.build_shift_read_index(shifts)
        matrix = app.create_shift_matrix(shifts, 2026, 7, index)
        with patch.object(app, "create_shift_matrix", side_effect=AssertionError("再生成")):
            result = app.create_shift_quality_check_table(shifts, 2026, 7, index, matrix)
        expected = pd.DataFrame([{
            "重要度": "高",
            "種類": "休み希望と勤務の重複",
            "日付": "2026-07-01",
            "職員名": "藤野",
            "内容": "休み・希望休・有休・その他と、日勤・夜勤・明けが同日に入っています。",
        }])
        pd.testing.assert_frame_equal(result, expected)

    def test_staff_filter_and_hot_loops_do_not_rescan_dataframe(self):
        entries = ((" 藤野 ", "藤野"), ("池田", "池田"), ("藤野", "藤野"))
        self.assertEqual(app._unique_index_staff_names(entries), ["藤野", "池田"])
        self.assertEqual(app._unique_index_staff_names(entries, "池田"), ["池田"])
        for target in (
            app.shift_day_labels_for_staff,
            app.shift_day_actual_label_for_staff,
            app.create_shift_matrix,
            app.create_editable_shift_matrix,
            app.render_shift_calendar,
        ):
            source = inspect.getsource(target)
            self.assertNotIn(".copy()", source)
            self.assertNotIn(".apply(normalize_staff_name)", source)
        calendar_source = inspect.getsource(app.render_shift_calendar)
        self.assertNotIn('shift_df[', calendar_source)
        self.assertNotIn("astype(str)", calendar_source)

    def test_edit_button_keeps_selected_date_without_explicit_rerun(self):
        source = inspect.getsource(app.render_shift_calendar)
        self.assertIn('st.session_state["selected_shift_date"]', source)
        self.assertNotIn("st.rerun()", source)

    def test_monthly_event_counts_execute_once_and_use_half_open_month_range(self):
        result_df = pd.DataFrame([
            {"event_date": "2026-07-01", "staff_name": "藤野", "event_count": 2},
            {"event_date": "2026-07-01", "staff_name": "池田", "event_count": 1},
            {"event_date": "2026-07-02", "staff_name": "藤野", "event_count": 1},
        ])
        app.get_staff_event_counts_month.clear()
        with patch.object(app, "fetch_df", return_value=result_df) as fetch:
            first = app.get_staff_event_counts_month(2026, 7)
            second = app.get_staff_event_counts_month(2026, 7)

        self.assertEqual(first, second)
        self.assertEqual(first["2026-07-01"], {"藤野": 2, "池田": 1})
        self.assertEqual(fetch.call_count, 1)
        sql, params = fetch.call_args.args
        self.assertIn("staff_name IS NOT NULL", sql)
        self.assertIn("staff_name <> ''", sql)
        self.assertEqual(params, ("2026-07-01", "2026-08-01"))

    def test_monthly_event_counts_exclude_next_year_january(self):
        app.get_staff_event_counts_month.clear()
        with patch.object(app, "fetch_df", return_value=pd.DataFrame()) as fetch:
            app.get_staff_event_counts_month(2026, 12)
        self.assertEqual(fetch.call_args.args[1], ("2026-12-01", "2027-01-01"))

    def test_candidate_order_matches_existing_daily_count_path(self):
        event = pd.Series({"event_date": "2026-07-03", "start_time": "10:00"})
        shifts = pd.DataFrame([
            {"shift_date": "2026-07-03", "staff_name": "藤野", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
            {"shift_date": "2026-07-03", "staff_name": "池田", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
        ])
        counts = {"藤野": 2, "池田": 1}
        expected = app._get_shift_candidates_for_event(event, shifts, counts)
        with patch.object(app, "get_staff_event_counts_for_date", return_value=counts):
            actual = app.get_shift_candidates_for_event(event, shifts)
        pd.testing.assert_frame_equal(actual, expected)
        self.assertEqual(actual.iloc[0]["職員名"], "池田")

    def test_assignment_preview_gets_month_counts_once_for_multiple_events(self):
        events = pd.DataFrame([
            {"id": 1, "event_date": "2026-07-01", "title": "A", "staff_name": "", "start_time": "10:00"},
            {"id": 2, "event_date": "2026-07-01", "title": "B", "staff_name": "", "start_time": "11:00"},
            {"id": 3, "event_date": "2026-07-02", "title": "C", "staff_name": "", "start_time": "12:00"},
        ])
        shifts = pd.DataFrame([
            {"shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
            {"shift_date": "2026-07-02", "staff_name": "藤野", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
        ])
        with patch.object(app, "get_staff_event_counts_month", return_value={"2026-07-01": {"藤野": 2}}) as get_counts:
            preview = app.build_event_assignment_preview(events, shifts)
        self.assertEqual(get_counts.call_count, 1)
        self.assertEqual(len(preview), 3)

    def test_assignment_preview_loads_each_month_once_and_uses_matching_counts(self):
        events = pd.DataFrame([
            {"id": 1, "event_date": "2026-12-31", "title": "年末A", "staff_name": "", "start_time": "10:00"},
            {"id": 2, "event_date": "2026-12-31", "title": "年末B", "staff_name": "", "start_time": "11:00"},
            {"id": 3, "event_date": "2027-01-01", "title": "年始", "staff_name": "", "start_time": "10:00"},
            {"id": 4, "event_date": None, "title": "日付なし", "staff_name": "", "start_time": "10:00"},
        ])
        shifts = pd.DataFrame([
            {"shift_date": "2026-12-31", "staff_name": "藤野", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
            {"shift_date": "2026-12-31", "staff_name": "池田", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
            {"shift_date": "2027-01-01", "staff_name": "藤野", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
            {"shift_date": "2027-01-01", "staff_name": "池田", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
        ])

        def monthly_counts(year, month):
            if (year, month) == (2026, 12):
                return {"2026-12-31": {"藤野": 2, "池田": 0}}
            return {"2027-01-01": {"藤野": 0, "池田": 2}}

        with patch.object(app, "get_staff_event_counts_month", side_effect=monthly_counts) as get_counts:
            preview = app.build_event_assignment_preview(events, shifts)

        self.assertEqual(get_counts.call_count, 2)
        self.assertEqual({call.args for call in get_counts.call_args_list}, {(2026, 12), (2027, 1)})
        self.assertEqual(preview.loc[preview["予定ID"] == 1, "AI候補"].iloc[0], "池田")
        self.assertEqual(preview.loc[preview["予定ID"] == 3, "AI候補"].iloc[0], "藤野")

    def test_cached_kot_export_runs_once_for_same_key(self):
        app._cached_king_of_time_clock_export.clear()
        generated = (pd.DataFrame(), pd.DataFrame(), b"csv")
        with patch.object(app, "build_king_of_time_clock_export", return_value=generated) as build:
            first = app._cached_king_of_time_clock_export(2026, 7, (1,), "same", "1")
            second = app._cached_king_of_time_clock_export(2026, 7, (1,), "same", "1")
            app._cached_king_of_time_clock_export(2026, 7, (1,), "same", "2")
        self.assertEqual(first[2], second[2])
        self.assertEqual(build.call_count, 2)

    def test_pdf_and_excel_cache_reuse_same_key_and_refresh_changed_key(self):
        app._cached_staff_shift_pdf.clear()
        app._cached_staff_shift_excel.clear()
        with patch.object(app, "make_staff_shift_pdf", return_value=b"pdf") as make_pdf:
            self.assertEqual(app._cached_staff_shift_pdf(2026, 7, "same", "1"), b"pdf")
            self.assertEqual(app._cached_staff_shift_pdf(2026, 7, "same", "1"), b"pdf")
            app._cached_staff_shift_pdf(2026, 7, "same", "2")
        with patch.object(app, "make_staff_shift_excel", return_value=b"xlsx") as make_excel:
            self.assertEqual(app._cached_staff_shift_excel(2026, 7, "same", "1"), b"xlsx")
            self.assertEqual(app._cached_staff_shift_excel(2026, 7, "same", "1"), b"xlsx")
            app._cached_staff_shift_excel(2026, 7, "same", "2")
        self.assertEqual(make_pdf.call_count, 2)
        self.assertEqual(make_excel.call_count, 2)

    def test_calendar_pdf_cache_key_includes_staff_selection_and_finalized_state(self):
        app._cached_shift_calendar_pdf.clear()
        rows = (("2026-07-01", "藤野", "日勤"),)
        with patch.object(app, "make_shift_calendar_pdf", return_value=b"calendar") as make_calendar:
            kwargs = {"_shift_rows": rows, "staff_list": ("藤野", "池田"), "selected_staff_names": (), "finalized": False}
            app._cached_shift_calendar_pdf(2026, 7, "hash", "1", **kwargs)
            app._cached_shift_calendar_pdf(2026, 7, "hash", "1", **kwargs)
            app._cached_shift_calendar_pdf(2026, 7, "hash", "1", **{**kwargs, "selected_staff_names": ("藤野",)})
            app._cached_shift_calendar_pdf(2026, 7, "hash", "1", **{**kwargs, "selected_staff_names": ("藤野",), "finalized": True})
            app._cached_shift_calendar_pdf(2026, 7, "hash", "2", **kwargs)
        self.assertEqual(make_calendar.call_count, 4)

    def test_report_signature_changes_with_shift_status_and_limit(self):
        shifts = pd.DataFrame([{"id": 1, "shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤"}])
        limits = pd.DataFrame([{"職員名": "藤野", "日勤上限": 10, "夜勤上限": 5, "合計上限": 15, "メモ": ""}])
        snapshot = app._build_shift_report_snapshot(shifts, 2026, 7)
        base = app._shift_report_signature(snapshot, {"is_confirmed": 0}, limits)
        changed_shift = shifts.copy()
        changed_shift.loc[0, "shift_kind"] = "夜勤"
        changed_limit = limits.copy()
        changed_limit.loc[0, "日勤上限"] = 9
        changed_snapshot = app._build_shift_report_snapshot(changed_shift, 2026, 7)
        self.assertNotEqual(base, app._shift_report_signature(changed_snapshot, {"is_confirmed": 0}, limits))
        self.assertNotEqual(base, app._shift_report_signature(snapshot, {"is_confirmed": 1}, limits))
        self.assertNotEqual(base, app._shift_report_signature(snapshot, {"is_confirmed": 0}, changed_limit))

    def test_shift_snapshot_is_order_independent_and_normalizes_scalar_types(self):
        first = pd.DataFrame([
            {"id": 2, "shift_date": pd.Timestamp("2026-07-02"), "staff_name": "池田", "shift_kind": "夜勤", "start_time": None, "end_time": "09:30", "next_day": 1.0},
            {"id": 1.0, "shift_date": date(2026, 7, 1), "staff_name": "藤野", "shift_kind": "日勤", "start_time": float("nan"), "end_time": "17:30", "next_day": 0},
        ])
        second = pd.DataFrame([
            {"id": 1, "shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤", "start_time": None, "end_time": "17:30", "next_day": 0.0},
            {"id": 2.0, "shift_date": date(2026, 7, 2), "staff_name": "池田", "shift_kind": "夜勤", "start_time": float("nan"), "end_time": "09:30", "next_day": 1},
        ])
        first_snapshot = app._build_shift_report_snapshot(first, 2026, 7)
        second_snapshot = app._build_shift_report_snapshot(second, 2026, 7)
        self.assertEqual(first_snapshot, second_snapshot)

    def test_previous_month_does_not_change_calendar_or_kot_month_hash(self):
        current = pd.DataFrame([
            {"id": 1, "shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
        ])
        with_previous = pd.concat([current, pd.DataFrame([
            {"id": 9, "shift_date": "2026-06-30", "staff_name": "池田", "shift_kind": "夜勤", "start_time": "16:30", "end_time": "09:30", "next_day": 1},
        ])], ignore_index=True)
        current_snapshot = app._build_shift_report_snapshot(current, 2026, 7)
        previous_snapshot = app._build_shift_report_snapshot(with_previous, 2026, 7)
        self.assertEqual(current_snapshot["base_hash"], previous_snapshot["base_hash"])
        self.assertEqual(current_snapshot["kot_detail_hash"], previous_snapshot["kot_detail_hash"])
        self.assertNotEqual(current_snapshot["previous_context_hash"], previous_snapshot["previous_context_hash"])

    def test_kot_detail_changes_do_not_change_shared_base_hash(self):
        base = pd.DataFrame([
            {"id": 1, "shift_date": "2026-07-01", "staff_name": "藤野", "shift_kind": "その他", "start_time": "08:30", "end_time": "17:30", "next_day": 0, "memo": "A", "updated_at": "old"},
        ])
        changed = base.copy()
        changed.loc[0, "start_time"] = "09:00"
        changed.loc[0, "memo"] = "B"
        changed.loc[0, "updated_at"] = "new"
        base_snapshot = app._build_shift_report_snapshot(base, 2026, 7)
        changed_snapshot = app._build_shift_report_snapshot(changed, 2026, 7)
        self.assertEqual(base_snapshot["base_hash"], changed_snapshot["base_hash"])
        self.assertNotEqual(base_snapshot["kot_detail_hash"], changed_snapshot["kot_detail_hash"])

    def test_kot_csv_remains_headerless_selected_staff_only_and_handles_overnight(self):
        shifts = pd.DataFrame([
            {"id": 1, "shift_date": "2026-07-02", "staff_name": "藤野", "shift_kind": "夜勤", "start_time": "16:30", "end_time": "09:30", "next_day": 1},
            {"id": 2, "shift_date": "2026-07-02", "staff_name": "池田", "shift_kind": "日勤", "start_time": "08:30", "end_time": "17:30", "next_day": 0},
        ])
        preview, errors, csv_bytes = app.report_build_king_of_time_clock_export(
            2026,
            7,
            lambda _start, _end: shifts,
            lambda active_only=False: {"藤野": "1001", "池田": "1002"},
            lambda value: str(value).strip(),
            default_shift_times=app.default_shift_times,
            selected_staff_keys=[1],
            get_staff_key_map=lambda active_only=False: {"藤野": 1, "池田": 2},
        )
        text = csv_bytes.decode("utf-8-sig")
        self.assertTrue(errors.empty)
        self.assertEqual(preview["職員名"].tolist(), ["藤野"])
        self.assertNotIn("従業員コード", text)
        self.assertNotIn("池田", text)
        self.assertIn("202607030930", text)


if __name__ == "__main__":
    unittest.main()
