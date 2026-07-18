import unittest
from unittest.mock import patch

import pandas as pd

import app
import db


class _FakeCursor:
    def __init__(self):
        self.queries = []
        self._result = None

    def execute(self, query, params=None):
        normalized = " ".join(str(query).split())
        self.queries.append((normalized, params))
        if "SELECT id, user_id FROM users" in normalized:
            self._result = []
        elif "SELECT COUNT(*) FROM categories" in normalized:
            self._result = (1,)
        else:
            self._result = None

    def fetchall(self):
        return list(self._result or [])

    def fetchone(self):
        return self._result

    def close(self):
        pass


class _FakeConnection:
    def __init__(self):
        self.cursor_instance = _FakeCursor()
        self.commit_count = 0

    def cursor(self):
        return self.cursor_instance

    def commit(self):
        self.commit_count += 1

    def close(self):
        pass


class KingOfTimeDatabaseTests(unittest.TestCase):
    def test_database_initialization_runs_once_per_session(self):
        app.st.session_state.pop("_hidamari_db_initialized", None)
        with patch.object(db, "init_db") as mocked_init:
            app.init_db_once()
            app.init_db_once()
        self.assertEqual(mocked_init.call_count, 1)
        app.st.session_state.pop("_hidamari_db_initialized", None)

    def test_setting_table_creation_is_idempotent(self):
        first = _FakeConnection()
        second = _FakeConnection()
        with patch.object(db, "get_conn", side_effect=[first, second]):
            db.init_db()
            db.init_db()
        for connection in (first, second):
            sql = "\n".join(query for query, _ in connection.cursor_instance.queries)
            self.assertIn("CREATE TABLE IF NOT EXISTS kot_auto_schedule_patterns", sql)
            self.assertIn("CREATE TABLE IF NOT EXISTS kot_auto_schedule_settings", sql)
            self.assertIn("ON CONFLICT (shift_kind) DO NOTHING", sql)
            self.assertIn("ON CONFLICT (setting_key) DO NOTHING", sql)
            self.assertEqual(connection.commit_count, 1)

    def test_saved_settings_can_be_read_back(self):
        pattern_state = {
            "日勤": {
                "pattern_code": "", "pattern_name": "", "day_type_code": "1",
                "day_type_name": "平日", "leave_name": "", "is_active": 1,
            }
        }
        setting_state = {"rest_leave_name": "公休"}

        def fake_execute(query, params=()):
            if "UPDATE kot_auto_schedule_patterns" in query:
                code, name, day_code, day_name, leave, _updated, shift_kind = params
                pattern_state[shift_kind] = {
                    "pattern_code": code, "pattern_name": name,
                    "day_type_code": day_code, "day_type_name": day_name,
                    "leave_name": leave, "is_active": 1,
                }
            elif "UPDATE kot_auto_schedule_settings" in query:
                value, _updated, key = params
                setting_state[key] = value

        def fake_fetch(query, params=()):
            if "FROM kot_auto_schedule_patterns" in query:
                return pd.DataFrame([
                    {"shift_kind": key, **value} for key, value in pattern_state.items()
                ])
            if "FROM kot_auto_schedule_settings" in query:
                return pd.DataFrame([
                    {"setting_key": key, "setting_value": value}
                    for key, value in setting_state.items()
                ])
            return pd.DataFrame()

        pattern_values = {
            "日勤": {
                "pattern_code": "DAY001", "pattern_name": "日勤",
                "day_type_code": "1", "day_type_name": "平日", "leave_name": "",
            }
        }
        setting_values = {"rest_leave_name": "所定公休"}
        with patch.object(app, "execute", side_effect=fake_execute), patch.object(app, "fetch_df", side_effect=fake_fetch):
            app.save_kot_auto_schedule_settings(pattern_values, setting_values)
            reloaded_patterns = app.get_kot_auto_schedule_patterns()
            reloaded_settings = app.get_kot_auto_schedule_settings()

        self.assertEqual(reloaded_patterns["日勤"]["pattern_code"], "DAY001")
        self.assertEqual(reloaded_settings["rest_leave_name"], "所定公休")


if __name__ == "__main__":
    unittest.main()
