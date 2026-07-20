# -*- coding: utf-8 -*-
"""Database access and small date/time helpers for ひだまり帳."""
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

try:
    import psycopg2
except ImportError:
    psycopg2 = None

from config import DEFAULT_CATEGORIES

DB_INTEGRITY_ERROR = psycopg2.IntegrityError if psycopg2 else Exception

def get_database_url():
    """
    PostgreSQL接続URLを取得する。
    Streamlit Cloudでは st.secrets["DATABASE_URL"] を推奨。
    ローカルでは環境変数 DATABASE_URL / SUPABASE_DB_URL でも動作する。
    """
    try:
        if "DATABASE_URL" in st.secrets:
            return st.secrets["DATABASE_URL"]
        if "SUPABASE_DB_URL" in st.secrets:
            return st.secrets["SUPABASE_DB_URL"]
        if "postgres" in st.secrets and "url" in st.secrets["postgres"]:
            return st.secrets["postgres"]["url"]
    except Exception:
        pass
    return os.getenv("DATABASE_URL") or os.getenv("SUPABASE_DB_URL")


def to_pg_query(query):
    """既存コードの ? プレースホルダを PostgreSQL/psycopg2 の %s へ変換する。"""
    return str(query).replace("?", "%s")


def get_conn():
    if psycopg2 is None:
        st.error("psycopg2 がインストールされていません。requirements.txt に psycopg2-binary を追加してください。")
        st.stop()

    database_url = get_database_url()
    if not database_url:
        st.error("PostgreSQL接続URLが未設定です。Streamlit secrets に DATABASE_URL を設定してください。")
        st.stop()

    # SupabaseではSSL必須のことが多いため、URLにsslmodeが無い場合はrequireを付ける。
    if "sslmode=" in database_url:
        return psycopg2.connect(database_url)
    return psycopg2.connect(database_url, sslmode=os.getenv("PGSSLMODE", "require"))


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        user_id TEXT UNIQUE,
        user_name TEXT NOT NULL UNIQUE,
        kana TEXT,
        room_no TEXT,
        note TEXT,
        is_active INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS staff (
        id SERIAL PRIMARY KEY,
        staff_name TEXT NOT NULL UNIQUE,
        staff_code TEXT,
        role TEXT,
        is_active INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS categories (
        id SERIAL PRIMARY KEY,
        category_name TEXT NOT NULL UNIQUE,
        mark TEXT,
        sort_order INTEGER DEFAULT 100,
        is_active INTEGER DEFAULT 1,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS events (
        id SERIAL PRIMARY KEY,
        event_date TEXT NOT NULL,
        category TEXT NOT NULL,
        title TEXT NOT NULL,
        user_id TEXT,
        user_name TEXT,
        staff_name TEXT,
        start_time TEXT,
        end_time TEXT,
        memo TEXT,
        important INTEGER DEFAULT 0,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS event_photos (
        id SERIAL PRIMARY KEY,
        event_id INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
        file_name TEXT NOT NULL,
        file_path TEXT NOT NULL,
        photo_memo TEXT,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS event_files (
        id SERIAL PRIMARY KEY,
        event_id INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
        file_name TEXT NOT NULL,
        file_path TEXT NOT NULL,
        file_type TEXT,
        file_memo TEXT,
        created_at TEXT NOT NULL
    )
    """)

    # 既存PostgreSQLテーブルからの移行：列がなければ追加する
    cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS user_id TEXT")
    cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS room_no TEXT")
    cur.execute("ALTER TABLE events ADD COLUMN IF NOT EXISTS user_id TEXT")
    cur.execute("ALTER TABLE staff ADD COLUMN IF NOT EXISTS staff_code TEXT")

    # 既存利用者にIDがない場合、U0001形式で仮IDを付与
    cur.execute("SELECT id, user_id FROM users")
    for row_id, user_id in cur.fetchall():
        if not user_id:
            cur.execute("UPDATE users SET user_id=%s WHERE id=%s", (f"U{int(row_id):04d}", int(row_id)))

    # 既存予定にIDがない場合、利用者名から補完
    cur.execute("""
        UPDATE events
        SET user_id = (
            SELECT users.user_id FROM users
            WHERE users.user_name = events.user_name
            LIMIT 1
        )
        WHERE (user_id IS NULL OR user_id = '') AND user_name IS NOT NULL
    """)

    # カテゴリマスタ初期投入
    cur.execute("SELECT COUNT(*) FROM categories")
    category_count = cur.fetchone()[0]
    if category_count == 0:
        default_marks = {
            "通院": "🏥",
            "面会": "👪",
            "行事": "🎉",
            "外出": "🚶",
            "注意": "⚠️",
            "申し送り": "📝",
            "面接": "・",
            "その他": "・",
        }
        for i, name in enumerate(DEFAULT_CATEGORIES, start=1):
            cur.execute("""
                INSERT INTO categories
                (category_name, mark, sort_order, is_active, created_at, updated_at)
                VALUES (%s, %s, %s, 1, %s, %s)
                ON CONFLICT (category_name) DO NOTHING
            """, (name, default_marks.get(name, "・"), i * 10, now_text(), now_text()))

    cur.execute("""
    CREATE TABLE IF NOT EXISTS staff_shifts (
        id SERIAL PRIMARY KEY,
        shift_date TEXT NOT NULL,
        staff_name TEXT NOT NULL,
        shift_kind TEXT NOT NULL,
        start_time TEXT,
        end_time TEXT,
        next_day INTEGER DEFAULT 0,
        memo TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS shift_month_status (
        id SERIAL PRIMARY KEY,
        shift_year INTEGER NOT NULL,
        shift_month INTEGER NOT NULL,
        is_confirmed INTEGER DEFAULT 0,
        confirmed_at TEXT,
        confirmed_by TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS staff_shift_limits (
        id SERIAL PRIMARY KEY,
        staff_name TEXT NOT NULL UNIQUE,
        max_day_shifts INTEGER DEFAULT 31,
        max_night_shifts INTEGER DEFAULT 31,
        max_total_shifts INTEGER DEFAULT 31,
        note TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS kot_auto_schedule_patterns (
        id SERIAL PRIMARY KEY,
        shift_kind TEXT NOT NULL UNIQUE,
        pattern_code TEXT,
        pattern_name TEXT,
        day_type_code TEXT DEFAULT '1',
        day_type_name TEXT DEFAULT '平日',
        leave_name TEXT,
        is_active INTEGER DEFAULT 1,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS kot_auto_schedule_settings (
        setting_key TEXT PRIMARY KEY,
        setting_value TEXT,
        updated_at TEXT NOT NULL
    )
    """)

    for shift_kind in ("日勤", "夜勤", "管"):
        cur.execute("""
            INSERT INTO kot_auto_schedule_patterns
            (shift_kind, pattern_code, pattern_name, day_type_code, day_type_name, leave_name, is_active, created_at, updated_at)
            VALUES (%s, '', '', '1', '平日', '', 1, %s, %s)
            ON CONFLICT (shift_kind) DO NOTHING
        """, (shift_kind, now_text(), now_text()))

    default_kot_settings = {
        "rest_day_type_code": "3",
        "rest_day_type_name": "法定外休日",
        "rest_leave_name": "公休",
        "paid_day_type_code": "1",
        "paid_day_type_name": "平日",
        "paid_leave_name": "有休",
        "statutory_weekday": "",
        "holiday_day_type_code": "1",
    }
    for setting_key, setting_value in default_kot_settings.items():
        cur.execute("""
            INSERT INTO kot_auto_schedule_settings (setting_key, setting_value, updated_at)
            VALUES (%s, %s, %s)
            ON CONFLICT (setting_key) DO NOTHING
        """, (setting_key, setting_value, now_text()))

    # よく使う検索用インデックス。既存DBにも安全に追加できる。
    cur.execute("CREATE INDEX IF NOT EXISTS idx_events_event_date ON events(event_date)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_events_date_start ON events(event_date, start_time)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_events_user_name ON events(user_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_event_photos_event_id ON event_photos(event_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_event_files_event_id ON event_files(event_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_categories_active_sort ON categories(is_active, sort_order)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_users_active_name ON users(is_active, user_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_staff_active_name ON staff(is_active, staff_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_staff_code ON staff(staff_code)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_staff_shifts_date ON staff_shifts(shift_date)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_staff_shifts_staff ON staff_shifts(staff_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_staff_shifts_date_kind ON staff_shifts(shift_date, shift_kind)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_shift_month_status_ym ON shift_month_status(shift_year, shift_month)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_staff_shift_limits_staff ON staff_shift_limits(staff_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_kot_auto_patterns_active ON kot_auto_schedule_patterns(is_active, shift_kind)")

    conn.commit()
    cur.close()
    conn.close()


JST = ZoneInfo("Asia/Tokyo")


def now_text():
    """
    Streamlit Cloud上でも日本時間で保存する。
    """
    return datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")


def today_jst():
    return datetime.now(JST).date()


def normalize_query_params(params=()):
    """cache用にSQLパラメータをtupleへ正規化する。"""
    if params is None:
        return ()
    if isinstance(params, tuple):
        return params
    if isinstance(params, list):
        return tuple(params)
    return (params,)


@st.cache_data(ttl=20, show_spinner=False)
def _fetch_df_cached(query, params_tuple=()):
    """
    読み取りSQLを短時間キャッシュする。
    予定表示・マスタ取得・件数確認の体感速度を上げる。
    """
    conn = get_conn()
    try:
        df = pd.read_sql_query(to_pg_query(query), conn, params=params_tuple)
    finally:
        conn.close()
    return df


def fetch_df(query, params=()):
    return _fetch_df_cached(str(query), normalize_query_params(params))


def clear_read_cache():
    """DB更新後に読み取りキャッシュをクリアする。"""
    try:
        _fetch_df_cached.clear()
    except Exception:
        pass


def execute(query, params=()):
    """
    INSERT/UPDATE/DELETEを実行する。
    INSERTの場合は自動で RETURNING id を付け、登録IDを返す。
    """
    conn = get_conn()
    cur = conn.cursor()
    q = to_pg_query(query).strip()
    q_no_semicolon = q[:-1].strip() if q.endswith(";") else q
    is_insert = q_no_semicolon.lower().startswith("insert")
    if is_insert and " returning " not in q_no_semicolon.lower():
        q_exec = q_no_semicolon + " RETURNING id"
    else:
        q_exec = q_no_semicolon

    try:
        cur.execute(q_exec, normalize_query_params(params))
        last_id = None
        if is_insert:
            row = cur.fetchone()
            last_id = row[0] if row else None
        conn.commit()
        clear_read_cache()
        return last_id
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


def execute_many(query, params_list):
    """複数行のINSERT/UPDATEをまとめて実行する。"""
    if not params_list:
        return 0
    conn = get_conn()
    cur = conn.cursor()
    q = to_pg_query(query).strip()
    try:
        cur.executemany(q, [normalize_query_params(p) for p in params_list])
        conn.commit()
        clear_read_cache()
        return len(params_list)
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


def execute_transaction(operations):
    """複数の更新SQLを同一トランザクションで実行する。

    operations は (query, params) または
    (query, params_or_params_list, use_executemany) の配列を受け取る。
    """
    if not operations:
        return 0
    conn = get_conn()
    cur = conn.cursor()
    executed = 0
    try:
        for operation in operations:
            if len(operation) == 2:
                query, params = operation
                use_executemany = False
            else:
                query, params, use_executemany = operation
            q = to_pg_query(query).strip()
            if use_executemany:
                normalized_rows = [normalize_query_params(row) for row in params]
                if normalized_rows:
                    cur.executemany(q, normalized_rows)
                    executed += len(normalized_rows)
            else:
                cur.execute(q, normalize_query_params(params))
                executed += 1
        conn.commit()
        clear_read_cache()
        return executed
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


def init_db_once():
    """
    init_dbはALTER TABLEや初期投入を含むため、毎回実行せずセッション内で1回だけ実行する。
    これだけで画面切替・再描画が1テンポ軽くなる。
    """
    if not st.session_state.get("_hidamari_db_initialized", False):
        init_db()
        st.session_state["_hidamari_db_initialized"] = True
        clear_read_cache()



