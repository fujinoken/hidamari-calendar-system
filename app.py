
# -*- coding: utf-8 -*-
"""
ひだまり帳 Ver1.2.0
超軽量・単独版
Python + Streamlit + SQLite

起動:
    streamlit run app.py

必要ライブラリ:
    pip install streamlit pandas openpyxl
"""

import sqlite3
import calendar
import re
import hashlib
from datetime import date, datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd
import streamlit as st


APP_TITLE = "ひだまり帳 Ver1.2.0"
DB_PATH = Path("hidamari_calendar.db")
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)
FILE_DIR = Path("attached_files")
FILE_DIR.mkdir(exist_ok=True)


# -----------------------------
# DB
# -----------------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        staff_name TEXT NOT NULL UNIQUE,
        role TEXT,
        is_active INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event_id INTEGER NOT NULL,
        file_name TEXT NOT NULL,
        file_path TEXT NOT NULL,
        photo_memo TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(event_id) REFERENCES events(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS event_files (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event_id INTEGER NOT NULL,
        file_name TEXT NOT NULL,
        file_path TEXT NOT NULL,
        file_type TEXT,
        file_memo TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(event_id) REFERENCES events(id) ON DELETE CASCADE
    )
    """)

    # 既存DBからの移行：列がなければ追加する
    def ensure_column(table_name, column_name, column_def):
        cur.execute(f"PRAGMA table_info({table_name})")
        cols = [r[1] for r in cur.fetchall()]
        if column_name not in cols:
            cur.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_def}")

    ensure_column("users", "user_id", "TEXT")
    ensure_column("users", "room_no", "TEXT")
    ensure_column("events", "user_id", "TEXT")

    # 既存利用者にIDがない場合、U0001形式で仮IDを付与
    cur.execute("SELECT id, user_id FROM users")
    for r in cur.fetchall():
        if not r[1]:
            cur.execute("UPDATE users SET user_id=? WHERE id=?", (f"U{int(r[0]):04d}", int(r[0])))

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
            "夜勤": "🌙",
            "その他": "・",
        }
        for i, name in enumerate(DEFAULT_CATEGORIES, start=1):
            cur.execute("""
                INSERT INTO categories
                (category_name, mark, sort_order, is_active, created_at, updated_at)
                VALUES (?, ?, ?, 1, ?, ?)
            """, (name, default_marks.get(name, "・"), i * 10, now_text(), now_text()))

    conn.commit()
    conn.close()


JST = ZoneInfo("Asia/Tokyo")


def now_text():
    """
    Streamlit Cloud上でも日本時間で保存する。
    """
    return datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")


def today_jst():
    return datetime.now(JST).date()


def fetch_df(query, params=()):
    conn = get_conn()
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def execute(query, params=()):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(query, params)
    conn.commit()
    last_id = cur.lastrowid
    conn.close()
    return last_id




def save_uploaded_photos(event_id, uploaded_files, photo_memo=""):
    """アップロード写真をuploadsフォルダに保存し、DBへ紐づける。"""
    if not uploaded_files:
        return

    for uploaded in uploaded_files:
        suffix = Path(uploaded.name).suffix.lower()
        safe_name = f"event_{event_id}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}{suffix}"
        save_path = UPLOAD_DIR / safe_name

        with open(save_path, "wb") as f:
            f.write(uploaded.getbuffer())

        execute("""
            INSERT INTO event_photos
            (event_id, file_name, file_path, photo_memo, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (
            int(event_id),
            uploaded.name,
            str(save_path),
            photo_memo.strip() or None,
            now_text(),
        ))


def get_event_photos(event_id):
    return fetch_df(
        "SELECT * FROM event_photos WHERE event_id=? ORDER BY id",
        (int(event_id),)
    )


def save_uploaded_files(event_id, uploaded_files, file_memo=""):
    """Excel等の添付ファイルをattached_filesフォルダに保存し、DBへ紐づける。"""
    if not uploaded_files:
        return

    for uploaded in uploaded_files:
        suffix = Path(uploaded.name).suffix.lower()
        safe_name = f"event_{event_id}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}{suffix}"
        save_path = FILE_DIR / safe_name

        with open(save_path, "wb") as f:
            f.write(uploaded.getbuffer())

        execute("""
            INSERT INTO event_files
            (event_id, file_name, file_path, file_type, file_memo, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            int(event_id),
            uploaded.name,
            str(save_path),
            suffix.replace(".", ""),
            file_memo.strip() or None,
            now_text(),
        ))


def get_event_files(event_id):
    return fetch_df(
        "SELECT * FROM event_files WHERE event_id=? ORDER BY id",
        (int(event_id),)
    )

def get_active_users():
    return fetch_df("SELECT user_id, user_name FROM users WHERE is_active=1 ORDER BY user_name")


def user_display_map():
    df = get_active_users()
    mapping = {"": ("", "")}
    if not df.empty:
        for _, r in df.iterrows():
            label = f"{r['user_name']}（ID:{r['user_id']}）"
            mapping[label] = (r["user_id"], r["user_name"])
    return mapping


def get_active_staff():
    df = fetch_df("SELECT staff_name FROM staff WHERE is_active=1 ORDER BY staff_name")
    return df["staff_name"].tolist() if not df.empty else []


# -----------------------------
# UI helpers
# -----------------------------
DEFAULT_CATEGORIES = ["通院", "面会", "行事", "外出", "注意", "申し送り", "夜勤", "その他"]

CATEGORY_MARK = {
    "通院": "🏥",
    "面会": "👪",
    "行事": "🎉",
    "外出": "🚶",
    "注意": "⚠️",
    "申し送り": "📝",
    "夜勤": "🌙",
    "その他": "・",
}


def get_categories(active_only=True):
    where = "WHERE is_active=1" if active_only else ""
    df = fetch_df(f"""
        SELECT category_name, mark, sort_order, is_active
        FROM categories
        {where}
        ORDER BY sort_order, category_name
    """)
    if df.empty and active_only:
        return DEFAULT_CATEGORIES
    return df["category_name"].tolist()


def get_category_mark(category_name):
    df = fetch_df(
        "SELECT mark FROM categories WHERE category_name=? LIMIT 1",
        (category_name,)
    )
    if not df.empty:
        mark = str(df.iloc[0]["mark"] or "").strip()
        if mark:
            return mark
    return CATEGORY_MARK.get(category_name, "・")


def add_css():
    st.markdown("""
    <style>
    .main .block-container {
        padding-top: 1.2rem;
        max-width: 1200px;
    }
    .calendar-title {
        font-size: 1.6rem;
        font-weight: 700;
        margin: 12px 0 10px 0;
        color: #3f3a35;
    }
    .calendar-grid {
        display: grid;
        grid-template-columns: repeat(7, minmax(0, 1fr));
        gap: 0;
        width: 100%;
        border-top: 1px solid #d8d0c4;
        border-left: 1px solid #d8d0c4;
        background: #fffdf8;
    }
    .day-head {
        font-weight: 700;
        text-align: center;
        padding: 8px 4px;
        border-right: 1px solid #d8d0c4;
        border-bottom: 1px solid #d8d0c4;
        background: #f3eee6;
        color: #4b4035;
        font-size: 0.82rem;
        letter-spacing: 0.08em;
    }
    .day-cell {
        min-height: 135px;
        border-right: 1px solid #d8d0c4;
        border-bottom: 1px solid #d8d0c4;
        padding: 8px;
        background: #fffdf8;
        box-sizing: border-box;
        overflow: hidden;
    }
    .blank-cell {
        background: #fffdf8;
    }
    .day-cell-muted {
        min-height: 150px;
        border: 1px solid #eee6dc;
        border-radius: 4px;
        padding: 8px;
        background: #fffdf8;
        color: #aaa;
    }
    .day-num {
        font-size: 2.4rem;
        font-weight: 700;
        line-height: 1;
        margin-bottom: 8px;
        letter-spacing: -1px;
    }
    .write-lines {
        height: 34px;
        margin: 4px 0 6px 0;
        background-image: repeating-linear-gradient(
            to bottom,
            transparent 0px,
            transparent 13px,
            #d9d2c8 14px
        );
        opacity: 0.75;
    }
    .sunday { color: #c0392b; }
    .saturday { color: #1f4e79; }
    .event-line {
        font-size: 0.78rem;
        line-height: 1.35;
        margin: 3px 0;
        padding: 3px 5px;
        border-radius: 5px;
        background: #f6efe6;
        overflow-wrap: anywhere;
        border-left: 3px solid #bfae9b;
    }
    .important {
        background: #ffe9e0;
        border-left: 4px solid #d65a31;
    }
    .small-note {
        color: #7a6a5b;
        font-size: 0.9rem;
    }

    .today-board-card {
        border: 2px solid #d8d0c4;
        border-left: 10px solid #bfae9b;
        background: #fffdf8;
        border-radius: 8px;
        padding: 14px 16px;
        margin: 10px 0;
        box-shadow: none;
    }
    .today-board-card-important {
        border-left: 10px solid #d65a31;
        background: #fff3ee;
    }
    .today-board-main {
        font-size: 1.35rem;
        font-weight: 800;
        line-height: 1.35;
        color: #3f3a35;
    }
    .today-board-time {
        display: inline-block;
        min-width: 76px;
        font-size: 1.45rem;
        font-weight: 900;
        color: #2f2a25;
    }
    .today-board-memo {
        font-size: 1.05rem;
        margin-top: 8px;
        padding-left: 82px;
        color: #4f463d;
    }
    .today-board-sub {
        font-size: 0.9rem;
        margin-top: 8px;
        padding-left: 82px;
        color: #7a6a5b;
    }
    .today-board-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 999px;
        background: #f2eadf;
        font-size: 0.82rem;
        margin-right: 4px;
    }
    .today-summary-box {
        border: 2px solid #d8d0c4;
        background: #f9f5ee;
        border-radius: 8px;
        padding: 12px 14px;
        margin: 8px 0 14px 0;
        font-size: 1.05rem;
        font-weight: 700;
    }

    </style>
    """, unsafe_allow_html=True)


def monthly_events(year, month):
    start = f"{year}-{month:02d}-01"
    last_day = calendar.monthrange(year, month)[1]
    end = f"{year}-{month:02d}-{last_day:02d}"

    df = fetch_df("""
        SELECT * FROM events
        WHERE event_date BETWEEN ? AND ?
        ORDER BY event_date, start_time, category, id
    """, (start, end))

    events_by_day = {}
    for _, row in df.iterrows():
        events_by_day.setdefault(row["event_date"], []).append(row)
    return events_by_day



def events_by_date(target_date):
    """指定日の予定一覧を取得する。"""
    if hasattr(target_date, "strftime"):
        target_date = target_date.strftime("%Y-%m-%d")
    return fetch_df("""
        SELECT *
        FROM events
        WHERE event_date = ?
        ORDER BY
            CASE WHEN start_time IS NULL OR start_time = '' THEN 1 ELSE 0 END,
            start_time,
            category,
            id
    """, (str(target_date),))


def render_event_button_list(df, empty_message="予定はありません。", include_date=True):
    """予定一覧をボタン表示し、押すと同一ページに詳細を出す。"""
    if df is None or df.empty:
        st.info(empty_message)
        return

    cols = st.columns(3)
    for i, (_, ev) in enumerate(df.iterrows()):
        mark = get_category_mark(ev["category"])
        label_date = str(ev["event_date"])[5:] if include_date and ev["event_date"] else ""
        label_time = f" {ev['start_time']}" if ev["start_time"] else ""
        label_user = f"／{ev['user_name']}" if ev["user_name"] else ""
        important = "⚠️ " if int(ev["important"] or 0) == 1 else ""
        label = f"{label_date}{label_time} {important}{mark}{ev['title']}{label_user}".strip()

        with cols[i % 3]:
            if st.button(label, key=f"event_btn_{ev['id']}", use_container_width=True):
                set_selected_event(int(ev["id"]))
                st.rerun()



def first_line_text(text, max_len=42):
    """メモの1行目だけを短く表示する。"""
    if text is None:
        return ""
    value = str(text).strip()
    if not value:
        return ""
    first = value.splitlines()[0].strip()
    if len(first) > max_len:
        return first[:max_len] + "…"
    return first


def get_attachment_counts(event_id):
    """写真・添付ファイル数を取得する。"""
    photos = fetch_df("SELECT COUNT(*) AS cnt FROM event_photos WHERE event_id=?", (int(event_id),))
    files = fetch_df("SELECT COUNT(*) AS cnt FROM event_files WHERE event_id=?", (int(event_id),))
    photo_count = int(photos.iloc[0]["cnt"]) if not photos.empty else 0
    file_count = int(files.iloc[0]["cnt"]) if not files.empty else 0
    return photo_count, file_count


def render_today_board(df):
    """
    今日画面専用のホワイトボード風一覧。
    クリックしなくても7割分かるように、時刻・カテゴリ・利用者・メモ1行・担当・添付を表示する。
    """
    if df is None or df.empty:
        st.info("今日の予定はありません。")
        return

    for _, ev in df.iterrows():
        important = int(ev["important"] or 0) == 1
        card_class = "today-board-card today-board-card-important" if important else "today-board-card"
        time_text = ev["start_time"] if ev["start_time"] else "時間未定"
        mark = get_category_mark(ev["category"])
        category = html_escape(ev["category"])
        title = html_escape(ev["title"])
        user_name = html_escape(ev["user_name"] or "")
        staff_name = html_escape(ev["staff_name"] or "")
        memo_line = html_escape(first_line_text(ev["memo"]))
        warning = "⚠️ " if important else ""

        main_line = f'{warning}<span class="today-board-time">{html_escape(time_text)}</span>{mark} {category}｜{title}'
        if user_name:
            main_line += f'｜{user_name}'

        memo_html = f'<div class="today-board-memo">メモ：{memo_line}</div>' if memo_line else ""

        sub_items = []
        if staff_name:
            sub_items.append(f'<span class="today-board-badge">担当：{staff_name}</span>')

        photo_count, file_count = get_attachment_counts(ev["id"])
        if photo_count:
            sub_items.append(f'<span class="today-board-badge">写真 {photo_count}</span>')
        if file_count:
            sub_items.append(f'<span class="today-board-badge">添付 {file_count}</span>')

        sub_html = f'<div class="today-board-sub">{" ".join(sub_items)}</div>' if sub_items else ""

        html = (
            f'<div class="{card_class}">'
            f'<div class="today-board-main">{main_line}</div>'
            f'{memo_html}'
            f'{sub_html}'
            f'</div>'
        )
        st.markdown(html, unsafe_allow_html=True)

        if st.button("詳細を見る", key=f"today_detail_{ev['id']}", use_container_width=True):
            set_selected_event(int(ev["id"]))
            st.rerun()


def html_escape(text):
    """HTML表示用の簡易エスケープ。"""
    if text is None:
        return ""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#x27;")
    )


def set_selected_event(event_id):
    st.session_state["selected_calendar_event_id"] = int(event_id)


def get_event_by_id(event_id):
    if not event_id:
        return pd.DataFrame()
    return fetch_df("SELECT * FROM events WHERE id=? LIMIT 1", (int(event_id),))


def render_event_detail_panel():
    event_id = st.session_state.get("selected_calendar_event_id")
    if not event_id:
        return

    ev_df = get_event_by_id(event_id)
    if ev_df.empty:
        st.warning("選択された予定が見つかりません。")
        st.session_state["selected_calendar_event_id"] = None
        return

    ev = ev_df.iloc[0]
    st.markdown("---")
    st.subheader("予定詳細")
    st.markdown(f"### {ev['event_date']}｜{ev['category']}｜{ev['title']}")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.write(f"**利用者**：{ev['user_name'] or ''}")
        st.caption(f"利用者ID：{ev['user_id'] or ''}")
    with c2:
        st.write(f"**時間**：{ev['start_time'] or ''} 〜 {ev['end_time'] or ''}")
    with c3:
        st.write(f"**担当**：{ev['staff_name'] or ''}")

    if ev["important"]:
        st.warning("重要マークあり")

    st.write("**メモ**")
    st.info(ev["memo"] or "メモはありません。")

    photos = get_event_photos(event_id)
    if not photos.empty:
        st.write("**写真メモ**")
        cols = st.columns(3)
        for i, (_, p) in enumerate(photos.iterrows()):
            with cols[i % 3]:
                img_path = Path(p["file_path"])
                if img_path.exists():
                    st.image(str(img_path), caption=p["photo_memo"] or p["file_name"], use_container_width=True)
                else:
                    st.warning(f"画像が見つかりません：{p['file_name']}")

    files = get_event_files(event_id)
    if not files.empty:
        st.write("**Excel・書類ファイル**")
        for _, frow in files.iterrows():
            f_path = Path(frow["file_path"])
            st.write(f"📎 {frow['file_name']}　{frow['file_memo'] or ''}")
            if f_path.exists():
                with open(f_path, "rb") as f:
                    st.download_button(
                        "ダウンロード",
                        data=f,
                        file_name=frow["file_name"],
                        key=f"detail_download_{frow['id']}"
                    )
            else:
                st.warning("ファイルが見つかりません。")

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("詳細を閉じる", use_container_width=True):
            st.session_state["selected_calendar_event_id"] = None
            st.rerun()
    with col_b:
        st.caption("編集・削除は「予定検索・更新・削除」メニューで行います。")



def render_calendar(year, month):
    """
    紙の壁カレンダー風レイアウト。
    カレンダーはHTMLで表示し、予定詳細は下部の予定ボタンから同じページ内に表示する。
    """
    events_by_day = monthly_events(year, month)
    first_weekday, last_day = calendar.monthrange(year, month)  # Monday=0, Sunday=6
    start_col = (first_weekday + 1) % 7

    cells = [""] * 42
    for day in range(1, last_day + 1):
        cells[start_col + day - 1] = day

    html = []
    html.append(f'<div class="calendar-title">{year}年 {month}月</div>')
    html.append('<div class="calendar-grid">')

    for h in ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]:
        html.append(f'<div class="day-head">{h}</div>')

    for idx, day in enumerate(cells):
        dow = idx % 7
        dow_cls = "sunday" if dow == 0 else ("saturday" if dow == 6 else "")

        if day == "":
            html.append('<div class="day-cell blank-cell"></div>')
            continue

        key = f"{year}-{month:02d}-{int(day):02d}"
        html.append('<div class="day-cell">')
        html.append(f'<div class="day-num {dow_cls}">{day}</div>')
        html.append('<div class="write-lines"></div>')

        for ev in events_by_day.get(key, []):
            mark = get_category_mark(ev["category"])
            time_part = f'{ev["start_time"]} ' if ev["start_time"] else ""
            user_part = f'／{ev["user_name"]}' if ev["user_name"] else ""
            imp_cls = " important" if int(ev["important"] or 0) == 1 else ""
            text = html_escape(f'{mark}{time_part}{ev["title"]}{user_part}')
            html.append(f'<div class="event-line{imp_cls}">{text}</div>')

        html.append('</div>')

    html.append('</div>')
    st.markdown("".join(html), unsafe_allow_html=True)

    # Streamlitの通常ボタンで詳細表示。HTML内クリックより安定。
    month_events = []
    for key in sorted(events_by_day.keys()):
        for ev in events_by_day[key]:
            month_events.append(ev)

    if month_events:
        st.markdown("### 予定詳細を表示")
        st.caption("下の予定ボタンから選ぶと、詳細が同じページ内に表示されます。")
        month_df = pd.DataFrame(month_events)
        render_event_button_list(month_df, empty_message="この月の予定はありません。", include_date=True)

    render_event_detail_panel()


# -----------------------------
# Pages
# -----------------------------
def page_calendar():
    today = today_jst()
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        year = st.number_input("", min_value=2020, max_value=2100, value=today.year, step=1)
    with col2:
        month = st.number_input("", min_value=1, max_value=12, value=today.month, step=1)
    with col3:
        st.markdown('<div class="small-note"></div>', unsafe_allow_html=True)

    render_calendar(int(year), int(month))



def page_today():
    st.subheader("今日は何ある")
    st.caption("朝の申し送り前に、今日の予定をホワイトボード感覚で確認できます。")

    target = today_jst()
    target_text = target.strftime("%Y-%m-%d")
    st.markdown(f"## {target_text} の予定")

    df = events_by_date(target)

    if not df.empty:
        summary = df.groupby("category").size().reset_index(name="件数")
        summary_text = "　".join([
            f"{get_category_mark(r['category'])}{r['category']} {r['件数']}件"
            for _, r in summary.iterrows()
        ])

        important_df = df[df["important"].fillna(0).astype(int) == 1]
        important_text = f"　⚠️重要 {len(important_df)}件" if not important_df.empty else ""

        st.markdown(
            f'<div class="today-summary-box">本日の予定：{len(df)}件　{summary_text}{important_text}</div>',
            unsafe_allow_html=True
        )

        st.markdown("### 今日のホワイトボード")
        render_today_board(df)
    else:
        st.info("今日の予定はありません。")

    render_event_detail_panel()


def page_event_register():
    st.subheader("予定登録")

    user_map = user_display_map()
    users = list(user_map.keys())
    staff = [""] + get_active_staff()

    with st.form("event_register_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            event_date = st.date_input("日付", value=today_jst())
            category = st.selectbox("カテゴリ", get_categories())
        with c2:
            start_time = st.text_input("開始時刻", placeholder="例：10:00")
            end_time = st.text_input("終了時刻", placeholder="例：11:00")
        with c3:
            user_name = st.selectbox("利用者", users)
            staff_name = st.selectbox("担当職員", staff)

        title = st.text_input("予定タイトル", placeholder="例：内科受診、家族面会、外出支援")
        memo = st.text_area("メモ", placeholder="必要な注意点、持ち物、申し送りなど")
        important = st.checkbox("重要マークを付ける")
        uploaded_photos = st.file_uploader(
            "写真メモ（複数可）",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            help="紙カレンダー、受診票、家族メモ、持ち物写真などを予定に紐づけできます。"
        )
        photo_memo = st.text_input("写真メモ補足", placeholder="例：受診予定表、家族からのメモ、持ち物確認")

        uploaded_files = st.file_uploader(
            "Excel・書類ファイル（複数可）",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="通院予定表、持ち物表、行事予定表、家族共有用Excelなどを予定に紐づけできます。"
        )
        file_memo = st.text_input("Excel・書類メモ補足", placeholder="例：通院予定表、買い物リスト、行事参加者表")

        submitted = st.form_submit_button("保存する")

    if submitted:
        if not title.strip():
            st.error("予定タイトルを入力してください。")
            return

        selected_user_id, selected_user_name = user_map.get(user_name, ("", ""))

        event_id = execute("""
            INSERT INTO events
            (event_date, category, title, user_id, user_name, staff_name, start_time, end_time, memo, important, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            event_date.strftime("%Y-%m-%d"),
            category,
            title.strip(),
            selected_user_id or None,
            selected_user_name or None,
            staff_name or None,
            start_time.strip() or None,
            end_time.strip() or None,
            memo.strip() or None,
            1 if important else 0,
            now_text(),
            now_text(),
        ))
        save_uploaded_photos(event_id, uploaded_photos, photo_memo)
        save_uploaded_files(event_id, uploaded_files, file_memo)
        st.success("予定を保存しました。写真メモ・Excelファイルも紐づけました。")


def page_event_manage():
    st.subheader("予定検索・更新・削除")

    c1, c2, c3 = st.columns(3)
    with c1:
        start = st.date_input("開始日", value=today_jst().replace(day=1))
    with c2:
        end = st.date_input("終了日", value=today_jst())
    with c3:
        category_filter = st.selectbox("カテゴリ絞り込み", ["すべて"] + get_categories())

    keyword = st.text_input("キーワード検索", placeholder="タイトル・メモ・利用者名・職員名")

    query = """
        SELECT * FROM events
        WHERE event_date BETWEEN ? AND ?
    """
    params = [start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")]

    if category_filter != "すべて":
        query += " AND category = ?"
        params.append(category_filter)

    if keyword.strip():
        query += " AND (title LIKE ? OR memo LIKE ? OR user_name LIKE ? OR staff_name LIKE ?)"
        kw = f"%{keyword.strip()}%"
        params.extend([kw, kw, kw, kw])

    query += " ORDER BY event_date DESC, start_time, id DESC"
    df = fetch_df(query, params)

    if df.empty:
        st.info("該当する予定はありません。")
        return

    show_cols = ["id", "event_date", "category", "title", "user_id", "user_name", "staff_name", "start_time", "end_time", "important", "memo"]
    st.dataframe(df[show_cols], use_container_width=True, hide_index=True)

    selected_id = st.selectbox("編集・削除する予定ID", df["id"].tolist())
    target = df[df["id"] == selected_id].iloc[0]

    st.markdown("---")
    st.write("選択中の予定の写真メモ")

    photos = get_event_photos(selected_id)
    if photos.empty:
        st.info("この予定に紐づく写真メモはありません。")
    else:
        for _, p in photos.iterrows():
            img_path = Path(p["file_path"])
            c_img, c_info = st.columns([1, 2])
            with c_img:
                if img_path.exists():
                    st.image(str(img_path), caption=p["file_name"], use_container_width=True)
                else:
                    st.warning("画像ファイルが見つかりません。")
            with c_info:
                st.write(f"メモ：{p['photo_memo'] or ''}")
                st.caption(f"登録日時：{p['created_at']}")
                if st.button(f"この写真を削除 ID:{p['id']}", key=f"delete_photo_{p['id']}"):
                    if img_path.exists():
                        try:
                            img_path.unlink()
                        except Exception:
                            pass
                    execute("DELETE FROM event_photos WHERE id=?", (int(p["id"]),))
                    st.success("写真メモを削除しました。画面を再読み込みしてください。")

    with st.expander("この予定に写真メモを追加する"):
        add_photos = st.file_uploader(
            "追加する写真",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key=f"add_photos_{selected_id}"
        )
        add_photo_memo = st.text_input("追加写真メモ", key=f"add_photo_memo_{selected_id}")
        if st.button("写真メモを追加", key=f"add_photo_btn_{selected_id}"):
            save_uploaded_photos(selected_id, add_photos, add_photo_memo)
            st.success("写真メモを追加しました。画面を再読み込みしてください。")

    st.markdown("---")
    st.write("選択中の予定のExcel・書類ファイル")

    files = get_event_files(selected_id)
    if files.empty:
        st.info("この予定に紐づくExcel・書類ファイルはありません。")
    else:
        for _, frow in files.iterrows():
            f_path = Path(frow["file_path"])
            c_file, c_action = st.columns([3, 1])
            with c_file:
                st.write(f"📎 **{frow['file_name']}**")
                st.write(f"メモ：{frow['file_memo'] or ''}")
                st.caption(f"登録日時：{frow['created_at']}")
                if f_path.exists():
                    with open(f_path, "rb") as f:
                        st.download_button(
                            "ダウンロード",
                            data=f,
                            file_name=frow["file_name"],
                            key=f"download_file_{frow['id']}"
                        )
                else:
                    st.warning("ファイルが見つかりません。")
            with c_action:
                if st.button(f"削除 ID:{frow['id']}", key=f"delete_file_{frow['id']}"):
                    if f_path.exists():
                        try:
                            f_path.unlink()
                        except Exception:
                            pass
                    execute("DELETE FROM event_files WHERE id=?", (int(frow["id"]),))
                    st.success("ファイルを削除しました。画面を再読み込みしてください。")

    with st.expander("この予定にExcel・書類ファイルを追加する"):
        add_files = st.file_uploader(
            "追加するExcel・書類",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key=f"add_files_{selected_id}"
        )
        add_file_memo = st.text_input("追加ファイルメモ", key=f"add_file_memo_{selected_id}")
        if st.button("Excel・書類ファイルを追加", key=f"add_file_btn_{selected_id}"):
            save_uploaded_files(selected_id, add_files, add_file_memo)
            st.success("Excel・書類ファイルを追加しました。画面を再読み込みしてください。")

    st.markdown("---")
    st.write("選択中の予定を編集")

    user_map = user_display_map()
    users = list(user_map.keys())
    staff = [""] + get_active_staff()

    with st.form("event_update_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            new_date = st.date_input("日付", value=datetime.strptime(target["event_date"], "%Y-%m-%d").date())
            category_options = get_categories()
            if target["category"] and target["category"] not in category_options:
                category_options = [target["category"]] + category_options
            new_category = st.selectbox(
                "カテゴリ",
                category_options,
                index=category_options.index(target["category"]) if target["category"] in category_options else 0
            )
        with c2:
            new_start_time = st.text_input("開始時刻", value=target["start_time"] or "")
            new_end_time = st.text_input("終了時刻", value=target["end_time"] or "")
        with c3:
            current_user_label = ""
            if target["user_name"] and target["user_id"]:
                current_user_label = f"{target['user_name']}（ID:{target['user_id']}）"
            elif target["user_name"]:
                for label, pair in user_map.items():
                    if pair[1] == target["user_name"]:
                        current_user_label = label
                        break
            user_index = users.index(current_user_label) if current_user_label in users else 0
            staff_index = staff.index(target["staff_name"]) if target["staff_name"] in staff else 0
            new_user = st.selectbox("利用者", users, index=user_index)
            new_staff = st.selectbox("担当職員", staff, index=staff_index)

        new_title = st.text_input("予定タイトル", value=target["title"])
        new_memo = st.text_area("メモ", value=target["memo"] or "")
        new_important = st.checkbox("重要マーク", value=bool(target["important"]))

        c_update, c_delete = st.columns(2)
        with c_update:
            update_btn = st.form_submit_button("更新する")
        with c_delete:
            delete_btn = st.form_submit_button("削除する")

    if update_btn:
        if not new_title.strip():
            st.error("予定タイトルを入力してください。")
            return

        new_user_id, new_user_name = user_map.get(new_user, ("", ""))

        execute("""
            UPDATE events
            SET event_date=?, category=?, title=?, user_id=?, user_name=?, staff_name=?,
                start_time=?, end_time=?, memo=?, important=?, updated_at=?
            WHERE id=?
        """, (
            new_date.strftime("%Y-%m-%d"),
            new_category,
            new_title.strip(),
            new_user_id or None,
            new_user_name or None,
            new_staff or None,
            new_start_time.strip() or None,
            new_end_time.strip() or None,
            new_memo.strip() or None,
            1 if new_important else 0,
            now_text(),
            int(selected_id),
        ))
        st.success("予定を更新しました。画面を再読み込みしてください。")

    if delete_btn:
        photos = get_event_photos(selected_id)
        for _, p in photos.iterrows():
            img_path = Path(p["file_path"])
            if img_path.exists():
                try:
                    img_path.unlink()
                except Exception:
                    pass
        files = get_event_files(selected_id)
        for _, frow in files.iterrows():
            f_path = Path(frow["file_path"])
            if f_path.exists():
                try:
                    f_path.unlink()
                except Exception:
                    pass
        execute("DELETE FROM event_photos WHERE event_id=?", (int(selected_id),))
        execute("DELETE FROM event_files WHERE event_id=?", (int(selected_id),))
        execute("DELETE FROM events WHERE id=?", (int(selected_id),))
        st.warning("予定と紐づく写真メモ・Excelファイルを削除しました。画面を再読み込みしてください。")



def page_category_master():
    st.subheader("予定カテゴリ設定")
    st.caption("予定登録で使うカテゴリを追加・編集・非表示にできます。")

    with st.form("category_add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            category_name = st.text_input("カテゴリ名", placeholder="例：訪問診療、往診、買い物、家族連絡")
        with c2:
            mark = st.text_input("マーク", placeholder="例：🏥", value="・")
        with c3:
            sort_order = st.number_input("並び順", min_value=1, max_value=999, value=100, step=10)

        add = st.form_submit_button("カテゴリを追加")

    if add:
        if not category_name.strip():
            st.error("カテゴリ名を入力してください。")
        else:
            try:
                execute("""
                    INSERT INTO categories
                    (category_name, mark, sort_order, is_active, created_at, updated_at)
                    VALUES (?, ?, ?, 1, ?, ?)
                """, (
                    category_name.strip(),
                    mark.strip() or "・",
                    int(sort_order),
                    now_text(),
                    now_text(),
                ))
                st.success("カテゴリを追加しました。")
            except Exception as e:
                st.error(f"追加できませんでした。同じカテゴリ名がある可能性があります：{e}")

    st.markdown("---")
    df = fetch_df("""
        SELECT id, category_name, mark, sort_order, is_active, created_at, updated_at
        FROM categories
        ORDER BY sort_order, category_name
    """)

    if df.empty:
        st.info("カテゴリが登録されていません。")
        return

    st.dataframe(df, use_container_width=True, hide_index=True)

    selected_id = st.selectbox("編集するカテゴリID", df["id"].tolist())
    target = df[df["id"] == selected_id].iloc[0]

    with st.form("category_edit_form"):
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        with c1:
            new_name = st.text_input("カテゴリ名", value=target["category_name"])
        with c2:
            new_mark = st.text_input("マーク", value=target["mark"] or "・")
        with c3:
            new_sort = st.number_input("並び順", min_value=1, max_value=999, value=int(target["sort_order"] or 100), step=10)
        with c4:
            new_active = st.selectbox(
                "状態",
                [1, 0],
                index=0 if int(target["is_active"] or 1) == 1 else 1,
                format_func=lambda x: "表示" if x == 1 else "非表示"
            )

        c_update, c_delete = st.columns(2)
        with c_update:
            update = st.form_submit_button("更新する")
        with c_delete:
            delete = st.form_submit_button("削除する")

    if update:
        if not new_name.strip():
            st.error("カテゴリ名を入力してください。")
        else:
            try:
                execute("""
                    UPDATE categories
                    SET category_name=?, mark=?, sort_order=?, is_active=?, updated_at=?
                    WHERE id=?
                """, (
                    new_name.strip(),
                    new_mark.strip() or "・",
                    int(new_sort),
                    int(new_active),
                    now_text(),
                    int(selected_id),
                ))
                st.success("カテゴリを更新しました。画面を再読み込みしてください。")
            except Exception as e:
                st.error(f"更新できませんでした：{e}")

    if delete:
        # 既存予定で使われているカテゴリは削除ではなく非表示推奨
        used = fetch_df("SELECT COUNT(*) AS cnt FROM events WHERE category=?", (target["category_name"],))
        count = int(used.iloc[0]["cnt"]) if not used.empty else 0
        if count > 0:
            execute("UPDATE categories SET is_active=0, updated_at=? WHERE id=?", (now_text(), int(selected_id)))
            st.warning(f"このカテゴリは既存予定で {count} 件使われているため、削除せず非表示にしました。")
        else:
            execute("DELETE FROM categories WHERE id=?", (int(selected_id),))
            st.warning("カテゴリを削除しました。画面を再読み込みしてください。")


def page_master_users():
    st.subheader("利用者マスタ")

    with st.form("user_add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            user_id = st.text_input("利用者ID", placeholder="例：U0001 / 健康管理アプリ側のID")
        with c2:
            user_name = st.text_input("利用者名")
        with c3:
            kana = st.text_input("ふりがな")
        room_no = st.text_input("居室番号", placeholder="例：101")
        note = st.text_area("メモ")
        add = st.form_submit_button("利用者を追加")

    if add:
        if not user_name.strip():
            st.error("利用者名を入力してください。")
        else:
            try:
                if not user_id.strip():
                    max_df = fetch_df("SELECT COALESCE(MAX(id), 0) + 1 AS next_id FROM users")
                    user_id_value = f"U{int(max_df.iloc[0]['next_id']):04d}"
                else:
                    user_id_value = user_id.strip()

                execute("""
                    INSERT INTO users (user_id, user_name, kana, room_no, note, is_active, created_at)
                    VALUES (?, ?, ?, ?, ?, 1, ?)
                """, (user_id_value, user_name.strip(), kana.strip(), room_no.strip(), note.strip(), now_text()))
                st.success("利用者を追加しました。")
            except sqlite3.IntegrityError:
                st.error("同じ利用者名がすでに登録されています。")

    df = fetch_df("SELECT * FROM users ORDER BY is_active DESC, user_name")
    st.dataframe(df, use_container_width=True, hide_index=True)

    if not df.empty:
        selected = st.selectbox("有効／無効を変更する利用者ID", df["id"].tolist())
        target = df[df["id"] == selected].iloc[0]
        new_active = st.radio("状態", [1, 0], index=0 if target["is_active"] == 1 else 1, format_func=lambda x: "有効" if x == 1 else "無効")
        if st.button("利用者状態を更新"):
            execute("UPDATE users SET is_active=? WHERE id=?", (int(new_active), int(selected)))
            st.success("状態を更新しました。")


def page_master_staff():
    st.subheader("職員マスタ")

    with st.form("staff_add_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            staff_name = st.text_input("職員名")
        with c2:
            role = st.text_input("役割", placeholder="例：管理者、日勤、夜勤")
        add = st.form_submit_button("職員を追加")

    if add:
        if not staff_name.strip():
            st.error("職員名を入力してください。")
        else:
            try:
                execute("""
                    INSERT INTO staff (staff_name, role, is_active, created_at)
                    VALUES (?, ?, 1, ?)
                """, (staff_name.strip(), role.strip(), now_text()))
                st.success("職員を追加しました。")
            except sqlite3.IntegrityError:
                st.error("同じ職員名がすでに登録されています。")

    df = fetch_df("SELECT * FROM staff ORDER BY is_active DESC, staff_name")
    st.dataframe(df, use_container_width=True, hide_index=True)

    if not df.empty:
        selected = st.selectbox("有効／無効を変更する職員ID", df["id"].tolist())
        target = df[df["id"] == selected].iloc[0]
        new_active = st.radio("状態", [1, 0], index=0 if target["is_active"] == 1 else 1, format_func=lambda x: "有効" if x == 1 else "無効")
        if st.button("職員状態を更新"):
            execute("UPDATE staff SET is_active=? WHERE id=?", (int(new_active), int(selected)))
            st.success("状態を更新しました。")



def page_photo_notes():
    st.subheader("写真メモ一覧")

    df = fetch_df("""
        SELECT
            p.id AS photo_id,
            p.file_name,
            p.file_path,
            p.photo_memo,
            p.created_at AS photo_created_at,
            e.id AS event_id,
            e.event_date,
            e.category,
            e.title,
            e.user_id,
            e.user_name,
            e.staff_name,
            e.memo
        FROM event_photos p
        JOIN events e ON p.event_id = e.id
        ORDER BY e.event_date DESC, p.id DESC
    """)

    if df.empty:
        st.info("写真メモはまだ登録されていません。")
        return

    keyword = st.text_input("写真メモ検索", placeholder="予定タイトル・利用者名・写真メモなど")
    if keyword.strip():
        kw = keyword.strip()
        mask = (
            df["title"].fillna("").str.contains(kw, case=False, na=False) |
            df["user_id"].fillna("").str.contains(kw, case=False, na=False) |
            df["user_name"].fillna("").str.contains(kw, case=False, na=False) |
            df["staff_name"].fillna("").str.contains(kw, case=False, na=False) |
            df["photo_memo"].fillna("").str.contains(kw, case=False, na=False) |
            df["memo"].fillna("").str.contains(kw, case=False, na=False)
        )
        df = df[mask]

    if df.empty:
        st.info("該当する写真メモはありません。")
        return

    for _, row in df.iterrows():
        st.markdown("---")
        c1, c2 = st.columns([1, 2])
        with c1:
            img_path = Path(row["file_path"])
            if img_path.exists():
                st.image(str(img_path), use_container_width=True)
            else:
                st.warning("画像ファイルが見つかりません。")
        with c2:
            st.write(f"**{row['event_date']}｜{row['category']}｜{row['title']}**")
            if row["user_name"]:
                st.write(f"利用者：{row['user_name']}（ID:{row['user_id'] or ''}）")
            if row["staff_name"]:
                st.write(f"担当：{row['staff_name']}")
            st.write(f"写真メモ：{row['photo_memo'] or ''}")
            if row["memo"]:
                st.caption(f"予定メモ：{row['memo']}")
            st.caption(f"写真登録：{row['photo_created_at']}")



def page_attached_files():
    st.subheader("Excel・書類ファイル一覧")

    df = fetch_df("""
        SELECT
            f.id AS file_id,
            f.file_name,
            f.file_path,
            f.file_type,
            f.file_memo,
            f.created_at AS file_created_at,
            e.id AS event_id,
            e.event_date,
            e.category,
            e.title,
            e.user_id,
            e.user_name,
            e.staff_name,
            e.memo
        FROM event_files f
        JOIN events e ON f.event_id = e.id
        ORDER BY e.event_date DESC, f.id DESC
    """)

    if df.empty:
        st.info("Excel・書類ファイルはまだ登録されていません。")
        return

    keyword = st.text_input("ファイル検索", placeholder="予定タイトル・利用者名・ファイル名・メモなど")
    if keyword.strip():
        kw = keyword.strip()
        mask = (
            df["title"].fillna("").str.contains(kw, case=False, na=False) |
            df["user_id"].fillna("").str.contains(kw, case=False, na=False) |
            df["user_name"].fillna("").str.contains(kw, case=False, na=False) |
            df["staff_name"].fillna("").str.contains(kw, case=False, na=False) |
            df["file_name"].fillna("").str.contains(kw, case=False, na=False) |
            df["file_memo"].fillna("").str.contains(kw, case=False, na=False) |
            df["memo"].fillna("").str.contains(kw, case=False, na=False)
        )
        df = df[mask]

    if df.empty:
        st.info("該当するファイルはありません。")
        return

    for _, row in df.iterrows():
        st.markdown("---")
        st.write(f"📎 **{row['file_name']}**")
        st.write(f"{row['event_date']}｜{row['category']}｜{row['title']}")
        if row["user_name"]:
            st.write(f"利用者：{row['user_name']}")
        if row["staff_name"]:
            st.write(f"担当：{row['staff_name']}")
        st.write(f"ファイルメモ：{row['file_memo'] or ''}")
        st.caption(f"登録：{row['file_created_at']}")

        f_path = Path(row["file_path"])
        if f_path.exists():
            with open(f_path, "rb") as f:
                st.download_button(
                    "ダウンロード",
                    data=f,
                    file_name=row["file_name"],
                    key=f"file_list_download_{row['file_id']}"
                )
        else:
            st.warning("ファイルが見つかりません。")



# -----------------------------
# 予定データ取込（健康管理アプリ → ひだまり帳）
# -----------------------------
def normalize_import_bool(value, default=True):
    """Excel/CSV上のチェック値をTrue/Falseへ寄せる。"""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    s = str(value).strip().lower()
    if s in ["true", "1", "yes", "y", "登録", "登録する", "する", "〇", "○", "✓", "☑"]:
        return True
    if s in ["false", "0", "no", "n", "登録しない", "しない", "×", "✕", "☐"]:
        return False
    return default


def normalize_import_date(value):
    """日付をYYYY-MM-DD文字列へ正規化。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        return ""
    # 2026-05-27 11:00 のような値は日付だけ使う
    dt = pd.to_datetime(text, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%Y-%m-%d")
    return text[:10]


def normalize_import_time(value):
    """時刻をHH:MMへ正規化。空なら空文字。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    text = str(value).strip()
    if not text or text.lower() in ["none", "nan", "nat"]:
        return ""
    # Excel由来で 11:00:00 になる場合
    if re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", text):
        parts = text.split(":")
        return f"{int(parts[0]):02d}:{parts[1]}"
    # 0.5 のようなExcel時刻シリアル対策
    try:
        f = float(text)
        if 0 <= f < 1:
            minutes = int(round(f * 24 * 60))
            return f"{minutes//60:02d}:{minutes%60:02d}"
    except Exception:
        pass
    return text


def pick_import_col(df, candidates):
    """候補列名のうち、存在する最初の列名を返す。"""
    for c in candidates:
        if c in df.columns:
            return c
    return None


def read_schedule_import_file(uploaded_file):
    """予定候補Excel/CSVを読み込む。複数シートExcelは先頭シートを読む。"""
    name = str(getattr(uploaded_file, "name", "")).lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, encoding="utf-8-sig")
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="cp932")
    return pd.read_excel(uploaded_file)


def clean_import_value(value):
    """nan / NaT / Noneを空文字に寄せる。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in ["nan", "nat", "none"]:
        return ""
    return text


def make_upload_key(uploaded_file):
    """ファイルが変わった時にdata_editorの古い状態が残らないようにする。"""
    name = str(getattr(uploaded_file, "name", "uploaded"))
    size = int(getattr(uploaded_file, "size", 0) or 0)
    return hashlib.md5(f"{name}-{size}".encode("utf-8")).hexdigest()[:10]


def normalize_schedule_import_df(raw_df):
    """
    健康管理アプリ側の出力列を、ひだまり帳events登録用に正規化する。
    元データとの対応が分かるように、元行番号を保持する。
    """
    base_cols = [
        "元行", "登録する", "event_date", "start_time", "end_time", "category", "title",
        "user_id", "user_name", "staff_name", "memo", "important", "取込状態"
    ]
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=base_cols)

    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 健康チェックアプリ側の列名ゆれを広めに吸収
    col_register = pick_import_col(df, ["登録する", "取込対象", "取込", "登録", "import", "selected"])
    col_date = pick_import_col(df, ["event_date", "予定日", "日付", "開始日", "日時", "開始日時", "予定日時", "実施日"])
    col_start = pick_import_col(df, ["start_time", "開始時刻", "開始時間", "時刻", "予定時刻", "時間"])
    col_end = pick_import_col(df, ["end_time", "終了時刻", "終了時間"])
    col_category = pick_import_col(df, ["category", "分類", "カテゴリ", "予定分類", "種別", "キーワード"])
    col_title = pick_import_col(df, ["title", "タイトル", "件名", "予定タイトル", "予定", "候補", "予定候補"])
    col_user_id = pick_import_col(df, ["user_id", "利用者ID", "利用者id", "入居者ID", "入居者id"])
    col_user_name = pick_import_col(df, ["user_name", "利用者名", "利用者", "入居者名", "入居者", "対象者"])
    col_staff = pick_import_col(df, ["staff_name", "担当", "担当者", "職員", "記入者", "作成者"])
    col_memo = pick_import_col(df, ["memo", "詳細", "メモ", "内容", "備考", "申し送り", "申し送り内容", "本文", "元の申し送り"])
    col_important = pick_import_col(df, ["important", "重要", "重要マーク", "注意", "要注意"])

    rows = []
    for original_index, r in df.iterrows():
        raw_date_value = r.get(col_date, "") if col_date else ""
        event_date = normalize_import_date(raw_date_value)

        start_time = normalize_import_time(r.get(col_start, "")) if col_start else ""
        end_time = normalize_import_time(r.get(col_end, "")) if col_end else ""

        # 日時列に時刻が含まれる場合は、開始時刻にも反映
        if not start_time and col_date:
            dt = pd.to_datetime(raw_date_value, errors="coerce")
            if pd.notna(dt) and (dt.hour != 0 or dt.minute != 0):
                start_time = dt.strftime("%H:%M")

        category = clean_import_value(r.get(col_category, "")) if col_category else ""
        title = clean_import_value(r.get(col_title, "")) if col_title else ""
        user_id = clean_import_value(r.get(col_user_id, "")) if col_user_id else ""
        user_name = clean_import_value(r.get(col_user_name, "")) if col_user_name else ""
        staff_name = clean_import_value(r.get(col_staff, "")) if col_staff else ""
        memo = clean_import_value(r.get(col_memo, "")) if col_memo else ""

        # タイトルが空なら、申し送り本文やカテゴリから予定名を補う
        if not category:
            category = "その他"
        if not title:
            if memo:
                title = memo.splitlines()[0][:30]
            else:
                title = category or "予定"

        important = 1 if normalize_import_bool(r.get(col_important, False), default=False) else 0
        register = normalize_import_bool(r.get(col_register, True), default=True) if col_register else True

        status = "OK"
        if not event_date:
            status = "日付なし"
        elif not title:
            status = "タイトルなし"

        rows.append({
            "元行": int(original_index) + 2,  # Excel上の見た目に近い行番号。1行目は見出し想定
            "登録する": bool(register),
            "event_date": event_date,
            "start_time": start_time,
            "end_time": end_time,
            "category": category,
            "title": title,
            "user_id": user_id,
            "user_name": user_name,
            "staff_name": staff_name,
            "memo": memo,
            "important": int(important),
            "取込状態": status,
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    keep_cols = ["event_date", "title", "memo", "user_id", "user_name"]
    out = out[out[keep_cols].astype(str).agg("".join, axis=1).str.strip() != ""].copy()
    return out.reset_index(drop=True)


def make_import_compare_df(raw_df, import_df):
    """元データと変換後予定候補を横並びで確認する表を作る。"""
    if raw_df is None or raw_df.empty or import_df is None or import_df.empty:
        return pd.DataFrame()
    raw = raw_df.copy().reset_index(drop=True)
    raw.insert(0, "元行", raw.index + 2)
    preview_cols = ["元行"] + [c for c in raw.columns if c != "元行"][:8]
    converted_cols = [
        "元行", "event_date", "start_time", "category", "title",
        "user_id", "user_name", "memo", "取込状態", "既存重複"
    ]
    left = raw[preview_cols]
    right = import_df[[c for c in converted_cols if c in import_df.columns]]
    return pd.merge(left, right, on="元行", how="outer", suffixes=("_元", "_変換後"))


def event_exists(row):
    """同一予定らしいものが既にあるか簡易チェック。"""
    df = fetch_df("""
        SELECT id FROM events
        WHERE event_date=?
          AND IFNULL(start_time, '')=?
          AND title=?
          AND IFNULL(user_id, '')=?
          AND IFNULL(user_name, '')=?
        LIMIT 1
    """, (
        row.get("event_date", ""),
        row.get("start_time", "") or "",
        row.get("title", ""),
        row.get("user_id", "") or "",
        row.get("user_name", "") or "",
    ))
    return not df.empty


def ensure_import_category(category_name):
    """取込データのカテゴリがマスタになければ追加する。"""
    category_name = str(category_name or "その他").strip() or "その他"
    hit = fetch_df("SELECT id FROM categories WHERE category_name=? LIMIT 1", (category_name,))
    if not hit.empty:
        return
    mark = CATEGORY_MARK.get(category_name, "・")
    try:
        execute("""
            INSERT INTO categories
            (category_name, mark, sort_order, is_active, created_at, updated_at)
            VALUES (?, ?, 900, 1, ?, ?)
        """, (category_name, mark, now_text(), now_text()))
    except Exception:
        pass


def upsert_import_user(user_id, user_name):
    """user_id付き利用者が未登録なら、ひだまり帳の利用者マスタにも補助登録する。"""
    user_id = str(user_id or "").strip()
    user_name = str(user_name or "").strip()
    if not user_id or not user_name:
        return
    hit = fetch_df("SELECT id FROM users WHERE user_id=? OR user_name=? LIMIT 1", (user_id, user_name))
    if not hit.empty:
        return
    try:
        execute("""
            INSERT INTO users (user_id, user_name, kana, room_no, note, is_active, created_at)
            VALUES (?, ?, '', '', '予定データ取込から自動追加', 1, ?)
        """, (user_id, user_name, now_text()))
    except Exception:
        pass


def insert_import_event(row):
    """正規化済み1行をeventsへ登録。"""
    ensure_import_category(row.get("category", "その他"))
    upsert_import_user(row.get("user_id", ""), row.get("user_name", ""))
    return execute("""
        INSERT INTO events
        (event_date, category, title, user_id, user_name, staff_name, start_time, end_time, memo, important, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("event_date", ""),
        row.get("category", "その他") or "その他",
        row.get("title", "予定") or "予定",
        row.get("user_id", "") or None,
        row.get("user_name", "") or None,
        row.get("staff_name", "") or None,
        row.get("start_time", "") or None,
        row.get("end_time", "") or None,
        row.get("memo", "") or None,
        int(row.get("important", 0) or 0),
        now_text(),
        now_text(),
    ))


def page_schedule_import():
    st.subheader("予定データ取込")
    st.caption("健康管理アプリで出力した予定候補Excel/CSVを読み込み、内容確認後にひだまり帳の予定へ登録します。")

    uploaded = st.file_uploader(
        "予定候補Excel/CSVを選択",
        type=["xlsx", "xls", "csv"],
        help="健康管理アプリの『予定候補一覧をExcelでダウンロード』『CSV出力』で作成したファイルを読み込めます。"
    )

    if not uploaded:
        st.info("ファイルを選択すると、予定候補を一覧表示します。")
        st.markdown("""
        **推奨列名**  
        `event_date / start_time / end_time / category / title / user_id / user_name / memo / important`  
        日本語列名（予定日、開始時刻、終了時刻、分類、タイトル、利用者ID、利用者名、詳細、重要）でも読み込めます。
        """)
        return

    try:
        raw_df = read_schedule_import_file(uploaded)
    except Exception as e:
        st.error(f"ファイルを読み込めませんでした：{e}")
        return

    import_df = normalize_schedule_import_df(raw_df)
    if import_df.empty:
        st.warning("取込できる予定候補がありません。")
        with st.expander("読み込んだ元データを確認"):
            st.dataframe(raw_df, use_container_width=True, hide_index=True)
        return

    # 重複チェック列を追加
    import_df["既存重複"] = import_df.apply(lambda r: "あり" if event_exists(r) else "", axis=1)
    # 重複ありは初期状態で登録対象から外す
    import_df.loc[import_df["既存重複"] == "あり", "登録する"] = False

    st.success(f"予定候補を {len(import_df)} 件読み込みました。内容を確認してから登録してください。")
    st.caption("上の件数は、元データを予定登録用に変換した後の件数です。空行や日付なし行は除外されます。")

    upload_key = make_upload_key(uploaded)
    editor_key = f"schedule_import_editor_{upload_key}"

    edited = st.data_editor(
        import_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "元行": st.column_config.NumberColumn("元行", disabled=True),
            "登録する": st.column_config.CheckboxColumn("登録する"),
            "event_date": st.column_config.TextColumn("予定日", help="YYYY-MM-DD"),
            "start_time": st.column_config.TextColumn("開始時刻", help="例：10:00"),
            "end_time": st.column_config.TextColumn("終了時刻", help="例：11:00"),
            "category": st.column_config.TextColumn("分類"),
            "title": st.column_config.TextColumn("タイトル"),
            "user_id": st.column_config.TextColumn("利用者ID"),
            "user_name": st.column_config.TextColumn("利用者名"),
            "staff_name": st.column_config.TextColumn("担当"),
            "memo": st.column_config.TextColumn("詳細"),
            "important": st.column_config.CheckboxColumn("重要"),
            "取込状態": st.column_config.TextColumn("状態", disabled=True),
            "既存重複": st.column_config.TextColumn("既存重複", disabled=True),
        },
        key=editor_key,
    )

    selected = edited[edited["登録する"].astype(bool)].copy()
    valid = selected[(selected["event_date"].astype(str).str.strip() != "") & (selected["title"].astype(str).str.strip() != "")].copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("登録対象", len(selected))
    with c2:
        st.metric("登録可能", len(valid))
    with c3:
        st.metric("重複候補", int((edited["既存重複"].astype(str) == "あり").sum()))

    with st.expander("元データと予定候補の対応を確認"):
        compare_df = make_import_compare_df(raw_df, import_df)
        if compare_df.empty:
            st.info("比較できるデータがありません。")
        else:
            st.dataframe(compare_df, use_container_width=True, hide_index=True)
        st.caption("左側が読み込んだ元データ、右側が予定候補へ変換した内容です。元行番号で対応を確認できます。")

    with st.expander("読み込んだ元データだけを確認"):
        st.dataframe(raw_df, use_container_width=True, hide_index=True)

    skip_duplicates = st.checkbox("既存重複ありの行は登録しない", value=True)
    confirm = st.checkbox("内容を確認しました。eventsテーブルへ登録します。")

    if st.button("選択した予定をひだまり帳へ登録", type="primary", use_container_width=True):
        if not confirm:
            st.error("確認チェックを入れてください。")
            return
        if valid.empty:
            st.error("登録可能な行がありません。予定日とタイトルを確認してください。")
            return

        inserted = 0
        skipped = 0
        errors = []
        for idx, row in valid.iterrows():
            try:
                if skip_duplicates and event_exists(row):
                    skipped += 1
                    continue
                insert_import_event(row)
                inserted += 1
            except Exception as e:
                row_no = row.get("元行", idx + 1)
                errors.append(f"元行{row_no}: {e}")

        if inserted:
            st.success(f"ひだまり帳へ {inserted} 件登録しました。")
        if skipped:
            st.info(f"重複候補のため {skipped} 件スキップしました。")
        if errors:
            st.error("一部登録できませんでした。")
            for err in errors[:10]:
                st.write(err)

def page_export():
    st.subheader("Excel出力")

    events = fetch_df("SELECT * FROM events ORDER BY event_date, start_time, id")
    photos = fetch_df("SELECT * FROM event_photos ORDER BY event_id, id")
    files = fetch_df("SELECT * FROM event_files ORDER BY event_id, id")
    categories = fetch_df("SELECT * FROM categories ORDER BY sort_order, category_name")
    users = fetch_df("SELECT * FROM users ORDER BY user_name")
    staff = fetch_df("SELECT * FROM staff ORDER BY staff_name")

    output = Path("hidamari_calendar_export.xlsx")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        events.to_excel(writer, sheet_name="予定", index=False)
        photos.to_excel(writer, sheet_name="写真メモ", index=False)
        files.to_excel(writer, sheet_name="添付ファイル", index=False)
        categories.to_excel(writer, sheet_name="カテゴリマスタ", index=False)
        users.to_excel(writer, sheet_name="利用者マスタ", index=False)
        staff.to_excel(writer, sheet_name="職員マスタ", index=False)

    with open(output, "rb") as f:
        st.download_button(
            label="Excelをダウンロード",
            data=f,
            file_name="hidamari_calendar_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.caption("予定・利用者マスタ・職員マスタをまとめて出力します。")




def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📅", layout="wide")
    init_db()
    add_css()

    st.title("📅 ひだまり帳 Ver1.2.0")
    st.caption("紙の壁カレンダー感覚で、通院・面会・行事・注意事項を一枚で")

    menu = st.sidebar.radio(
        "メニュー",
        [
            "月間カレンダー",
            "今日は何ある",
            "予定登録",
            "予定検索・更新・削除",
            "予定データ取込",
            "写真メモ一覧",
            "Excel・書類ファイル一覧",
            "予定カテゴリ設定",
            "利用者マスタ",
            "職員マスタ",
            "Excel出力",
        ],
    )

    if menu == "月間カレンダー":
        page_calendar()
    elif menu == "今日は何ある":
        page_today()
    elif menu == "予定登録":
        page_event_register()
    elif menu == "予定検索・更新・削除":
        page_event_manage()
    elif menu == "予定データ取込":
        page_schedule_import()
    elif menu == "写真メモ一覧":
        page_photo_notes()
    elif menu == "Excel・書類ファイル一覧":
        page_attached_files()
    elif menu == "予定カテゴリ設定":
        page_category_master()
    elif menu == "利用者マスタ":
        page_master_users()
    elif menu == "職員マスタ":
        page_master_staff()
    elif menu == "Excel出力":
        page_export()


if __name__ == "__main__":
    main()
