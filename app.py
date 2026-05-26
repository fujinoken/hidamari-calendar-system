
# -*- coding: utf-8 -*-
"""
ひだまり現場カレンダー Ver1.3.2
超軽量・単独版
Python + Streamlit + SQLite

起動:
    streamlit run app.py

必要ライブラリ:
    pip install streamlit pandas openpyxl
"""

import sqlite3
import calendar
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st


APP_TITLE = "ひだまり現場カレンダー Ver1.3.2"
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

    conn.commit()
    conn.close()


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


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
CATEGORIES = ["通院", "面会", "行事", "外出", "注意", "申し送り", "夜勤", "その他"]

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
        font-size: 0.82rem;
        line-height: 1.35;
        margin: 3px 0;
        padding: 3px 5px;
        border-radius: 7px;
        background: #f6efe6;
        overflow-wrap: anywhere;
    }
    .important {
        background: #ffe9e0;
        border-left: 4px solid #d65a31;
    }
    .small-note {
        color: #7a6a5b;
        font-size: 0.9rem;
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


def render_calendar(year, month):
    """
    紙の壁カレンダー風レイアウト。
    Streamlitでグリッドが崩れないよう、カレンダー全体を1つのHTMLとして出力する。
    """
    events_by_day = monthly_events(year, month)
    first_weekday, last_day = calendar.monthrange(year, month)  # Monday=0, Sunday=6

    # Sunday start に変換：Sunday=0, Monday=1...
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
            mark = CATEGORY_MARK.get(ev["category"], "・")
            time_part = f'{ev["start_time"]} ' if ev["start_time"] else ""
            if ev["user_name"] and ev["user_id"]:
                user_part = f'／{ev["user_name"]}({ev["user_id"]})'
            elif ev["user_name"]:
                user_part = f'／{ev["user_name"]}'
            else:
                user_part = ""

            imp_cls = " important" if int(ev["important"] or 0) == 1 else ""
            text = html_escape(f'{mark}{time_part}{ev["title"]}{user_part}')
            html.append(f'<div class="event-line{imp_cls}">{text}</div>')

        html.append('</div>')

    html.append('</div>')

    st.markdown("".join(html), unsafe_allow_html=True)


# -----------------------------
# Pages
# -----------------------------
def page_calendar():
    today = date.today()
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        year = st.number_input("年", min_value=2020, max_value=2100, value=today.year, step=1)
    with col2:
        month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1)
    with col3:
        st.markdown('<div class="small-note">紙の壁カレンダーのように、月全体の流れを一目で確認する画面です。</div>', unsafe_allow_html=True)

    render_calendar(int(year), int(month))


def page_event_register():
    st.subheader("予定登録")

    user_map = user_display_map()
    users = list(user_map.keys())
    staff = [""] + get_active_staff()

    with st.form("event_register_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            event_date = st.date_input("日付", value=date.today())
            category = st.selectbox("カテゴリ", CATEGORIES)
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
        start = st.date_input("開始日", value=date.today().replace(day=1))
    with c2:
        end = st.date_input("終了日", value=date.today())
    with c3:
        category_filter = st.selectbox("カテゴリ絞り込み", ["すべて"] + CATEGORIES)

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
            new_category = st.selectbox("カテゴリ", CATEGORIES, index=CATEGORIES.index(target["category"]) if target["category"] in CATEGORIES else 0)
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


def page_export():
    st.subheader("Excel出力")

    events = fetch_df("SELECT * FROM events ORDER BY event_date, start_time, id")
    photos = fetch_df("SELECT * FROM event_photos ORDER BY event_id, id")
    files = fetch_df("SELECT * FROM event_files ORDER BY event_id, id")
    users = fetch_df("SELECT * FROM users ORDER BY user_name")
    staff = fetch_df("SELECT * FROM staff ORDER BY staff_name")

    output = Path("hidamari_calendar_export.xlsx")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        events.to_excel(writer, sheet_name="予定", index=False)
        photos.to_excel(writer, sheet_name="写真メモ", index=False)
        files.to_excel(writer, sheet_name="添付ファイル", index=False)
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


def page_about():
    st.subheader("このアプリについて")
    st.markdown("""
    **ひだまり現場カレンダー Ver1.3.2** は、  
    グループホームなど小規模施設向けの、超軽量な予定共有アプリです。

    目的は、高機能なスケジュール管理ではなく、  
    **紙の壁カレンダーの見やすさを残したまま、通院・面会・行事・注意事項を共有すること**です。

    ### Ver1.0でできること
    - 月間カレンダー表示
    - 予定登録
    - 予定の検索・更新・削除
    - 写真メモの登録・一覧表示
    - Excel・CSVファイルの添付・一覧表示・ダウンロード
    - 利用者IDによる予定紐づけ
    - 利用者マスタ
    - 職員マスタ
    - Excel出力

    ### 今は入れていないこと
    - 健康チェックアプリとの自動連携
    - Googleカレンダー連携
    - シフト自動作成
    - AI分析

    まずは単独で安定運用し、必要になったら健康管理アプリと連携する設計です。
    """)


# -----------------------------
# Main
# -----------------------------
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📅", layout="wide")
    init_db()
    add_css()

    st.title("📅 ひだまり現場カレンダー Ver1.3.2")
    st.caption("紙の壁カレンダー感覚で、通院・面会・行事・注意事項を一枚で見るための超軽量アプリ")

    menu = st.sidebar.radio(
        "メニュー",
        [
            "月間カレンダー",
            "予定登録",
            "予定検索・更新・削除",
            "写真メモ一覧",
            "Excel・書類ファイル一覧",
            "利用者マスタ",
            "職員マスタ",
            "Excel出力",
            "このアプリについて",
        ],
    )

    if menu == "月間カレンダー":
        page_calendar()
    elif menu == "予定登録":
        page_event_register()
    elif menu == "予定検索・更新・削除":
        page_event_manage()
    elif menu == "写真メモ一覧":
        page_photo_notes()
    elif menu == "Excel・書類ファイル一覧":
        page_attached_files()
    elif menu == "利用者マスタ":
        page_master_users()
    elif menu == "職員マスタ":
        page_master_staff()
    elif menu == "Excel出力":
        page_export()
    elif menu == "このアプリについて":
        page_about()


if __name__ == "__main__":
    main()
