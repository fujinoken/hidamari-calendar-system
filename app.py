
# -*- coding: utf-8 -*-
"""
ひだまり帳 Ver1.4.7
PostgreSQL永続化版
Python + Streamlit + PostgreSQL

起動:
    streamlit run app.py

必要ライブラリ:
    pip install streamlit pandas openpyxl psycopg2-binary reportlab requests
"""

import calendar
import re
import hashlib
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

import report_service
from king_of_time_auto_schedule import build_auto_schedule_export

from config import (
    AI_SHIFT_RULE_VERSION,
    APP_TITLE,
    CATEGORY_MARK,
    DEFAULT_CATEGORIES,
    FILE_DIR,
    SHIFT_EDITOR_OPTIONS,
    SHIFT_KINDS,
    UPLOAD_DIR,
)
from db import (
    DB_INTEGRITY_ERROR,
    execute,
    execute_many,
    execute_transaction,
    fetch_df,
    init_db_once,
    now_text,
    today_jst,
)
from report_service import (
    OPENPYXL_AVAILABLE,
    REPORTLAB_AVAILABLE,
    build_king_of_time_clock_export as report_build_king_of_time_clock_export,
    make_calendar_pdf as report_make_calendar_pdf,
    make_king_of_time_shift_csv as report_make_king_of_time_shift_csv,
    make_staff_shift_excel as report_make_staff_shift_excel,
    make_staff_shift_pdf as report_make_staff_shift_pdf,
)
from storage_service import (
    REQUESTS_AVAILABLE,
    check_storage_bucket_access,
    clear_event_files_cache,
    clear_event_photos_cache,
    delete_saved_file,
    get_event_files,
    get_event_photos,
    get_supabase_storage_bucket,
    get_supabase_storage_key,
    get_supabase_url,
    is_storage_file,
    render_saved_download_button,
    render_saved_image,
    save_uploaded_files,
    save_uploaded_photos,
    storage_is_configured,
    storage_object_exists,
)
from ui_styles import add_css

UPLOAD_DIR.mkdir(exist_ok=True)
FILE_DIR.mkdir(exist_ok=True)


def count_query(sql, params=()):
    """COUNT系SQLを安全にintで返す。"""
    try:
        df = fetch_df(sql, params)
        if df is None or df.shape[0] == 0 or df.shape[1] == 0:
            return 0
        value = df.iat[0, 0]
        if pd.isna(value):
            return 0
        return int(value or 0)
    except Exception as e:
        st.warning(f"件数取得に失敗しました：{e}")
        return 0




EVENT_LIST_COLUMNS = """
    id, event_date, category, title, user_id, user_name, staff_name,
    start_time, end_time, important, created_at, updated_at
"""

EVENT_TODAY_COLUMNS = """
    id, event_date, category, title, user_id, user_name, staff_name,
    start_time, end_time, memo, important, created_at, updated_at
"""

EVENT_DETAIL_COLUMNS = """
    id, event_date, category, title, user_id, user_name, staff_name,
    start_time, end_time, memo, important, created_at, updated_at
"""

USER_COLUMNS = "id, user_id, user_name, kana, room_no, note, is_active, created_at"
STAFF_COLUMNS = "id, staff_name, staff_code, role, is_active, created_at"
CATEGORY_COLUMNS = "id, category_name, mark, sort_order, is_active, created_at, updated_at"
SHIFT_COLUMNS = """
    id, shift_date, staff_name, shift_kind, start_time, end_time,
    next_day, memo, created_at, updated_at
"""

DAY_STAFFING_SHIFT_KINDS = ["日勤"]
DAY_LIMIT_SHIFT_KINDS = ["日勤"]
NIGHT_LIMIT_SHIFT_KINDS = ["夜勤"]
MANAGEMENT_SHIFT_KINDS = ["管", "管理業務"]
WORKDAY_SHIFT_KINDS = ["日勤", "管", "管理業務", "夜勤", "夜勤明け"]


@dataclass(frozen=True)
class ShiftSaveResult:
    status: str
    count: int = 0
    message: str = ""
    changed_cells: int = 0


class ShiftUpdateBlockedError(RuntimeError):
    """確定済み、または確定状態不明の月に対する更新拒否。"""


SHIFT_SAVE_SAVED = "saved"
SHIFT_SAVE_NO_CHANGE = "no_change"
SHIFT_SAVE_DUPLICATE = "duplicate"
SHIFT_SAVE_BLOCKED = "blocked"


def show_shift_save_result(result, success_message=None):
    """保存結果を誤解のない日本語メッセージで表示する。"""
    if not isinstance(result, ShiftSaveResult):
        st.success(success_message or "勤務データを保存しました。")
        return
    message = success_message or result.message
    if result.status == SHIFT_SAVE_SAVED:
        st.success(message or "勤務データを保存しました。")
    elif result.status == SHIFT_SAVE_DUPLICATE:
        st.warning(message or "重複しているため保存しませんでした。")
    elif result.status == SHIFT_SAVE_BLOCKED:
        st.error(message or "重大な問題があるため処理できませんでした。")
    else:
        st.info(message or "変更はありませんでした。")


def queue_shift_save_result(result):
    if isinstance(result, ShiftSaveResult):
        st.session_state["shift_save_flash"] = {
            "status": result.status,
            "count": result.count,
            "message": result.message,
            "changed_cells": result.changed_cells,
        }


def render_queued_shift_save_result():
    value = st.session_state.pop("shift_save_flash", None)
    if not value:
        return
    show_shift_save_result(ShiftSaveResult(**value))

# 帳票生成ロジックやレイアウトを変更した場合は、対応する値を更新する。
STAFF_SHIFT_PDF_CACHE_VERSION = "1"
STAFF_SHIFT_EXCEL_CACHE_VERSION = "1"
SHIFT_CALENDAR_PDF_CACHE_VERSION = "1"
KING_OF_TIME_CACHE_VERSION = "1"


def note_perf(label, started_at):
    elapsed = time.perf_counter() - started_at
    st.session_state.setdefault("perf_log", [])
    st.session_state["perf_log"] = (
        [{"label": label, "sec": elapsed, "at": datetime.now().strftime("%H:%M:%S")}]
        + st.session_state["perf_log"]
    )[:8]
    return elapsed


def show_perf_log():
    rows = st.session_state.get("perf_log", [])
    if not rows:
        return
    with st.sidebar.expander("処理時間", expanded=False):
        for row in rows:
            st.caption(f"{row['at']} {row['label']}: {row['sec']:.2f}s")


def clear_event_caches():
    for fn in (monthly_events, events_by_date, get_event_by_id, get_attachment_counts, get_staff_event_counts_month):
        try:
            fn.clear()
        except Exception:
            pass


def clear_master_caches():
    for fn in (get_active_users, get_active_staff, get_staff_code_map, get_staff_key_map, get_categories, get_category_mark_map):
        try:
            fn.clear()
        except Exception:
            pass


def clear_shift_caches():
    for fn in (get_shift_month_status, get_staff_shifts_month, get_staff_shift_limits):
        try:
            fn.clear()
        except Exception:
            pass


@st.cache_data(ttl=60, show_spinner=False)
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


@st.cache_data(ttl=60, show_spinner=False)
def get_active_staff():
    """
    有効な職員名を取得する。
    「職員 A」と「職員A」のような空白違いは同一職員として扱い、
    シフト表に重複行が出ないようにする。
    """
    df = fetch_df("SELECT staff_name FROM staff WHERE is_active=1 ORDER BY staff_name")
    if df.empty:
        return []

    names = []
    seen = set()
    for raw_name in df["staff_name"].dropna().astype(str).tolist():
        name = normalize_staff_name(raw_name)
        if not name or name in seen:
            continue
        seen.add(name)
        names.append(name)
    return names


@st.cache_data(ttl=60, show_spinner=False)
def get_staff_code_map(active_only=True):
    """
    KING OF TIME CSV出力用に、職員名（正規化後）と従業員コードの対応表を取得する。
    空白違いの職員名がある場合は、先に見つかった有効な従業員コードを優先する。
    """
    where = "WHERE is_active=1" if active_only else ""
    try:
        df = fetch_df(f"""
            SELECT staff_name, staff_code, is_active
            FROM staff
            {where}
            ORDER BY is_active DESC, staff_name
        """)
    except Exception:
        return {}

    mapping = {}
    if df is None or df.empty:
        return mapping

    for _, r in df.iterrows():
        staff_name = normalize_staff_name(r.get("staff_name", ""))
        staff_code = str(r.get("staff_code") or "").strip()
        if not staff_name:
            continue
        if staff_name not in mapping or (not mapping.get(staff_name) and staff_code):
            mapping[staff_name] = staff_code
    return mapping


@st.cache_data(ttl=60, show_spinner=False)
def get_staff_key_map(active_only=True):
    """職員名（正規化後）から内部キー staff.id への対応表を返す。"""
    where = "WHERE is_active=1" if active_only else ""
    df = fetch_df(f"SELECT id, staff_name, is_active FROM staff {where} ORDER BY is_active DESC, id")
    mapping = {}
    if df is None or df.empty:
        return mapping
    for _, row in df.iterrows():
        staff_name = normalize_staff_name(row.get("staff_name", ""))
        if staff_name and staff_name not in mapping:
            mapping[staff_name] = int(row["id"])
    return mapping


def _get_king_of_time_month_staff(year, month):
    """対象月のシフトに存在する職員を staff.id と表示名で返す。"""
    shift_df = get_staff_shifts_month(year, month)
    if shift_df is None or shift_df.empty:
        return []
    staff_key_map = get_staff_key_map(active_only=False)
    result = []
    seen = set()
    for raw_name in shift_df["staff_name"].dropna().astype(str):
        staff_name = normalize_staff_name(raw_name)
        staff_id = staff_key_map.get(staff_name)
        if staff_id is None or staff_id in seen:
            continue
        seen.add(staff_id)
        result.append((staff_id, staff_name))
    return sorted(result, key=lambda item: item[1])


def get_king_of_time_export_staff(year, month):
    """対象月の確定済みシフトに存在する職員を返す。"""
    if not shift_month_is_confirmed(year, month):
        return []
    return _get_king_of_time_month_staff(year, month)


def get_missing_staff_code_names(staff_names):
    """KING OF TIME従業員コードが未登録の職員名一覧を返す。"""
    code_map = get_staff_code_map(active_only=True)
    missing = []
    seen = set()
    for name in staff_names or []:
        staff_name = normalize_staff_name(name)
        if not staff_name or staff_name in seen:
            continue
        seen.add(staff_name)
        if not str(code_map.get(staff_name, "") or "").strip():
            missing.append(staff_name)
    return missing


# -----------------------------
# UI helpers
# -----------------------------
@st.cache_data(ttl=60, show_spinner=False)
def get_categories(active_only=True):
    where = "WHERE is_active=1" if active_only else ""
    df = fetch_df(f"""
        SELECT category_name, mark, sort_order, is_active
        FROM categories
        {where}
        ORDER BY sort_order, category_name
    """)
    if df.empty and active_only:
        categories = list(DEFAULT_CATEGORIES)
    else:
        categories = df["category_name"].dropna().astype(str).tolist()
    if active_only:
        categories = [category for category in categories if category != "夜勤"]
        if "面接" not in categories:
            insert_at = categories.index("申し送り") + 1 if "申し送り" in categories else len(categories)
            categories.insert(insert_at, "面接")
    return categories


@st.cache_data(ttl=60, show_spinner=False)
def get_category_mark_map():
    df = fetch_df("""
        SELECT category_name, mark
        FROM categories
        WHERE is_active=1
    """)
    mapping = {}
    if not df.empty:
        for _, r in df.iterrows():
            name = str(r["category_name"] or "").strip()
            mark = str(r["mark"] or "").strip()
            if name and mark:
                mapping[name] = mark
    return mapping


def get_category_mark(category_name):
    mark_map = get_category_mark_map()
    category_name = str(category_name or "").strip()
    return mark_map.get(category_name, CATEGORY_MARK.get(category_name, "・"))


@st.cache_data(ttl=60, show_spinner=False)
def monthly_events(year, month):
    start = f"{year}-{month:02d}-01"
    last_day = calendar.monthrange(year, month)[1]
    end = f"{year}-{month:02d}-{last_day:02d}"

    df = fetch_df(f"""
        SELECT {EVENT_LIST_COLUMNS}
        FROM events
        WHERE event_date BETWEEN ? AND ?
        ORDER BY event_date, start_time, category, id
    """, (start, end))

    events_by_day = {}
    for _, row in df.iterrows():
        events_by_day.setdefault(row["event_date"], []).append(row)
    return events_by_day



@st.cache_data(ttl=60, show_spinner=False)
def events_by_date(target_date):
    """指定日の予定一覧を取得する。"""
    if hasattr(target_date, "strftime"):
        target_date = target_date.strftime("%Y-%m-%d")
    return fetch_df(f"""
        SELECT {EVENT_TODAY_COLUMNS}
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
        start_text = format_time_for_display(ev.get("start_time"))
        label_time = f" {start_text}" if start_text else ""
        label_user = f"／{ev['user_name']}" if ev["user_name"] else ""
        important = "⚠️ " if int(ev["important"] or 0) == 1 else ""
        label = f"{label_date}{label_time} {important}{mark}{ev['title']}{label_user}".strip()

        with cols[i % 3]:
            if st.button(label, key=f"event_btn_{ev['id']}", use_container_width=True):
                set_selected_event(int(ev["id"]))



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


def format_time_for_display(value):
    """DBやDataFrameから来た時刻値を画面表示用の文字列に安全変換する。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if value == "":
        return ""
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%H:%M")
        except Exception:
            return str(value)
    return str(value).strip()


def format_time_range_for_display(start_value, end_value, empty_text="時刻未設定", separator=" - "):
    start_text = format_time_for_display(start_value)
    end_text = format_time_for_display(end_value)
    time_parts = [x for x in [start_text, end_text] if x]
    return separator.join(time_parts) if time_parts else empty_text


@st.cache_data(ttl=20, show_spinner=False)
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
        time_text = format_time_for_display(ev.get("start_time")) or "時間未定"
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


@st.cache_data(ttl=60, show_spinner=False)
def get_event_by_id(event_id):
    if not event_id:
        return pd.DataFrame()
    return fetch_df(f"SELECT {EVENT_DETAIL_COLUMNS} FROM events WHERE id=? LIMIT 1", (int(event_id),))


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
        st.write(f"**時間**：{format_time_range_for_display(ev.get('start_time'), ev.get('end_time'), empty_text='')}")
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
                render_saved_image(
                    p["file_path"],
                    caption=p["photo_memo"] or p["file_name"],
                    use_container_width=True,
                )

    files = get_event_files(event_id)
    if not files.empty:
        st.write("**Excel・書類ファイル**")
        for _, frow in files.iterrows():
            st.write(f"📎 {frow['file_name']}　{frow['file_memo'] or ''}")
            render_saved_download_button(
                "ダウンロード",
                frow["file_path"],
                frow["file_name"],
                key=f"detail_download_{frow['id']}",
            )

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("詳細を閉じる", use_container_width=True):
            st.session_state["selected_calendar_event_id"] = None
            st.rerun()
    with col_b:
        st.caption("編集・削除は「予定検索・更新・削除」メニューで行います。")



def clear_last_saved_event_state():
    """保存後確認エリア用の一時情報を消す。"""
    for key in [
        "last_saved_event_id",
        "last_saved_photo_saved",
        "last_saved_photo_failed",
        "last_saved_file_saved",
        "last_saved_file_failed",
        "last_saved_had_photos",
        "last_saved_had_files",
    ]:
        st.session_state.pop(key, None)


def render_saved_event_confirmation(event_id=None):
    """
    予定登録直後に、保存した予定・写真メモ・添付ファイルの紐づきをその場で確認する。
    入力ミスやStorage保存失敗にすぐ気づけるようにするための確認パネル。
    """
    if event_id is None:
        event_id = st.session_state.get("last_saved_event_id")
    if not event_id:
        return

    ev_df = get_event_by_id(event_id)
    if ev_df.empty:
        st.warning("保存後確認：保存した予定が見つかりません。")
        return

    ev = ev_df.iloc[0]
    photos = get_event_photos(event_id)
    files = get_event_files(event_id)

    photo_saved = int(st.session_state.get("last_saved_photo_saved", len(photos)) or 0)
    photo_failed = int(st.session_state.get("last_saved_photo_failed", 0) or 0)
    file_saved = int(st.session_state.get("last_saved_file_saved", len(files)) or 0)
    file_failed = int(st.session_state.get("last_saved_file_failed", 0) or 0)
    had_photos = bool(st.session_state.get("last_saved_had_photos", False))
    had_files = bool(st.session_state.get("last_saved_had_files", False))

    st.markdown("---")
    st.subheader("保存後の確認")
    st.caption("いま保存した予定の内容と、写真・添付ファイルの紐づき状態を確認できます。")

    if photo_failed or file_failed:
        st.warning("予定本体は保存されていますが、写真メモまたは添付ファイルで保存失敗があります。")
    else:
        st.success("予定本体と、保存できた写真・添付ファイルの紐づきを確認できます。")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("予定ID", int(event_id))
    with c2:
        st.metric("写真紐づき", len(photos))
    with c3:
        st.metric("添付紐づき", len(files))
    with c4:
        ng_count = int(photo_failed + file_failed)
        st.metric("保存失敗", ng_count)

    if had_photos:
        if photo_failed:
            st.warning(f"写真メモ：{photo_saved}件保存、{photo_failed}件失敗")
        else:
            st.info(f"写真メモ：{photo_saved}件保存")
    if had_files:
        if file_failed:
            st.warning(f"Excel・書類：{file_saved}件保存、{file_failed}件失敗")
        else:
            st.info(f"Excel・書類：{file_saved}件保存")

    st.markdown("#### 保存した予定の詳細")
    d1, d2, d3 = st.columns(3)
    with d1:
        st.write(f"**日付**：{ev['event_date']}")
        st.write(f"**カテゴリ**：{ev['category']}")
        st.write(f"**タイトル**：{ev['title']}")
    with d2:
        st.write(f"**利用者**：{ev['user_name'] or ''}")
        st.caption(f"利用者ID：{ev['user_id'] or ''}")
        st.write(f"**担当**：{ev['staff_name'] or ''}")
    with d3:
        st.write(f"**時間**：{format_time_range_for_display(ev.get('start_time'), ev.get('end_time'), empty_text='')}")
        st.write(f"**重要**：{'あり' if int(ev['important'] or 0) == 1 else 'なし'}")
        st.caption(f"登録日時：{ev['created_at']}")

    st.write("**メモ**")
    st.info(ev["memo"] or "メモはありません。")

    st.markdown("#### 写真メモの確認")
    if photos.empty:
        if had_photos:
            st.warning("写真を選択していましたが、DB上の写真メモ紐づきは0件です。Storage保存失敗の可能性があります。")
        else:
            st.info("写真メモは添付されていません。")
    else:
        st.caption(f"DBに紐づいた写真メモ：{len(photos)}件")
        cols = st.columns(3)
        for i, (_, p) in enumerate(photos.iterrows()):
            with cols[i % 3]:
                render_saved_image(
                    p["file_path"],
                    caption=p["photo_memo"] or p["file_name"],
                    use_container_width=True,
                )
                st.caption(f"写真ID：{p['id']}")

    st.markdown("#### Excel・書類ファイルの確認")
    if files.empty:
        if had_files:
            st.warning("ファイルを選択していましたが、DB上の添付ファイル紐づきは0件です。Storage保存失敗の可能性があります。")
        else:
            st.info("Excel・書類ファイルは添付されていません。")
    else:
        st.caption(f"DBに紐づいた添付ファイル：{len(files)}件")
        for _, frow in files.iterrows():
            st.write(f"📎 **{frow['file_name']}**　{frow['file_memo'] or ''}")
            st.caption(f"ファイルID：{frow['id']} / 保存先：{frow['file_path']}")
            render_saved_download_button(
                "ダウンロードして確認",
                frow["file_path"],
                frow["file_name"],
                key=f"saved_confirm_download_{frow['id']}",
            )

    cc1, cc2 = st.columns(2)
    with cc1:
        if st.button("この予定を下の詳細パネルでも表示", use_container_width=True):
            set_selected_event(int(event_id))
            st.rerun()
    with cc2:
        if st.button("保存後確認を閉じる", use_container_width=True):
            clear_last_saved_event_state()
            st.rerun()





def get_weekday_label(target_date):
    labels = ["月", "火", "水", "木", "金", "土", "日"]
    return labels[target_date.weekday()]


def format_event_for_calendar(event):
    mark = get_category_mark(event.get("category", ""))
    time_text = format_time_for_display(event.get("start_time"))
    title = str(event.get("title") or "").strip()
    category = str(event.get("category") or "").strip()
    important = "!" if int(event.get("important") or 0) == 1 else ""
    head = f"{time_text} " if time_text else ""
    cat = f"［{category}］" if category else ""
    return f"{important}{mark} {head}{cat}{title}".strip()


def render_day_cell(target_date, day_events):
    date_key = target_date.strftime("%Y-%m-%d")
    is_selected = st.session_state.get("selected_calendar_date") == date_key
    is_today = target_date == today_jst()
    weekday = target_date.weekday()
    classes = ["hm-day-card"]
    if weekday == 6:
        classes.append("hm-sunday")
    if weekday == 5:
        classes.append("hm-saturday")
    if is_today:
        classes.append("hm-today")
    if is_selected:
        classes.append("hm-selected")
    if day_events:
        classes.append("hm-has-events")

    event_lines = []
    for ev in day_events[:3]:
        event_lines.append(
            f'<div class="hm-event-pill">{html_escape(format_event_for_calendar(ev))}</div>'
        )
    if len(day_events) > 3:
        event_lines.append(f'<div class="hm-more-events">ほか {len(day_events) - 3} 件</div>')

    count_badge = f'<span class="hm-event-count">{len(day_events)}件</span>' if day_events else ""
    html = (
        f'<div class="{" ".join(classes)}">'
        f'<div class="hm-day-top"><span class="hm-day-number">{target_date.day}</span>'
        f'<span class="hm-weekday">{get_weekday_label(target_date)}</span>{count_badge}</div>'
        f'<div class="hm-day-events">{"".join(event_lines)}</div>'
        f'</div>'
    )
    st.markdown(html, unsafe_allow_html=True)
    if st.button("この日を選択", key=f"select_day_{date_key}", use_container_width=True):
        st.session_state["selected_calendar_date"] = date_key
        st.session_state["editing_event_id"] = None
        st.session_state["pending_delete_event_id"] = None
        st.rerun()


def render_month_calendar(year, month, events=None):
    events = events if events is not None else monthly_events(year, month)
    first_weekday, last_day = calendar.monthrange(year, month)
    start_col = (first_weekday + 1) % 7
    days = [None] * start_col
    days.extend(date(year, month, day) for day in range(1, last_day + 1))
    while len(days) % 7 != 0:
        days.append(None)

    st.markdown(f'<div class="hm-calendar-title">{year}年 {month}月</div>', unsafe_allow_html=True)
    header_cols = st.columns(7)
    for idx, label in enumerate(["日", "月", "火", "水", "木", "金", "土"]):
        cls = "hm-cal-head hm-sunday-text" if idx == 0 else "hm-cal-head hm-saturday-text" if idx == 6 else "hm-cal-head"
        header_cols[idx].markdown(f'<div class="{cls}">{label}</div>', unsafe_allow_html=True)

    for week_start in range(0, len(days), 7):
        cols = st.columns(7)
        for idx, target_date in enumerate(days[week_start:week_start + 7]):
            with cols[idx]:
                if target_date is None:
                    st.markdown('<div class="hm-day-card hm-blank"></div>', unsafe_allow_html=True)
                else:
                    key = target_date.strftime("%Y-%m-%d")
                    render_day_cell(target_date, events.get(key, []))


def _event_user_label(user_map, event_row):
    user_id = str(event_row.get("user_id") or "")
    user_name = str(event_row.get("user_name") or "")
    if user_id and user_name:
        expected = f"{user_name}（ID:{user_id}）"
        if expected in user_map:
            return expected
    for label, pair in user_map.items():
        if pair[0] == user_id or pair[1] == user_name:
            return label
    return ""


def _event_date_value(value, fallback_date):
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except Exception:
        return fallback_date


def render_selected_day_events(selected_date):
    st.markdown("### 選択日の予定詳細")
    date_text = selected_date.strftime("%Y-%m-%d")
    st.caption(f"{date_text}（{get_weekday_label(selected_date)}）")
    df = events_by_date(date_text)
    if df.empty:
        st.info("この日の予定はまだありません。下のフォームから登録できます。")
        return

    for _, ev in df.iterrows():
        event_id = int(ev["id"])
        important = "重要 " if int(ev.get("important") or 0) == 1 else ""
        time_text = format_time_range_for_display(ev.get("start_time"), ev.get("end_time"))
        title = html_escape(ev.get("title") or "")
        category = html_escape(ev.get("category") or "")
        memo = html_escape(ev.get("memo") or "")
        staff = html_escape(ev.get("staff_name") or "")
        user = html_escape(ev.get("user_name") or "")
        mark = get_category_mark(ev.get("category"))

        st.markdown(
            f'<div class="hm-detail-card">'
            f'<div class="hm-detail-title">{important}{mark} {title}</div>'
            f'<div class="hm-detail-meta">{time_text}　{category}'
            f'{"　利用者: " + user if user else ""}'
            f'{"　担当: " + staff if staff else ""}</div>'
            f'<div class="hm-detail-memo">{memo if memo else "メモなし"}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns([1, 1, 4])
        with c1:
            if st.button("編集", key=f"edit_event_{event_id}", use_container_width=True):
                st.session_state["editing_event_id"] = event_id
                st.session_state["pending_delete_event_id"] = None
                st.rerun()
        with c2:
            if st.button("削除", key=f"ask_delete_event_{event_id}", use_container_width=True):
                st.session_state["pending_delete_event_id"] = event_id
                st.rerun()
        if st.session_state.get("pending_delete_event_id") == event_id:
            st.warning(f"予定ID {event_id} を削除します。写真メモ・添付ファイルも一緒に削除されます。")
            d1, d2 = st.columns(2)
            with d1:
                if st.button("削除を確定", key=f"confirm_delete_event_{event_id}", type="primary", use_container_width=True):
                    photos = get_event_photos(event_id)
                    for _, photo in photos.iterrows():
                        delete_saved_file(photo["file_path"])
                    files = get_event_files(event_id)
                    for _, file_row in files.iterrows():
                        delete_saved_file(file_row["file_path"])
                    execute("DELETE FROM event_photos WHERE event_id=?", (event_id,))
                    execute("DELETE FROM event_files WHERE event_id=?", (event_id,))
                    execute("DELETE FROM events WHERE id=?", (event_id,))
                    clear_event_photos_cache()
                    clear_event_files_cache()
                    clear_event_caches()
                    st.session_state["pending_delete_event_id"] = None
                    st.session_state["editing_event_id"] = None
                    st.success("予定を削除しました。")
                    st.rerun()
            with d2:
                if st.button("キャンセル", key=f"cancel_delete_event_{event_id}", use_container_width=True):
                    st.session_state["pending_delete_event_id"] = None
                    st.rerun()


def render_event_form(selected_date, editing_event_id=None):
    editing_event = None
    if editing_event_id:
        edit_df = get_event_by_id(editing_event_id)
        if not edit_df.empty:
            editing_event = edit_df.iloc[0]
        else:
            st.session_state["editing_event_id"] = None
            st.warning("編集中の予定が見つかりません。")

    st.markdown("### 選択日の予定登録・編集フォーム")
    if editing_event is not None:
        st.info(f"予定ID {int(editing_event['id'])} を編集中です。")

    user_map = user_display_map()
    users = list(user_map.keys())
    staff = [""] + get_active_staff()
    category_options = get_categories()
    current_category = str(editing_event.get("category") or "") if editing_event is not None else ""
    if current_category and current_category != "夜勤" and current_category not in category_options:
        category_options = [current_category] + category_options

    default_date = _event_date_value(editing_event.get("event_date"), selected_date) if editing_event is not None else selected_date
    default_category = editing_event.get("category") if editing_event is not None else (category_options[0] if category_options else "")
    default_staff = editing_event.get("staff_name") if editing_event is not None else ""
    staff_index = staff.index(default_staff) if default_staff in staff else 0
    user_label = _event_user_label(user_map, editing_event) if editing_event is not None else ""
    user_index = users.index(user_label) if user_label in users else 0
    form_key = f"selected_day_event_form_{selected_date.strftime('%Y%m%d')}_{editing_event_id or 'new'}"

    with st.form(form_key, clear_on_submit=editing_event is None):
        c1, c2, c3 = st.columns(3)
        with c1:
            event_date = st.date_input("日付", value=default_date, key=f"{form_key}_date")
            category = st.selectbox(
                "種別",
                category_options,
                index=category_options.index(default_category) if default_category in category_options else 0,
                key=f"{form_key}_category",
            )
        with c2:
            start_time = st.text_input("開始時刻", value=format_time_for_display(editing_event.get("start_time")) if editing_event is not None else "", placeholder="例：10:00")
            end_time = st.text_input("終了時刻", value=format_time_for_display(editing_event.get("end_time")) if editing_event is not None else "", placeholder="例：11:00")
        with c3:
            user_name = st.selectbox("利用者", users, index=user_index, key=f"{form_key}_user")
            staff_name = st.selectbox("担当職員", staff, index=staff_index, key=f"{form_key}_staff")

        title = st.text_input("予定タイトル", value=str(editing_event.get("title") or "") if editing_event is not None else "", placeholder="例：内科受診、家族面会、外出支援")
        memo = st.text_area("メモ", value=str(editing_event.get("memo") or "") if editing_event is not None else "", placeholder="持ち物、注意点、申し送りなど")
        important = st.checkbox("重要マーク", value=bool(editing_event.get("important")) if editing_event is not None else False)

        uploaded_photos = uploaded_files = []
        photo_memo = file_memo = ""
        if editing_event is None:
            uploaded_photos = st.file_uploader("写真メモ（複数可）", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key=f"{form_key}_photos")
            photo_memo = st.text_input("写真メモ補足", key=f"{form_key}_photo_memo")
            uploaded_files = st.file_uploader("Excel・書類ファイル（複数可）", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key=f"{form_key}_files")
            file_memo = st.text_input("Excel・書類メモ補足", key=f"{form_key}_file_memo")
        else:
            st.caption("写真メモ・添付ファイルの追加や削除は、従来の予定検索・更新・削除画面でも行えます。")

        submit_label = "予定を更新" if editing_event is not None else "予定を登録"
        submitted = st.form_submit_button(submit_label, type="primary", use_container_width=True)

    if not submitted:
        if editing_event is not None and st.button("編集をやめる", key=f"cancel_edit_{editing_event_id}", use_container_width=True):
            st.session_state["editing_event_id"] = None
            st.rerun()
        return

    if not title.strip():
        st.error("予定タイトルを入力してください。")
        return

    selected_user_id, selected_user_name = user_map.get(user_name, ("", ""))
    if editing_event is not None:
        execute("""
            UPDATE events
            SET event_date=?, category=?, title=?, user_id=?, user_name=?, staff_name=?,
                start_time=?, end_time=?, memo=?, important=?, updated_at=?
            WHERE id=?
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
            int(editing_event_id),
        ))
        clear_event_caches()
        st.session_state["selected_calendar_date"] = event_date.strftime("%Y-%m-%d")
        st.session_state["editing_event_id"] = None
        st.success("予定を更新しました。")
        st.rerun()

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
    photo_saved, photo_failed = save_uploaded_photos(event_id, uploaded_photos, photo_memo)
    file_saved, file_failed = save_uploaded_files(event_id, uploaded_files, file_memo)
    clear_event_caches()
    st.session_state["selected_calendar_date"] = event_date.strftime("%Y-%m-%d")
    st.session_state["selected_calendar_event_id"] = int(event_id)
    st.session_state["last_saved_event_id"] = int(event_id)
    st.session_state["last_saved_photo_saved"] = int(photo_saved)
    st.session_state["last_saved_photo_failed"] = int(photo_failed)
    st.session_state["last_saved_file_saved"] = int(file_saved)
    st.session_state["last_saved_file_failed"] = int(file_failed)
    st.session_state["last_saved_had_photos"] = bool(uploaded_photos)
    st.session_state["last_saved_had_files"] = bool(uploaded_files)
    if photo_failed or file_failed:
        st.warning("予定は保存しましたが、一部の写真メモまたは添付ファイルの保存に失敗しました。")
    else:
        st.success("予定を保存しました。")
    st.rerun()


def render_calendar(year, month):
    render_month_calendar(year, month, monthly_events(year, month))


def page_calendar():
    today = today_jst()
    st.subheader("月間カレンダー")
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        year = st.number_input("年", min_value=2020, max_value=2100, value=today.year, step=1, key="calendar_year")
    with c2:
        month = st.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="calendar_month")
    with c3:
        st.caption("日付を選ぶと、下にその日の予定詳細と登録フォームが表示されます。")

    default_date = date(int(year), int(month), min(today.day, calendar.monthrange(int(year), int(month))[1]))
    selected_text = st.session_state.get("selected_calendar_date")
    try:
        selected_date = datetime.strptime(selected_text, "%Y-%m-%d").date() if selected_text else default_date
    except Exception:
        selected_date = default_date
    if selected_date.year != int(year) or selected_date.month != int(month):
        selected_date = default_date
        st.session_state["selected_calendar_date"] = selected_date.strftime("%Y-%m-%d")

    render_calendar(int(year), int(month))
    st.markdown("---")
    render_selected_day_events(selected_date)
    st.markdown("---")
    render_event_form(selected_date, st.session_state.get("editing_event_id"))
    render_saved_event_confirmation()

    st.markdown("### PDF/Excel出力")
    out1, out2 = st.columns(2)
    with out1:
        if REPORTLAB_AVAILABLE:
            try:
                pdf_bytes = make_calendar_pdf(int(year), int(month), include_detail=True)
                st.download_button(
                    "月間カレンダーPDFをダウンロード",
                    data=pdf_bytes,
                    file_name=f"hidamari_calendar_{int(year)}_{int(month):02d}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"PDFを作成できませんでした：{e}")
        else:
            st.warning("reportlab が未導入のためPDF出力できません。")
    with out2:
        month_events = []
        for day_events in monthly_events(int(year), int(month)).values():
            for ev in day_events:
                month_events.append(dict(ev))
        if month_events:
            try:
                from io import BytesIO
                bio = BytesIO()
                pd.DataFrame(month_events).to_excel(bio, index=False, sheet_name="予定")
                st.download_button(
                    "月間予定Excelをダウンロード",
                    data=bio.getvalue(),
                    file_name=f"hidamari_events_{int(year)}_{int(month):02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Excelを作成できませんでした：{e}")
        else:
            st.download_button("月間予定Excelをダウンロード", data=b"", file_name="hidamari_events_empty.xlsx", disabled=True, use_container_width=True)




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
    if storage_is_configured():
        st.caption(f"写真・添付ファイルは新規登録分からSupabase Storageへ保存されます。Storage: {get_supabase_storage_bucket()} / URL: {get_supabase_url()}")
    else:
        st.warning("Supabase Storage設定が未完了です。写真・添付ファイルを保存するには SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY / SUPABASE_STORAGE_BUCKET を設定してください。")

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
        photo_saved, photo_failed = save_uploaded_photos(event_id, uploaded_photos, photo_memo)
        file_saved, file_failed = save_uploaded_files(event_id, uploaded_files, file_memo)
        clear_event_caches()

        st.session_state["selected_calendar_event_id"] = int(event_id)
        st.session_state["last_saved_event_id"] = int(event_id)
        st.session_state["last_saved_photo_saved"] = int(photo_saved)
        st.session_state["last_saved_photo_failed"] = int(photo_failed)
        st.session_state["last_saved_file_saved"] = int(file_saved)
        st.session_state["last_saved_file_failed"] = int(file_failed)
        st.session_state["last_saved_had_photos"] = bool(uploaded_photos)
        st.session_state["last_saved_had_files"] = bool(uploaded_files)

        messages = [f"予定を保存しました（予定ID: {event_id}）。"]
        if uploaded_photos:
            messages.append(f"写真メモ：{photo_saved}件保存")
            if photo_failed:
                messages.append(f"写真メモ：{photo_failed}件失敗")
        if uploaded_files:
            messages.append(f"Excel・書類：{file_saved}件保存")
            if file_failed:
                messages.append(f"Excel・書類：{file_failed}件失敗")

        if photo_failed or file_failed:
            st.warning(" / ".join(messages))
        else:
            st.success(" / ".join(messages))


    render_saved_event_confirmation()


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

    query = f"""
        SELECT {EVENT_TODAY_COLUMNS}
        FROM events
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
            c_img, c_info = st.columns([1, 2])
            with c_img:
                render_saved_image(
                    p["file_path"],
                    caption=p["file_name"],
                    use_container_width=True,
                )
            with c_info:
                st.write(f"メモ：{p['photo_memo'] or ''}")
                st.caption(f"登録日時：{p['created_at']}")
                if st.button(f"この写真を削除 ID:{p['id']}", key=f"delete_photo_{p['id']}"):
                    delete_saved_file(p["file_path"])
                    execute("DELETE FROM event_photos WHERE id=?", (int(p["id"]),))
                    clear_event_photos_cache()
                    clear_event_caches()
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
            saved, failed = save_uploaded_photos(selected_id, add_photos, add_photo_memo)
            if saved:
                clear_event_caches()
            if failed:
                st.warning(f"写真メモ：{saved}件保存、{failed}件失敗しました。")
            else:
                st.success(f"写真メモを{saved}件追加しました。画面を再読み込みしてください。")

    st.markdown("---")
    st.write("選択中の予定のExcel・書類ファイル")

    files = get_event_files(selected_id)
    if files.empty:
        st.info("この予定に紐づくExcel・書類ファイルはありません。")
    else:
        for _, frow in files.iterrows():
            c_file, c_action = st.columns([3, 1])
            with c_file:
                st.write(f"📎 **{frow['file_name']}**")
                st.write(f"メモ：{frow['file_memo'] or ''}")
                st.caption(f"登録日時：{frow['created_at']}")
                render_saved_download_button(
                    "ダウンロード",
                    frow["file_path"],
                    frow["file_name"],
                    key=f"download_file_{frow['id']}",
                )
            with c_action:
                if st.button(f"削除 ID:{frow['id']}", key=f"delete_file_{frow['id']}"):
                    delete_saved_file(frow["file_path"])
                    execute("DELETE FROM event_files WHERE id=?", (int(frow["id"]),))
                    clear_event_files_cache()
                    clear_event_caches()
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
            saved, failed = save_uploaded_files(selected_id, add_files, add_file_memo)
            if saved:
                clear_event_caches()
            if failed:
                st.warning(f"Excel・書類ファイル：{saved}件保存、{failed}件失敗しました。")
            else:
                st.success(f"Excel・書類ファイルを{saved}件追加しました。画面を再読み込みしてください。")

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
            current_category = str(target["category"] or "")
            if current_category and current_category != "夜勤" and current_category not in category_options:
                category_options = [current_category] + category_options
            new_category = st.selectbox(
                "カテゴリ",
                category_options,
                index=category_options.index(target["category"]) if target["category"] in category_options else 0
            )
        with c2:
            new_start_time = st.text_input("開始時刻", value=format_time_for_display(target.get("start_time")))
            new_end_time = st.text_input("終了時刻", value=format_time_for_display(target.get("end_time")))
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
        clear_event_caches()
        st.success("予定を更新しました。画面を再読み込みしてください。")

    if delete_btn:
        photos = get_event_photos(selected_id)
        for _, p in photos.iterrows():
            delete_saved_file(p["file_path"])
        files = get_event_files(selected_id)
        for _, frow in files.iterrows():
            delete_saved_file(frow["file_path"])
        execute("DELETE FROM event_photos WHERE event_id=?", (int(selected_id),))
        execute("DELETE FROM event_files WHERE event_id=?", (int(selected_id),))
        execute("DELETE FROM events WHERE id=?", (int(selected_id),))
        clear_event_photos_cache()
        clear_event_files_cache()
        clear_event_caches()
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
                clear_master_caches()
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
                clear_master_caches()
                st.success("カテゴリを更新しました。画面を再読み込みしてください。")
            except Exception as e:
                st.error(f"更新できませんでした：{e}")

    if delete:
        # 既存予定で使われているカテゴリは削除ではなく非表示推奨
        used = fetch_df("SELECT COUNT(*) AS cnt FROM events WHERE category=?", (target["category_name"],))
        count = int(used.iloc[0]["cnt"]) if not used.empty else 0
        if count > 0:
            execute("UPDATE categories SET is_active=0, updated_at=? WHERE id=?", (now_text(), int(selected_id)))
            clear_master_caches()
            st.warning(f"このカテゴリは既存予定で {count} 件使われているため、削除せず非表示にしました。")
        else:
            execute("DELETE FROM categories WHERE id=?", (int(selected_id),))
            clear_master_caches()
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
                clear_master_caches()
                st.success("利用者を追加しました。")
            except DB_INTEGRITY_ERROR:
                st.error("同じ利用者名がすでに登録されています。")

    df = fetch_df(f"SELECT {USER_COLUMNS} FROM users ORDER BY is_active DESC, user_name")
    st.dataframe(df, use_container_width=True, hide_index=True)

    if not df.empty:
        selected = st.selectbox("有効／無効を変更する利用者ID", df["id"].tolist())
        target = df[df["id"] == selected].iloc[0]
        new_active = st.radio("状態", [1, 0], index=0 if target["is_active"] == 1 else 1, format_func=lambda x: "有効" if x == 1 else "無効")
        if st.button("利用者状態を更新"):
            execute("UPDATE users SET is_active=? WHERE id=?", (int(new_active), int(selected)))
            clear_master_caches()
            st.success("状態を更新しました。")


def page_master_staff():
    st.subheader("職員マスタ")
    st.caption("KING OF TIMEへシフトCSVを出力する場合は、従業員コードを登録しておくとCSVの「従業員コード」に反映されます。未入力でも保存できます。")

    with st.form("staff_add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            staff_name = st.text_input("職員名")
        with c2:
            staff_code = st.text_input("従業員コード（KING OF TIME）", placeholder="例：001（未入力可）")
        with c3:
            role = st.text_input("役割", placeholder="例：管理者、日勤、夜勤")
        add = st.form_submit_button("職員を追加")

    if add:
        clean_staff_name = normalize_staff_name(staff_name)
        clean_staff_code = str(staff_code or "").strip()
        if not clean_staff_name:
            st.error("職員名を入力してください。")
        else:
            existing_staff = fetch_df("SELECT staff_name, staff_code FROM staff")
            existing_norms = set()
            if existing_staff is not None and not existing_staff.empty:
                existing_norms = set(existing_staff["staff_name"].dropna().astype(str).apply(normalize_staff_name).tolist())
            if clean_staff_name in existing_norms:
                st.error("同じ職員名がすでに登録されています。（空白違いも同一職員として扱います）")
            elif clean_staff_code and existing_staff is not None and not existing_staff.empty and clean_staff_code in existing_staff["staff_code"].fillna("").astype(str).str.strip().tolist():
                st.error("同じKING OF TIME従業員コードがすでに登録されています。")
            else:
                try:
                    execute("""
                        INSERT INTO staff (staff_name, staff_code, role, is_active, created_at)
                        VALUES (?, ?, ?, 1, ?)
                    """, (clean_staff_name, clean_staff_code or None, role.strip(), now_text()))
                    clear_master_caches()
                    st.success("職員を追加しました。")
                except DB_INTEGRITY_ERROR:
                    st.error("同じ職員名がすでに登録されています。")

    df = fetch_df("SELECT id, staff_name, staff_code, role, is_active, created_at FROM staff ORDER BY is_active DESC, staff_name")
    st.dataframe(df, use_container_width=True, hide_index=True)

    duplicate_df = get_duplicate_staff_name_groups()
    if not duplicate_df.empty:
        st.warning("職員マスタに空白違いの重複があります。シフト表では1人として表示します。不要な重複行は無効にしてください。")
        st.dataframe(duplicate_df, use_container_width=True, hide_index=True)

    if not df.empty:
        selected = st.selectbox("編集する職員ID", df["id"].tolist())
        target = df[df["id"] == selected].iloc[0]

        with st.form("staff_edit_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                edit_staff_name = st.text_input("職員名", value=str(target.get("staff_name") or ""))
            with c2:
                edit_staff_code = st.text_input("従業員コード（KING OF TIME）", value=str(target.get("staff_code") or ""))
            with c3:
                edit_role = st.text_input("役割", value=str(target.get("role") or ""))

            new_active = st.radio(
                "状態",
                [1, 0],
                index=0 if int(target["is_active"] or 0) == 1 else 1,
                format_func=lambda x: "有効" if x == 1 else "無効",
            )
            update_staff = st.form_submit_button("職員情報を更新")

        if update_staff:
            clean_staff_name = normalize_staff_name(edit_staff_name)
            clean_staff_code = str(edit_staff_code or "").strip()
            if not clean_staff_name:
                st.error("職員名を入力してください。")
            else:
                other_df = fetch_df("SELECT id, staff_name, staff_code FROM staff WHERE id<>?", (int(selected),))
                other_norms = set()
                other_codes = set()
                if other_df is not None and not other_df.empty:
                    other_norms = set(other_df["staff_name"].dropna().astype(str).apply(normalize_staff_name).tolist())
                    other_codes = set([str(x or "").strip() for x in other_df["staff_code"].tolist() if str(x or "").strip()])
                if clean_staff_name in other_norms:
                    st.error("同じ職員名がすでに登録されています。（空白違いも同一職員として扱います）")
                elif clean_staff_code and clean_staff_code in other_codes:
                    st.error("同じKING OF TIME従業員コードがすでに登録されています。")
                else:
                    execute("""
                        UPDATE staff
                        SET staff_name=?, staff_code=?, role=?, is_active=?
                        WHERE id=?
                    """, (
                        clean_staff_name,
                        clean_staff_code or None,
                        edit_role.strip() or None,
                        int(new_active),
                        int(selected),
                    ))
                    clear_master_caches()
                    st.success("職員情報を更新しました。")



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
            e.staff_name
        FROM event_photos p
        JOIN events e ON p.event_id = e.id
        ORDER BY e.event_date DESC, p.id DESC
    """)

    if df.empty:
        st.info("写真メモはまだ登録されていません。写真アップロード時にStorage保存が失敗している場合、予定本体だけが保存され、写真メモは登録されません。")
        return

    keyword = st.text_input("写真メモ検索", placeholder="予定タイトル・利用者名・写真メモなど")
    if keyword.strip():
        kw = keyword.strip()
        mask = (
            df["title"].fillna("").str.contains(kw, case=False, na=False) |
            df["user_id"].fillna("").str.contains(kw, case=False, na=False) |
            df["user_name"].fillna("").str.contains(kw, case=False, na=False) |
            df["staff_name"].fillna("").str.contains(kw, case=False, na=False) |
            df["photo_memo"].fillna("").str.contains(kw, case=False, na=False)
        )
        df = df[mask]

    if df.empty:
        st.info("該当する写真メモはありません。")
        return

    show_df = df.drop(columns=["file_path"], errors="ignore")
    st.dataframe(show_df, use_container_width=True, hide_index=True)

    selected_photo_id = st.selectbox("表示する写真ID", df["photo_id"].tolist())
    row = df[df["photo_id"] == selected_photo_id].iloc[0]
    st.markdown("---")
    c1, c2 = st.columns([1, 2])
    with c1:
        render_saved_image(
            row["file_path"],
            use_container_width=True,
        )
    with c2:
        st.write(f"**{row['event_date']}｜{row['category']}｜{row['title']}**")
        if row["user_name"]:
            st.write(f"利用者：{row['user_name']}（ID:{row['user_id'] or ''}）")
        if row["staff_name"]:
            st.write(f"担当：{row['staff_name']}")
        st.write(f"写真メモ：{row['photo_memo'] or ''}")
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
            e.staff_name
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
            df["file_memo"].fillna("").str.contains(kw, case=False, na=False)
        )
        df = df[mask]

    if df.empty:
        st.info("該当するファイルはありません。")
        return

    show_df = df.drop(columns=["file_path"], errors="ignore")
    st.dataframe(show_df, use_container_width=True, hide_index=True)

    selected_file_id = st.selectbox("ダウンロードするファイルID", df["file_id"].tolist())
    row = df[df["file_id"] == selected_file_id].iloc[0]
    st.markdown("---")
    st.write(f"📎 **{row['file_name']}**")
    st.write(f"{row['event_date']}｜{row['category']}｜{row['title']}")
    if row["user_name"]:
        st.write(f"利用者：{row['user_name']}")
    if row["staff_name"]:
        st.write(f"担当：{row['staff_name']}")
    st.write(f"ファイルメモ：{row['file_memo'] or ''}")
    st.caption(f"登録：{row['file_created_at']}")

    render_saved_download_button(
        "ダウンロード",
        row["file_path"],
        row["file_name"],
        key=f"file_list_download_{row['file_id']}",
    )



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
    col_register = pick_import_col(df, ["登録する", "取込対象", "取込", "登録", "import", "selected", "All Day Event"])
    col_date = pick_import_col(df, ["event_date", "予定日", "日付", "開始日", "日時", "開始日時", "予定日時", "実施日", "Start Date", "start date", "DATE", "Date"])
    col_start = pick_import_col(df, ["start_time", "開始時刻", "開始時間", "時刻", "予定時刻", "時間", "Start Time", "start time", "START TIME"])
    col_end = pick_import_col(df, ["end_time", "終了時刻", "終了時間", "End Time", "end time", "END TIME"])
    col_category = pick_import_col(df, ["category", "分類", "カテゴリ", "予定分類", "種別", "キーワード", "Category", "CATEGORIES"])
    col_title = pick_import_col(df, ["title", "タイトル", "件名", "予定タイトル", "予定", "候補", "予定候補", "Subject", "subject", "SUMMARY", "Summary"])
    col_user_id = pick_import_col(df, ["user_id", "利用者ID", "利用者id", "入居者ID", "入居者id"])
    col_user_name = pick_import_col(df, ["user_name", "利用者名", "利用者", "入居者名", "入居者", "対象者"])
    col_staff = pick_import_col(df, ["staff_name", "担当", "担当者", "職員", "記入者", "作成者"])
    col_memo = pick_import_col(df, ["memo", "詳細", "メモ", "内容", "備考", "申し送り", "申し送り内容", "本文", "元の申し送り", "Description", "description", "Location", "location"])
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
          AND COALESCE(start_time, '')=?
          AND title=?
          AND COALESCE(user_id, '')=?
          AND COALESCE(user_name, '')=?
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
        clear_master_caches()
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
        clear_master_caches()
    except Exception:
        pass


def insert_import_event(row):
    """正規化済み1行をeventsへ登録。"""
    ensure_import_category(row.get("category", "その他"))
    upsert_import_user(row.get("user_id", ""), row.get("user_name", ""))
    event_id = execute("""
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
    clear_event_caches()
    return event_id


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
    # 空文字だけでなく、nan / NaT / None 文字列も登録不可として扱う
    date_ok = ~selected["event_date"].astype(str).str.strip().str.lower().isin(["", "nan", "nat", "none"])
    title_ok = ~selected["title"].astype(str).str.strip().str.lower().isin(["", "nan", "nat", "none"])
    valid = selected[date_ok & title_ok].copy()

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

# -----------------------------
# PDFカレンダー出力
# -----------------------------
def make_calendar_pdf(year, month, include_detail=True):
    return report_make_calendar_pdf(year, month, include_detail=include_detail)


# -----------------------------


def default_shift_times(shift_kind):
    """基本勤務時間。日勤2名、夜勤1名の運用を想定。"""
    if shift_kind == "日勤" or shift_kind in MANAGEMENT_SHIFT_KINDS:
        return "08:30", "17:30", 0
    if shift_kind == "夜勤":
        return "16:30", "09:30", 1
    if shift_kind == "夜勤明け":
        return "", "", 0
    return "", "", 0


def shift_short_label(shift_kind):
    return {
        "日勤": "日",
        "管": "管",
        "管理業務": "管",
        "夜勤": "夜",
        "夜勤明け": "明",
        "休み": "",
        "希望休": "希",
        "有休": "有",
        "その他": "他",
    }.get(str(shift_kind), str(shift_kind)[:1])


def shift_kind_from_editor_label(label):
    """月間シフト直接入力用の短縮ラベルを勤務区分へ変換する。"""
    value = str(label or "").strip()
    mapping = {
        "日": ["日勤"],
        "管": ["管"],
        "夜": ["夜勤"],
        "明": ["夜勤明け"],
        "休": ["休み"],
        "希": ["希望休"],
        "有": ["有休"],
        "他": ["その他"],
    }
    return mapping.get(value, [])


def shift_kind_for_selectbox(shift_kind):
    """古い保存値の管理業務も、画面上の選択肢では管として扱う。"""
    value = str(shift_kind or "").strip()
    if value in MANAGEMENT_SHIFT_KINDS:
        return "管"
    return value


@st.cache_data(ttl=60, show_spinner=False)
def get_shift_month_status(year, month):
    """指定月のシフト確定状態を取得する。"""
    try:
        df = fetch_df("""
            SELECT id, shift_year, shift_month, is_confirmed, confirmed_at, confirmed_by, created_at, updated_at
            FROM shift_month_status
            WHERE shift_year=? AND shift_month=?
            LIMIT 1
        """, (int(year), int(month)))
        if df.empty:
            return {"is_confirmed": 0, "confirmed_at": "", "confirmed_by": "", "status_error": ""}
        r = df.iloc[0]
        return {
            "is_confirmed": int(r.get("is_confirmed", 0) or 0),
            "confirmed_at": str(r.get("confirmed_at") or ""),
            "confirmed_by": str(r.get("confirmed_by") or ""),
            "status_error": "",
        }
    except Exception as e:
        return {
            "is_confirmed": 0,
            "confirmed_at": "",
            "confirmed_by": "",
            "status_error": str(e) or "確定状態を取得できませんでした。",
        }


def ensure_shift_month_editable(year, month):
    """対象月を更新できることを共通判定し、安全でない場合は拒否する。"""
    year = int(year)
    month = int(month)
    status = get_shift_month_status(year, month)
    if status.get("status_error"):
        raise ShiftUpdateBlockedError(
            f"{year}年{month}月の確定状態を確認できないため、勤務データを更新できません。"
            "時間をおいて再度お試しください。"
        )
    if status.get("is_confirmed"):
        raise ShiftUpdateBlockedError(
            f"{year}年{month}月の勤務表は確定済みのため更新できません。"
            "修正する場合は、先に確定を解除してください。"
        )
    return status


def ensure_shift_date_editable(shift_date):
    """日付から対象月を求め、共通の更新可否判定を行う。"""
    try:
        target = pd.to_datetime(str(shift_date), errors="raise").date()
    except Exception as e:
        raise ValueError(f"勤務日を読み取れません：{shift_date}") from e
    ensure_shift_month_editable(target.year, target.month)
    return target


def shift_month_is_read_only(status):
    return bool(status.get("is_confirmed") or status.get("status_error"))


def set_shift_month_status(year, month, is_confirmed, confirmed_by=""):
    """指定月のシフト確定／未確定を保存する。"""
    execute_transaction([
        ("""
            DELETE FROM shift_month_status
            WHERE shift_year=? AND shift_month=?
        """, (int(year), int(month))),
        ("""
            INSERT INTO shift_month_status
            (shift_year, shift_month, is_confirmed, confirmed_at, confirmed_by, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            int(year),
            int(month),
            1 if is_confirmed else 0,
            now_text() if is_confirmed else None,
            confirmed_by or "",
            now_text(),
            now_text(),
        )),
    ])
    clear_shift_caches()
    message = "この月のシフトを確定しました。" if is_confirmed else "この月のシフト確定を解除しました。"
    return ShiftSaveResult(SHIFT_SAVE_SAVED, count=1, message=message)

@st.cache_data(ttl=60, show_spinner=False)
def get_staff_shift_limits():
    """
    職員別の月間勤務回数上限を取得する。
    未設定職員は、日勤31回・夜勤31回・合計31回として扱う。
    """
    staff_names = get_active_staff()
    try:
        df = fetch_df("""
            SELECT staff_name, max_day_shifts, max_night_shifts, max_total_shifts, note
            FROM staff_shift_limits
            ORDER BY staff_name
        """)
    except Exception:
        df = pd.DataFrame(columns=["staff_name", "max_day_shifts", "max_night_shifts", "max_total_shifts", "note"])

    rows = []
    for staff_name in staff_names:
        hit = df[df["staff_name"].astype(str) == str(staff_name)] if not df.empty else pd.DataFrame()
        if hit.empty:
            rows.append({
                "職員名": staff_name,
                "日勤上限": 31,
                "夜勤上限": 31,
                "合計上限": 31,
                "メモ": "",
            })
        else:
            r = hit.iloc[0]
            rows.append({
                "職員名": staff_name,
                "日勤上限": safe_shift_limit_value(r.get("max_day_shifts", 31), 31),
                "夜勤上限": safe_shift_limit_value(r.get("max_night_shifts", 31), 31),
                "合計上限": safe_shift_limit_value(r.get("max_total_shifts", 31), 31),
                "メモ": str(r.get("note") or ""),
            })
    return pd.DataFrame(rows)


def save_staff_shift_limits_from_editor(limits_df):
    """職員別勤務上限を保存する。"""
    if limits_df is None or limits_df.empty:
        return 0
    execute("DELETE FROM staff_shift_limits")
    params = []
    seen_staff_names = set()
    for _, r in limits_df.iterrows():
        staff_name = normalize_staff_name(r.get("職員名", ""))
        if not staff_name or staff_name in seen_staff_names:
            continue
        seen_staff_names.add(staff_name)
        params.append((
            staff_name,
            safe_shift_limit_value(r.get("日勤上限", 31), 31),
            safe_shift_limit_value(r.get("夜勤上限", 31), 31),
            safe_shift_limit_value(r.get("合計上限", 31), 31),
            str(r.get("メモ", "") or "").strip() or None,
            now_text(),
            now_text(),
        ))
    if not params:
        clear_shift_caches()
        return 0
    saved = execute_many("""
        INSERT INTO staff_shift_limits
        (staff_name, max_day_shifts, max_night_shifts, max_total_shifts, note, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, params)
    clear_shift_caches()
    return saved




def normalize_staff_name(value):
    """
    職員名を比較・保存用に正規化する。
    「職員 A」「職員　A」「職員A」のような空白違いを同一職員として扱う。
    シフト表・上限表・AI案ではこの正規化名を使い、重複行を防ぐ。
    """
    if value is None:
        return ""
    v = str(value).replace("\u3000", " ").strip()
    v = re.sub(r"\s+", "", v)
    return v


def shift_duplicate_key(shift_date, shift_kind, staff_name, staff_id=None):
    staff_key = str(staff_id or "").strip()
    if staff_key:
        staff_part = ("id", staff_key)
    else:
        staff_part = ("name", normalize_staff_name(staff_name))
    return (str(shift_date or "").strip(), str(shift_kind or "").strip(), staff_part)


def dedupe_shift_insert_params(params):
    unique = []
    seen = set()
    for p in params or []:
        key = shift_duplicate_key(p[0], p[2], p[1])
        if not key[0] or not key[1] or not key[2][1] or key in seen:
            continue
        seen.add(key)
        unique.append(p)
    return unique


def dedupe_shift_dataframe(df):
    if df is None or df.empty:
        return df
    rows = []
    seen = set()
    for _, row in df.iterrows():
        staff_id = row.get("staff_id", None) if "staff_id" in df.columns else None
        key = shift_duplicate_key(row.get("shift_date"), row.get("shift_kind"), row.get("staff_name"), staff_id)
        if not key[0] or not key[1] or not key[2][1] or key in seen:
            continue
        seen.add(key)
        rows.append(row)
    if not rows:
        return df.iloc[0:0].copy()
    return pd.DataFrame(rows).reset_index(drop=True)


def unique_shift_staff_names(df):
    if df is None or df.empty:
        return []
    names = []
    seen = set()
    for _, row in df.iterrows():
        raw_name = str(row.get("staff_name") or "").strip()
        display_name = normalize_staff_name(raw_name)
        staff_id = row.get("staff_id", None) if "staff_id" in df.columns else None
        staff_key = str(staff_id or "").strip() or display_name
        if not display_name or staff_key in seen:
            continue
        seen.add(staff_key)
        names.append(display_name)
    return names


def get_duplicate_staff_name_groups():
    """職員マスタ内の空白違い重複を確認するための一覧を作る。"""
    try:
        df = fetch_df("SELECT id, staff_name, role, is_active, created_at FROM staff ORDER BY is_active DESC, staff_name")
    except Exception:
        return pd.DataFrame(columns=["正規化名", "登録ID", "登録名", "状態"])
    if df is None or df.empty:
        return pd.DataFrame(columns=["正規化名", "登録ID", "登録名", "状態"])

    tmp = df.copy()
    tmp["正規化名"] = tmp["staff_name"].apply(normalize_staff_name)
    dup_names = tmp[tmp["正規化名"].duplicated(keep=False)]["正規化名"].unique().tolist()
    if not dup_names:
        return pd.DataFrame(columns=["正規化名", "登録ID", "登録名", "状態"])

    rows = []
    for norm in dup_names:
        part = tmp[tmp["正規化名"] == norm]
        for _, r in part.iterrows():
            rows.append({
                "正規化名": norm,
                "登録ID": int(r.get("id")),
                "登録名": str(r.get("staff_name") or ""),
                "状態": "有効" if int(r.get("is_active") or 0) == 1 else "無効",
            })
    return pd.DataFrame(rows, columns=["正規化名", "登録ID", "登録名", "状態"])


def safe_shift_limit_value(value, default=31):
    """
    職員別勤務回数上限を安全に数値化する。
    重要：0回指定は「勤務不可」として有効な設定なので、0を31へ戻さない。
    空欄・NaN・不正値だけを初期値31として扱う。
    """
    try:
        if value is None:
            return int(default)
        if isinstance(value, str) and value.strip() == "":
            return int(default)
        if pd.isna(value):
            return int(default)
        v = int(float(value))
        return max(0, min(31, v))
    except Exception:
        return int(default)


def normalize_shift_limits_for_staff(limit_map, staff_name):
    """limit_mapから職員の上限を取り出し、day/night/totalを必ず安全なintにそろえる。"""
    limit_map = limit_map or {}
    raw = limit_map.get(normalize_staff_name(staff_name), {"day": 31, "night": 31, "total": 31})
    day_limit = safe_shift_limit_value(raw.get("day", 31), 31)
    night_limit = safe_shift_limit_value(raw.get("night", 31), 31)
    total_limit = safe_shift_limit_value(raw.get("total", 31), 31)
    return {"day": day_limit, "night": night_limit, "total": total_limit}


def staff_month_work_counts(df, staff_name, year, month):
    """
    指定職員の指定月内の日勤・夜勤・合計勤務数を数える。
    職員名は正規化して比較するため、空白違いでも上限判定から漏れない。
    """
    if df is None or df.empty:
        return 0, 0, 0
    target_staff = normalize_staff_name(staff_name)
    work = df.copy()
    work["_staff_norm"] = work["staff_name"].apply(normalize_staff_name)
    work = work[work["_staff_norm"] == target_staff]
    if work.empty:
        return 0, 0, 0
    ym = f"{int(year)}-{int(month):02d}-"
    work = work[work["shift_date"].astype(str).str.startswith(ym)]
    if work.empty:
        return 0, 0, 0
    day_count = int(len(work[work["shift_kind"].astype(str).isin(DAY_LIMIT_SHIFT_KINDS)]))
    night_count = int(len(work[work["shift_kind"].astype(str).isin(NIGHT_LIMIT_SHIFT_KINDS)]))
    total_count = int(len(work[work["shift_kind"].astype(str).isin(WORKDAY_SHIFT_KINDS)]))
    return day_count, night_count, total_count

def would_exceed_staff_shift_limit(df, staff_name, target_date, shift_kind, limit_map=None):
    """
    その勤務を1件追加した場合に、職員別の日勤・夜勤・合計上限を超えるか判定する。
    AIシフト案作成時の最終ガードとして使う。
    """
    if shift_kind not in WORKDAY_SHIFT_KINDS:
        return False, ""
    target = pd.to_datetime(str(target_date)).date()
    limits = normalize_shift_limits_for_staff(limit_map or get_staff_shift_limit_map(), staff_name)
    day_count, night_count, total_count = staff_month_work_counts(df, staff_name, target.year, target.month)

    add_day = 1 if shift_kind in DAY_LIMIT_SHIFT_KINDS else 0
    add_night = 1 if shift_kind in NIGHT_LIMIT_SHIFT_KINDS else 0

    if day_count + add_day > limits["day"]:
        return True, f"日勤上限{limits['day']}回を超えるため候補外"
    if night_count + add_night > limits["night"]:
        return True, f"夜勤上限{limits['night']}回を超えるため候補外"
    if total_count + 1 > limits["total"]:
        return True, f"合計上限{limits['total']}回を超えるため候補外"
    return False, ""


def get_staff_shift_limit_map():
    """
    職員別勤務回数上限をDBから直接取得し、職員名を正規化して辞書化する。
    画面表示用の結合結果ではなく保存値そのものを見ることで、
    日勤上限0・夜勤上限0がAI判定で31扱いになる事故を防ぐ。
    """
    mapping = {}
    try:
        df = fetch_df("""
            SELECT staff_name, max_day_shifts, max_night_shifts, max_total_shifts
            FROM staff_shift_limits
        """)
    except Exception:
        df = pd.DataFrame(columns=["staff_name", "max_day_shifts", "max_night_shifts", "max_total_shifts"])

    if df is not None and not df.empty:
        for _, r in df.iterrows():
            staff_name = normalize_staff_name(r.get("staff_name", ""))
            if not staff_name:
                continue
            mapping[staff_name] = {
                "day": safe_shift_limit_value(r.get("max_day_shifts", 31), 31),
                "night": safe_shift_limit_value(r.get("max_night_shifts", 31), 31),
                "total": safe_shift_limit_value(r.get("max_total_shifts", 31), 31),
            }

    # 未設定職員だけ31回にする。設定済みの0回は絶対に上書きしない。
    for staff_name in get_active_staff():
        s = normalize_staff_name(staff_name)
        if s and s not in mapping:
            mapping[s] = {"day": 31, "night": 31, "total": 31}

    return mapping


def debug_shift_limit_summary_for_ai(limit_map=None):
    """AIが実際に参照している上限表を確認するための表示用データ。"""
    limit_map = limit_map or get_staff_shift_limit_map()
    rows = []
    for staff_name in sorted(limit_map.keys()):
        limits = normalize_shift_limits_for_staff(limit_map, staff_name)
        rows.append({
            "職員名": staff_name,
            "AI参照_日勤上限": limits["day"],
            "AI参照_夜勤上限": limits["night"],
            "AI参照_合計上限": limits["total"],
        })
    return pd.DataFrame(rows)

def create_shift_limit_check_table(matrix, limit_map=None):
    """月間勤務回数が職員別上限を超えていないか確認する。"""
    if matrix is None or matrix.empty:
        return pd.DataFrame(columns=["重要度", "種類", "職員名", "内容"])
    limit_map = limit_map or get_staff_shift_limit_map()
    rows = []
    for _, r in matrix.iterrows():
        staff_name = normalize_staff_name(r.get("職員名", ""))
        limits = normalize_shift_limits_for_staff(limit_map, staff_name)
        day_count = int(r.get("日勤", 0) or 0)
        night_count = int(r.get("夜勤", 0) or 0)
        total_count = int(r.get("合計", 0) or 0)
        if day_count > limits["day"]:
            rows.append({"重要度": "高", "種類": "日勤上限超過", "職員名": staff_name, "内容": f"日勤 {day_count}回 / 上限 {limits['day']}回"})
        if night_count > limits["night"]:
            rows.append({"重要度": "高", "種類": "夜勤上限超過", "職員名": staff_name, "内容": f"夜勤 {night_count}回 / 上限 {limits['night']}回"})
        if total_count > limits["total"]:
            rows.append({"重要度": "中", "種類": "合計上限超過", "職員名": staff_name, "内容": f"合計 {total_count}回 / 上限 {limits['total']}回"})
    return pd.DataFrame(rows, columns=["重要度", "種類", "職員名", "内容"])




def clear_month_staff_shifts(year, month):
    """
    指定月のシフトを全削除する。
    職員別勤務回数上限は残し、月間シフト表だけを空にして再入力できるようにする。
    """
    year = int(year)
    month = int(month)
    ensure_shift_month_editable(year, month)
    last_day = calendar.monthrange(year, month)[1]
    start = f"{year}-{month:02d}-01"
    end = f"{year}-{month:02d}-{last_day:02d}"
    execute("DELETE FROM staff_shifts WHERE shift_date BETWEEN ? AND ?", (start, end))
    clear_shift_caches()
    try:
        st.session_state.pop("ai_shift_draft", None)
        st.session_state["shift_editor_reset_counter"] = int(st.session_state.get("shift_editor_reset_counter", 0) or 0) + 1
    except Exception:
        pass
    return ShiftSaveResult(SHIFT_SAVE_SAVED, message="この月のシフトを全クリアしました。")


def shift_month_is_confirmed(year, month):
    return bool(get_shift_month_status(year, month).get("is_confirmed"))


def fetch_staff_shifts_raw(start_date, end_date, keyword=""):
    query = f"""
        SELECT {SHIFT_COLUMNS}
        FROM staff_shifts
        WHERE shift_date BETWEEN ? AND ?
    """
    params = [str(start_date), str(end_date)]
    if keyword.strip():
        query += " AND (staff_name LIKE ? OR shift_kind LIKE ? OR memo LIKE ?)"
        kw = f"%{keyword.strip()}%"
        params.extend([kw, kw, kw])
    query += " ORDER BY shift_date, staff_name, shift_kind, id"
    return fetch_df(query, params)


def get_staff_shifts(start_date, end_date, keyword=""):
    return dedupe_shift_dataframe(fetch_staff_shifts_raw(start_date, end_date, keyword))


@st.cache_data(ttl=60, show_spinner=False)
def get_staff_shifts_month(year, month, include_prev_day=False):
    last_day = calendar.monthrange(int(year), int(month))[1]
    start_date = date(int(year), int(month), 1)
    if include_prev_day:
        start_date = start_date - pd.Timedelta(days=1)
        start = start_date.strftime("%Y-%m-%d")
    else:
        start = f"{int(year)}-{int(month):02d}-01"
    end = f"{int(year)}-{int(month):02d}-{last_day:02d}"
    return get_staff_shifts(start, end)


def save_single_shift(shift_date, staff_name, shift_kind, start_time=None, end_time=None, next_day=0, memo=""):
    """1件のシフトを追加保存する。"""
    ensure_shift_date_editable(shift_date)
    staff_name = normalize_staff_name(staff_name)
    if not staff_name or not shift_kind:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="保存する勤務情報がありません。")
    existing_df = fetch_staff_shifts_raw(shift_date, shift_date)
    if existing_df is not None and not existing_df.empty:
        new_key = shift_duplicate_key(shift_date, shift_kind, staff_name)
        for _, row in existing_df.iterrows():
            if shift_duplicate_key(row.get("shift_date"), row.get("shift_kind"), row.get("staff_name")) == new_key:
                return ShiftSaveResult(
                    SHIFT_SAVE_DUPLICATE,
                    message="同じ職員・日付・勤務区分が登録済みのため、重複保存しませんでした。",
                )
    if start_time is None or end_time is None:
        start_time, end_time, default_next = default_shift_times(shift_kind)
        next_day = default_next
    shift_id = execute("""
        INSERT INTO staff_shifts
        (shift_date, staff_name, shift_kind, start_time, end_time, next_day, memo, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        shift_date,
        staff_name,
        shift_kind,
        start_time or None,
        end_time or None,
        int(next_day or 0),
        memo.strip() or None,
        now_text(),
        now_text(),
    ))
    clear_shift_caches()
    return ShiftSaveResult(SHIFT_SAVE_SAVED, count=1, message="勤務を保存しました。")


def save_day_shift_assignments(shift_date, assignments, memo="日付カードから編集"):
    """
    選択日の職員別シフトをまとめて保存する。
    空欄は削除扱いにし、既存の管理業務表記も同じ職員・同じ日付として置き換える。
    """
    ensure_shift_date_editable(shift_date)
    shift_date = str(shift_date)
    assignments = assignments or {}
    staff_names = []
    seen_staff_names = set()
    for staff_name in assignments.keys():
        normalized = normalize_staff_name(staff_name)
        if normalized and normalized not in seen_staff_names:
            seen_staff_names.add(normalized)
            staff_names.append(normalized)
    if not staff_names:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="保存対象の職員がいません。")

    existing_day_df = fetch_staff_shifts_raw(shift_date, shift_date)
    operations = []
    for staff_name in staff_names:
        aliases = {staff_name}
        if existing_day_df is not None and not existing_day_df.empty:
            for old_name in existing_day_df["staff_name"].dropna().astype(str).unique().tolist():
                if normalize_staff_name(old_name) == staff_name:
                    aliases.add(old_name)
        for alias in aliases:
            operations.append((
                "DELETE FROM staff_shifts WHERE shift_date=? AND staff_name=?",
                (shift_date, alias),
            ))

    params = []
    for staff_name in staff_names:
        label = str(assignments.get(staff_name, "") or "").strip()
        if not label:
            continue
        kinds = []
        labels = [label] if "/" not in label else [x.strip() for x in label.split("/") if x.strip()]
        for lab in labels:
            kinds.extend(shift_kind_from_editor_label(lab))
        for kind in kinds:
            stime, etime, next_day = default_shift_times(kind)
            params.append((
                shift_date,
                staff_name,
                kind,
                stime or None,
                etime or None,
                int(next_day or 0),
                memo.strip() or None,
                now_text(),
                now_text(),
            ))

    params = dedupe_shift_insert_params(params)
    if params:
        operations.append(("""
            INSERT INTO staff_shifts
            (shift_date, staff_name, shift_kind, start_time, end_time, next_day, memo, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, params, True))
    execute_transaction(operations)
    clear_shift_caches()
    return ShiftSaveResult(
        SHIFT_SAVE_SAVED,
        count=len(params),
        changed_cells=len(staff_names),
        message=f"この日のシフトを保存しました。登録件数：{len(params)}件",
    )


def save_basic_day_shift(shift_date, day_staff_1, day_staff_2, night_staff, memo=""):
    """
    1日分の基本シフトを保存する。
    同日の既存日勤・夜勤は一度削除し、日勤2名・夜勤1名として登録する。
    """
    ensure_shift_date_editable(shift_date)
    params = []
    for staff_name in [day_staff_1, day_staff_2]:
        staff_name = normalize_staff_name(staff_name)
        if staff_name:
            stime, etime, nd = default_shift_times("日勤")
            params.append((shift_date, staff_name, "日勤", stime, etime, nd, memo.strip() or None, now_text(), now_text()))
    night_staff = normalize_staff_name(night_staff)
    if night_staff:
        stime, etime, nd = default_shift_times("夜勤")
        params.append((shift_date, night_staff, "夜勤", stime, etime, nd, memo.strip() or None, now_text(), now_text()))
    params = dedupe_shift_insert_params(params)
    if not params:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="保存する基本シフトがありません。")
    execute_transaction([
        (
            "DELETE FROM staff_shifts WHERE shift_date=? AND shift_kind IN ('日勤', '夜勤')",
            (shift_date,),
        ),
        ("""
            INSERT INTO staff_shifts
            (shift_date, staff_name, shift_kind, start_time, end_time, next_day, memo, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, params, True),
    ])
    clear_shift_caches()
    return ShiftSaveResult(
        SHIFT_SAVE_SAVED,
        count=len(params),
        message=f"基本シフトを保存しました。登録件数：{len(params)}件",
    )


def update_staff_shift(
    shift_id, shift_date, staff_name, shift_kind, start_time=None, end_time=None,
    next_day=0, memo="",
):
    """過去シフト更新を確定状態の共通ガード経由で実行する。"""
    current_df = fetch_df(f"SELECT {SHIFT_COLUMNS} FROM staff_shifts WHERE id=? LIMIT 1", (int(shift_id),))
    if current_df is None or current_df.empty:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="更新対象のシフトが見つかりません。")
    current = current_df.iloc[0]
    current_date = ensure_shift_date_editable(current.get("shift_date"))
    new_date = ensure_shift_date_editable(shift_date)
    # 月をまたいで移動する場合も、移動元・移動先の両方を検査済みとする。
    _ = current_date, new_date

    staff_name = normalize_staff_name(staff_name)
    if not staff_name or not shift_kind:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="職員または勤務区分が未入力です。")

    duplicate_df = fetch_staff_shifts_raw(str(shift_date), str(shift_date))
    if duplicate_df is not None and not duplicate_df.empty:
        new_key = shift_duplicate_key(shift_date, shift_kind, staff_name)
        for _, row in duplicate_df.iterrows():
            if int(row.get("id") or 0) == int(shift_id):
                continue
            if shift_duplicate_key(row.get("shift_date"), row.get("shift_kind"), row.get("staff_name")) == new_key:
                return ShiftSaveResult(
                    SHIFT_SAVE_DUPLICATE,
                    message="同じ職員・日付・勤務区分が登録済みのため、更新しませんでした。",
                )

    execute("""
        UPDATE staff_shifts
        SET shift_date=?, staff_name=?, shift_kind=?, start_time=?, end_time=?, next_day=?, memo=?, updated_at=?
        WHERE id=?
    """, (
        str(shift_date), staff_name, shift_kind, start_time or None, end_time or None,
        int(next_day or 0), str(memo or "").strip() or None, now_text(), int(shift_id),
    ))
    clear_shift_caches()
    return ShiftSaveResult(SHIFT_SAVE_SAVED, count=1, message="シフトを更新しました。")


def delete_staff_shift(shift_id):
    """過去シフト削除を確定状態の共通ガード経由で実行する。"""
    current_df = fetch_df(f"SELECT {SHIFT_COLUMNS} FROM staff_shifts WHERE id=? LIMIT 1", (int(shift_id),))
    if current_df is None or current_df.empty:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="削除対象のシフトが見つかりません。")
    ensure_shift_date_editable(current_df.iloc[0].get("shift_date"))
    execute("DELETE FROM staff_shifts WHERE id=?", (int(shift_id),))
    clear_shift_caches()
    return ShiftSaveResult(SHIFT_SAVE_SAVED, count=1, message="シフトを削除しました。")


MANAGEMENT_INDEX_KIND = "__management__"


def build_shift_read_index(df, report_year=None, report_month=None):
    """月間シフトを1回走査し、画面表示と帳票署名で共有する読み取り索引を作る。"""
    kinds_by_staff_date = {}
    raw_kinds_by_staff_date = {}
    entries_by_date_kind = {}
    kind_counts_by_date = {}
    normalized_by_raw_staff = {}
    normalized_staff_names = set()
    raw_staff_names = set()
    report_month_rows = []
    report_kot_rows = []
    report_previous_rows = []
    month_prefix = None
    if report_year is not None and report_month is not None:
        month_prefix = f"{int(report_year)}-{int(report_month):02d}-"

    if df is not None and not df.empty:
        required = ["shift_date", "staff_name", "shift_kind"]
        for column in required:
            if column not in df.columns:
                raise KeyError(column)
        report_columns = [column for column in SHIFT_REPORT_SNAPSHOT_COLUMNS if column in df.columns]
        read_columns = list(dict.fromkeys(required + report_columns))
        positions = {column: index for index, column in enumerate(read_columns)}
        report_positions = {column: SHIFT_REPORT_SNAPSHOT_COLUMNS.index(column) for column in report_columns}

        for values in df[read_columns].itertuples(index=False, name=None):
            shift_date_value = values[positions["shift_date"]]
            staff_name_value = values[positions["staff_name"]]
            shift_kind_value = values[positions["shift_kind"]]
            date_text = str(shift_date_value)
            kind_text = str(shift_kind_value)
            staff_name_for_normalization = None if pd.isna(staff_name_value) else staff_name_value
            raw_staff_name = "" if staff_name_for_normalization is None else str(staff_name_value)
            normalized_staff_name = normalize_staff_name(staff_name_for_normalization)

            counts = kind_counts_by_date.setdefault(date_text, {})
            counts[kind_text] = counts.get(kind_text, 0) + 1

            if raw_staff_name:
                raw_staff_names.add(raw_staff_name)
                normalized_by_raw_staff[raw_staff_name] = normalized_staff_name
                raw_dates = raw_kinds_by_staff_date.setdefault(raw_staff_name, {})
                raw_dates.setdefault(date_text, []).append(kind_text)

            if normalized_staff_name:
                normalized_staff_names.add(normalized_staff_name)
                staff_dates = kinds_by_staff_date.setdefault(normalized_staff_name, {})
                staff_dates.setdefault(date_text, []).append(kind_text)
                date_entries = entries_by_date_kind.setdefault(date_text, {})
                entry = (raw_staff_name, normalized_staff_name)
                date_entries.setdefault(kind_text, []).append(entry)
                if kind_text in MANAGEMENT_SHIFT_KINDS:
                    date_entries.setdefault(MANAGEMENT_INDEX_KIND, []).append(entry)

            if month_prefix is not None:
                report_row = [""] * len(SHIFT_REPORT_SNAPSHOT_COLUMNS)
                for column, report_position in report_positions.items():
                    report_row[report_position] = _normalize_shift_snapshot_value(column, values[positions[column]])
                report_row = tuple(report_row)
                core_row = report_row[1:4]
                if report_row[1].startswith(month_prefix):
                    report_month_rows.append(core_row)
                    report_kot_rows.append(report_row)
                else:
                    report_previous_rows.append(core_row)

    def freeze_nested(source):
        return {
            key: {nested_key: tuple(sorted(values)) for nested_key, values in nested.items()}
            for key, nested in source.items()
        }

    report_snapshot = None
    if month_prefix is not None:
        report_snapshot = _make_shift_report_snapshot(
            report_month_rows, report_kot_rows, report_previous_rows
        )
    return {
        "kinds_by_staff_date": freeze_nested(kinds_by_staff_date),
        "raw_kinds_by_staff_date": freeze_nested(raw_kinds_by_staff_date),
        "entries_by_date_kind": freeze_nested(entries_by_date_kind),
        "kind_counts_by_date": {key: dict(value) for key, value in kind_counts_by_date.items()},
        "normalized_by_raw_staff": dict(normalized_by_raw_staff),
        "normalized_staff_names": tuple(sorted(normalized_staff_names)),
        "raw_staff_names": tuple(sorted(raw_staff_names)),
        "report_snapshot": report_snapshot,
    }


def _unique_index_staff_names(entries, staff_filter=None):
    names = []
    seen = set()
    for raw_name, normalized_name in entries or ():
        if staff_filter and staff_filter != "全職員" and raw_name != staff_filter:
            continue
        if normalized_name and normalized_name not in seen:
            seen.add(normalized_name)
            names.append(normalized_name)
    return names


def shift_day_labels_for_staff(df, staff_name, target_date, read_index=None):
    """
    指定職員・指定日のラベル一覧。
    前日夜勤があれば明を補完するが、当日に希望休・有休・休み・その他がある場合は明を表示しない。
    """
    labels = []
    read_index = read_index or build_shift_read_index(df)
    target_staff = normalize_staff_name(staff_name)
    staff_dates = read_index["kinds_by_staff_date"].get(target_staff, {})
    day_shift_kinds = list(staff_dates.get(str(target_date), ()))
    for kind in day_shift_kinds:
        label = shift_short_label(kind)
        if label and label not in labels:
            labels.append(label)

    # 当日に休み系がある場合、明け表示は重ねない。
    has_off_label = any(x in labels for x in ["希", "有", "他"]) or any(k in OFF_OR_BLOCKING_SHIFT_KINDS for k in day_shift_kinds)

    # 前日夜勤の翌日は「明」を自動表示する。ただし希望休・有休・休み・その他がある日は表示しない。
    try:
        prev_date = (pd.to_datetime(str(target_date)).date() - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        if "夜勤" in staff_dates.get(prev_date, ()) and "明" not in labels and not has_off_label:
            labels.insert(0, "明")
    except Exception:
        pass

    # 既存データで明けと希望休等が既に重なっている場合も、表示上は休み系を優先して明を消す。
    if has_off_label and "明" in labels:
        labels = [x for x in labels if x != "明"]
    return labels


def shift_day_actual_label_for_staff(df, staff_name, target_date, read_index=None):
    """指定日のDB登録済みシフトだけを、日付編集フォーム用の短縮ラベルにする。"""
    labels = []
    read_index = read_index or build_shift_read_index(df)
    target_staff = normalize_staff_name(staff_name)
    kinds = read_index["kinds_by_staff_date"].get(target_staff, {}).get(str(target_date), ())
    for kind in kinds:
        label = "休" if str(kind or "") == "休み" else shift_short_label(str(kind or ""))
        if label and label not in labels:
            labels.append(label)
    return labels_to_cell_value(labels)


def shift_day_editor_cell_value(df, staff_name, target_date, read_index=None):
    """直接入力用。帳票では非表示の通常休も「休」として区別する。"""
    read_index = read_index or build_shift_read_index(df)
    labels = shift_day_labels_for_staff(df, staff_name, target_date, read_index)
    target_staff = normalize_staff_name(staff_name)
    kinds = read_index["kinds_by_staff_date"].get(target_staff, {}).get(str(target_date), ())
    if "休み" in kinds and "休" not in labels:
        labels.append("休")
    return labels_to_cell_value(labels)


def labels_to_cell_value(labels):
    labels = [x for x in labels if x]
    if not labels:
        return ""

    # 希望休・有休・その他がある日は、明けや勤務より休み系表示を優先する。
    for off_label in ["希", "有", "他"]:
        if off_label in labels:
            return off_label

    # 勤務系の同日入力は禁止。既存データで重なっている場合は日、管、夜の順に優先表示し、チェックで警告する。
    priority = ["日", "管", "夜", "明", "休", "希", "有", "他"]
    if sum(1 for x in ["日", "管", "夜"] if x in labels) >= 2:
        return next(x for x in priority if x in labels)

    ordered = [x for x in priority if x in labels]
    if len(ordered) == 1:
        return ordered[0]
    return "/".join(ordered[:2])

def create_shift_matrix(df, year, month, read_index=None):
    """写真の勤務表イメージに近い、職員×日付の横長表を作る。"""
    last_day = calendar.monthrange(int(year), int(month))[1]
    if read_index is None:
        read_index = build_shift_read_index(df)
    staff_names = list(read_index["normalized_staff_names"])
    columns = ["職員名"] + [str(d) for d in range(1, last_day + 1)] + ["日勤", "管", "夜勤", "明", "休み", "希望休", "有休", "他", "合計", "最大連勤"]
    rows = []

    for staff_name in staff_names:
        row = {"職員名": staff_name}
        day_count = management_count = night_count = ake_count = rest_count = hope_count = paid_count = other_count = total_count = 0
        work_flags = []
        for d in range(1, last_day + 1):
            target_date = f"{int(year)}-{int(month):02d}-{d:02d}"
            labels = shift_day_labels_for_staff(df, staff_name, target_date, read_index)
            row[str(d)] = labels_to_cell_value(labels)
            if "日" in labels:
                day_count += 1
                total_count += 1
            if "管" in labels:
                management_count += 1
                total_count += 1
            if "夜" in labels:
                night_count += 1
                total_count += 1
            if "明" in labels:
                ake_count += 1
                total_count += 1
            if "休" in labels:
                rest_count += 1
            if "希" in labels:
                hope_count += 1
            if "有" in labels:
                paid_count += 1
            if "他" in labels:
                other_count += 1
            work_flags.append(1 if any(label in labels for label in ["日", "管", "夜", "明"]) else 0)

        row["日勤"] = day_count
        row["管"] = management_count
        row["夜勤"] = night_count
        row["明"] = ake_count
        row["休み"] = rest_count
        row["希望休"] = hope_count
        row["有休"] = paid_count
        row["他"] = other_count
        row["合計"] = total_count
        row["最大連勤"] = max_consecutive_ones(work_flags)
        rows.append(row)

    return pd.DataFrame(rows, columns=columns)


def max_consecutive_ones(values):
    best = cur = 0
    for v in values:
        if int(v or 0) == 1:
            cur += 1
            best = max(best, cur)
        else:
            cur = 0
    return best


def create_editable_shift_matrix(staff_names, df, year, month, read_index=None):
    """
    st.data_editorで直接入力しやすい月間シフト表を作る。
    各セルは「」「日」「管」「夜」「明」「休」「希」「有」「他」からプルダウン入力する。
    """
    last_day = calendar.monthrange(int(year), int(month))[1]
    # 職員名は正規化し、空白違いの重複行を1行にまとめる。
    unique_staff_names = []
    seen_staff_names = set()
    for s in staff_names:
        ns = normalize_staff_name(s)
        if ns and ns not in seen_staff_names:
            seen_staff_names.add(ns)
            unique_staff_names.append(ns)
    staff_names = unique_staff_names

    if read_index is None:
        read_index = build_shift_read_index(df)
    for s in read_index["raw_staff_names"]:
        ns = read_index["normalized_by_raw_staff"].get(s, "")
        if ns and ns not in seen_staff_names:
            seen_staff_names.add(ns)
            staff_names.append(ns)

    columns = ["職員名"] + [str(d) for d in range(1, last_day + 1)]
    rows = []

    for staff_name in staff_names:
        row = {"職員名": staff_name}
        for d in range(1, last_day + 1):
            target_date = f"{int(year)}-{int(month):02d}-{d:02d}"
            row[str(d)] = shift_day_editor_cell_value(
                df, staff_name, target_date, read_index
            )
        rows.append(row)

    return pd.DataFrame(rows, columns=columns)


def save_shift_matrix_from_editor(year, month, edited_df):
    """
    月間シフト表の直接編集内容を差分保存する。
    表示値が変わった職員・日付だけを同一トランザクションで置き換える。
    """
    year = int(year)
    month = int(month)
    ensure_shift_month_editable(year, month)
    if edited_df is None or edited_df.empty:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="保存する変更がありません。")

    last_day = calendar.monthrange(year, month)[1]
    staff_names = []
    seen_staff_names = set()
    for v in edited_df["職員名"].tolist():
        ns = normalize_staff_name(v)
        if ns and ns not in seen_staff_names:
            seen_staff_names.add(ns)
            staff_names.append(ns)
    if not staff_names:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="保存対象の職員がいません。")

    first_date = date(year, month, 1)
    start = first_date.strftime("%Y-%m-%d")
    comparison_start = (first_date - timedelta(days=1)).strftime("%Y-%m-%d")
    end = f"{year}-{month:02d}-{last_day:02d}"
    existing_month_df = fetch_staff_shifts_raw(start, end)
    comparison_df = fetch_staff_shifts_raw(comparison_start, end)
    comparison_index = build_shift_read_index(comparison_df)

    aliases_by_staff = {}
    for staff_name in staff_names:
        aliases = {staff_name}
        if existing_month_df is not None and not existing_month_df.empty:
            for old_name in existing_month_df["staff_name"].dropna().astype(str).unique().tolist():
                if normalize_staff_name(old_name) == staff_name:
                    aliases.add(old_name)
        aliases_by_staff[staff_name] = aliases

    operations = []
    insert_params = []
    changed_cells = 0
    for _, row in edited_df.iterrows():
        staff_name = normalize_staff_name(row.get("職員名", ""))
        if not staff_name:
            continue
        for d in range(1, last_day + 1):
            target_date = f"{year}-{month:02d}-{d:02d}"
            raw_label = row.get(str(d), "")
            label = "" if pd.isna(raw_label) else str(raw_label or "").strip()
            original_label = shift_day_editor_cell_value(
                comparison_df,
                staff_name,
                target_date,
                comparison_index,
            )
            if label == original_label:
                continue

            changed_cells += 1
            for alias in aliases_by_staff.get(staff_name, {staff_name}):
                operations.append((
                    "DELETE FROM staff_shifts WHERE shift_date=? AND staff_name=?",
                    (target_date, alias),
                ))

            labels = [label] if "/" not in label else [x.strip() for x in label.split("/") if x.strip()]
            kinds = []
            for lab in labels:
                kinds.extend(shift_kind_from_editor_label(lab))
            for kind in kinds:
                stime, etime, next_day = default_shift_times(kind)
                insert_params.append((
                    target_date,
                    staff_name,
                    kind,
                    stime or None,
                    etime or None,
                    int(next_day or 0),
                    "月間表から直接入力",
                    now_text(),
                    now_text(),
                ))

    if not changed_cells:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="変更されたセルはありません。")

    insert_params = dedupe_shift_insert_params(insert_params)
    if insert_params:
        operations.append(("""
            INSERT INTO staff_shifts
            (shift_date, staff_name, shift_kind, start_time, end_time, next_day, memo, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, insert_params, True))
    execute_transaction(operations)
    clear_shift_caches()
    return ShiftSaveResult(
        SHIFT_SAVE_SAVED,
        count=len(insert_params),
        changed_cells=changed_cells,
        message=f"月間シフト表の変更セルを保存しました。変更セル：{changed_cells}件",
    )


def create_shift_shortage_table(df, year, month, read_index=None):
    """日勤2名・夜勤1名を基準に、不足日を確認する。"""
    last_day = calendar.monthrange(int(year), int(month))[1]
    rows = []
    if read_index is None:
        read_index = build_shift_read_index(df)
    for d in range(1, last_day + 1):
        target_date = f"{int(year)}-{int(month):02d}-{d:02d}"
        counts = read_index["kind_counts_by_date"].get(target_date, {})
        day_count = sum(counts.get(kind, 0) for kind in DAY_STAFFING_SHIFT_KINDS)
        night_count = counts.get("夜勤", 0)
        status = "OK" if day_count >= 2 and night_count >= 1 else "要確認"
        rows.append({
            "日付": target_date,
            "日勤人数": day_count,
            "夜勤人数": night_count,
            "状態": status,
            "不足": f"日勤あと{max(0, 2-day_count)} / 夜勤あと{max(0, 1-night_count)}" if status != "OK" else "",
        })
    return pd.DataFrame(rows)


def create_shift_quality_check_table(df, year, month, read_index=None, matrix=None):
    """夜勤明け・連勤・希望休・勤務偏りを確認する。"""
    rows = []
    last_day = calendar.monthrange(int(year), int(month))[1]
    if df is None or df.empty:
        return pd.DataFrame(columns=["重要度", "種類", "日付", "職員名", "内容"])

    if read_index is None:
        read_index = build_shift_read_index(df)
    raw_kinds_by_staff_date = read_index["raw_kinds_by_staff_date"]
    staff_names = list(read_index["raw_staff_names"])
    for staff_name in staff_names:
        work_flags = []
        for d in range(1, last_day + 1):
            target_date = f"{int(year)}-{int(month):02d}-{d:02d}"
            staff_dates = raw_kinds_by_staff_date.get(staff_name, {})
            kinds = list(staff_dates.get(target_date, ()))

            # 希望休・有休に勤務が重なっている
            if any(k in kinds for k in OFF_OR_BLOCKING_SHIFT_KINDS) and any(k in kinds for k in ACTIVE_SHIFT_KINDS):
                rows.append({
                    "重要度": "高",
                    "種類": "休み希望と勤務の重複",
                    "日付": target_date,
                    "職員名": staff_name,
                    "内容": "休み・希望休・有休・その他と、日勤・夜勤・明けが同日に入っています。",
                })

            # 同じ日に日勤と夜勤が重なっている
            if "日勤" in kinds and "夜勤" in kinds:
                rows.append({
                    "重要度": "高",
                    "種類": "日勤夜勤重複",
                    "日付": target_date,
                    "職員名": staff_name,
                    "内容": "同じ日に日勤と夜勤が入っています。日/夜入力は禁止です。",
                })

            # 夜勤翌日は「明」、明け翌日は「休み」を基本にする
            try:
                prev_date = (date(int(year), int(month), d) - timedelta(days=1)).strftime("%Y-%m-%d")
                prev_kinds = list(staff_dates.get(prev_date, ()))

                if "夜勤" in prev_kinds:
                    if any(k in WORKDAY_SHIFT_KINDS for k in kinds):
                        rows.append({
                            "重要度": "高",
                            "種類": "夜勤翌日勤務",
                            "日付": target_date,
                            "職員名": staff_name,
                            "内容": "夜勤の翌日に勤務が入っています。翌日は明け扱いにしてください。",
                        })
                    if "夜勤明け" not in kinds:
                        rows.append({
                            "重要度": "中",
                            "種類": "夜勤翌日明け未登録",
                            "日付": target_date,
                            "職員名": staff_name,
                            "内容": "夜勤の翌日に「明」が登録されていません。",
                        })

                if "夜勤明け" in prev_kinds:
                    if any(k in WORKDAY_SHIFT_KINDS for k in kinds):
                        rows.append({
                            "重要度": "高",
                            "種類": "明け翌日勤務",
                            "日付": target_date,
                            "職員名": staff_name,
                            "内容": "明けの翌日に勤務が入っています。原則として休みにしてください。",
                        })
                    if not any(k in kinds for k in ["休み", "有休", "希望休"]):
                        rows.append({
                            "重要度": "中",
                            "種類": "明け翌日休み未登録",
                            "日付": target_date,
                            "職員名": staff_name,
                            "内容": "明けの翌日に休みが登録されていません。",
                        })
            except Exception:
                pass

            work_flags.append(1 if any(k in WORKDAY_SHIFT_KINDS for k in kinds) else 0)

        # 5連勤・6連勤以上
        cur = 0
        for idx, v in enumerate(work_flags, start=1):
            if v:
                cur += 1
                if cur == 5:
                    rows.append({
                        "重要度": "中",
                        "種類": "5連勤",
                        "日付": f"{int(year)}-{int(month):02d}-{idx:02d}",
                        "職員名": staff_name,
                        "内容": "5連勤になっています。疲労・調整を確認してください。",
                    })
                elif cur >= 6:
                    rows.append({
                        "重要度": "高",
                        "種類": f"{cur}連勤",
                        "日付": f"{int(year)}-{int(month):02d}-{idx:02d}",
                        "職員名": staff_name,
                        "内容": "6連勤以上になっています。要調整です。",
                    })
            else:
                cur = 0

    # 夜勤回数の偏り
    if matrix is None:
        matrix = create_shift_matrix(df, year, month, read_index)
    if not matrix.empty and "夜勤" in matrix.columns:
        night_vals = matrix["夜勤"].fillna(0).astype(int)
        if len(night_vals) >= 2:
            avg = night_vals.mean()
            for _, r in matrix.iterrows():
                if int(r["夜勤"]) >= avg + 3 and int(r["夜勤"]) >= 4:
                    rows.append({
                        "重要度": "中",
                        "種類": "夜勤回数偏り",
                        "日付": f"{int(year)}-{int(month):02d}",
                        "職員名": r["職員名"],
                        "内容": f"夜勤が{int(r['夜勤'])}回で、平均との差が大きい可能性があります。",
                    })

    return pd.DataFrame(rows, columns=["重要度", "種類", "日付", "職員名", "内容"])


def create_shift_duplicate_check_table(raw_df, year, month):
    """同一職員・同一日に複数の登録がある箇所を確定前に検出する。"""
    columns = ["重要度", "種類", "日付", "職員名", "内容"]
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=columns)
    work = raw_df.copy()
    work["_staff"] = work["staff_name"].apply(normalize_staff_name)
    work["_date"] = work["shift_date"].astype(str)
    month_prefix = f"{int(year)}-{int(month):02d}-"
    work = work[work["_date"].str.startswith(month_prefix) & work["_staff"].astype(bool)]
    rows = []
    for (shift_date, staff_name), group in work.groupby(["_date", "_staff"], sort=True):
        if len(group) <= 1:
            continue
        kinds = "、".join(group["shift_kind"].fillna("").astype(str).tolist())
        rows.append({
            "重要度": "高",
            "種類": "同一職員・同一日の重複",
            "日付": shift_date,
            "職員名": staff_name,
            "内容": f"同じ日に{len(group)}件の勤務が登録されています：{kinds}",
        })
    return pd.DataFrame(rows, columns=columns)


def build_shift_confirmation_errors(
    shortage_ng, quality_checks, limit_checks, duplicate_checks,
    kot_error_df=None, kot_generation_error=None,
):
    """既存チェックのうち、勤務表の確定を禁止する重大エラーを一覧化する。"""
    columns = ["区分", "重要度", "日付", "職員名", "内容"]
    rows = []

    for _, row in (shortage_ng if shortage_ng is not None else pd.DataFrame()).iterrows():
        rows.append({
            "区分": "必要人数不足", "重要度": "高",
            "日付": str(row.get("日付", "")), "職員名": "",
            "内容": str(row.get("不足", "必要人数を満たしていません。")),
        })

    for _, row in (duplicate_checks if duplicate_checks is not None else pd.DataFrame()).iterrows():
        rows.append({
            "区分": str(row.get("種類", "重複")), "重要度": "高",
            "日付": str(row.get("日付", "")), "職員名": str(row.get("職員名", "")),
            "内容": str(row.get("内容", "同一職員・同一日に重複があります。")),
        })

    high_checks = quality_checks if quality_checks is not None else pd.DataFrame()
    if not high_checks.empty and "重要度" in high_checks.columns:
        high_checks = high_checks[high_checks["重要度"].astype(str) == "高"]
    for _, row in high_checks.iterrows():
        rows.append({
            "区分": str(row.get("種類", "勤務チェック")), "重要度": "高",
            "日付": str(row.get("日付", "")), "職員名": str(row.get("職員名", "")),
            "内容": str(row.get("内容", "重大な勤務上の問題があります。")),
        })

    # 上限値は設定された最大値であり、重要度表示にかかわらず超過を確定禁止とする。
    for _, row in (limit_checks if limit_checks is not None else pd.DataFrame()).iterrows():
        rows.append({
            "区分": str(row.get("種類", "勤務回数上限超過")), "重要度": "高",
            "日付": "", "職員名": str(row.get("職員名", "")),
            "内容": str(row.get("内容", "勤務回数上限を超えています。")),
        })

    for _, row in (kot_error_df if kot_error_df is not None else pd.DataFrame()).iterrows():
        detail_parts = []
        for column in row.index:
            value = row.get(column, "")
            if pd.isna(value) or not str(value).strip():
                continue
            detail_parts.append(f"{column}:{value}")
        detail = " / ".join(detail_parts)
        rows.append({
            "区分": "KING OF TIME重大エラー", "重要度": "高",
            "日付": str(row.get("勤務日", row.get("日付", ""))),
            "職員名": str(row.get("職員名", row.get("名前", ""))),
            "内容": detail or "KING OF TIME出力に重大エラーがあります。",
        })
    if kot_generation_error:
        rows.append({
            "区分": "KING OF TIME重大エラー", "重要度": "高",
            "日付": "", "職員名": "", "内容": str(kot_generation_error),
        })
    return pd.DataFrame(rows, columns=columns)


def confirm_shift_month_if_valid(year, month, confirmation_errors, confirmed_by=""):
    """重大エラーが0件の場合だけ対象月を確定する。"""
    status = get_shift_month_status(int(year), int(month))
    if status.get("status_error"):
        raise ShiftUpdateBlockedError(
            f"{int(year)}年{int(month)}月の確定状態を確認できないため、確定処理を実行できません。"
        )
    if status.get("is_confirmed"):
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="この月はすでに確定済みです。")
    if confirmation_errors is not None and not confirmation_errors.empty:
        return ShiftSaveResult(
            SHIFT_SAVE_BLOCKED,
            count=len(confirmation_errors),
            message="重大な問題があるため勤務表を確定できません。内容を修正して再確認してください。",
        )
    set_shift_month_status(int(year), int(month), True, confirmed_by)
    return ShiftSaveResult(SHIFT_SAVE_SAVED, count=1, message="この月のシフトを確定しました。")


def parse_hour_from_time_text(value):
    text = str(value or "").strip()
    if not text:
        return None
    m = re.match(r"^(\d{1,2})[:：](\d{2})", text)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def get_staff_event_counts_for_date(target_date):
    df = fetch_df("""
        SELECT staff_name, COUNT(*) AS cnt
        FROM events
        WHERE event_date = ? AND staff_name IS NOT NULL AND staff_name <> ''
        GROUP BY staff_name
    """, (str(target_date),))
    if df.empty:
        return {}
    return {str(r["staff_name"]): int(r["cnt"]) for _, r in df.iterrows()}


@st.cache_data(ttl=60, show_spinner=False)
def get_staff_event_counts_month(year, month):
    """対象月の予定担当件数を日付・職員別に1回のSQLで取得する。"""
    year = int(year)
    month = int(month)
    start = date(year, month, 1)
    next_month = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
    df = fetch_df("""
        SELECT event_date, staff_name, COUNT(*) AS event_count
        FROM events
        WHERE event_date >= ? AND event_date < ?
          AND staff_name IS NOT NULL AND staff_name <> ''
        GROUP BY event_date, staff_name
    """, (start.strftime("%Y-%m-%d"), next_month.strftime("%Y-%m-%d")))
    counts_by_date = {}
    if df is None or df.empty:
        return counts_by_date
    for _, row in df.iterrows():
        event_date = str(row["event_date"])
        staff_name = str(row["staff_name"])
        counts_by_date.setdefault(event_date, {})[staff_name] = int(row["event_count"])
    return counts_by_date


def _get_shift_candidates_for_event(event_row, shift_df, counts):
    """取得済みの日別担当件数を使って予定の担当候補を作る。"""
    if shift_df is None or shift_df.empty:
        return pd.DataFrame()

    event_date = str(event_row["event_date"])
    start_hour = parse_hour_from_time_text(event_row.get("start_time", ""))
    night_event = start_hour is not None and (start_hour >= 16 or start_hour <= 9)
    day_df = shift_df[shift_df["shift_date"].astype(str) == event_date].copy()
    if day_df.empty:
        return pd.DataFrame()

    # 休み系・明けは候補外
    day_df = day_df[~day_df["shift_kind"].isin(["休み", "希望休", "有休", "夜勤明け"])].copy()
    if day_df.empty:
        return pd.DataFrame()

    rows = []
    for _, r in day_df.iterrows():
        staff_name = str(r["staff_name"])
        shift_kind = str(r["shift_kind"])
        score = 50
        reason = ["勤務表に登録あり"]

        if night_event:
            if shift_kind == "夜勤":
                score += 40
                reason.append("夜間予定と夜勤が一致")
            elif shift_kind == "日勤":
                score -= 10
                reason.append("夜間予定だが日勤")
        else:
            if shift_kind == "日勤":
                score += 35
                reason.append("日中予定と日勤が一致")
            elif shift_kind == "夜勤":
                score -= 5
                reason.append("日中予定だが夜勤")

        assigned = counts.get(staff_name, 0)
        if assigned:
            score -= assigned * 8
            reason.append(f"同日すでに{assigned}件担当")

        rows.append({
            "職員名": staff_name,
            "シフト": shift_kind,
            "時間": f"{format_time_for_display(r.get('start_time'))}〜{format_time_for_display(r.get('end_time'))}{'翌' if int(r['next_day'] or 0) else ''}",
            "スコア": int(score),
            "理由": " / ".join(reason),
        })

    if not rows:
        return pd.DataFrame()
    result = pd.DataFrame(rows)
    result = result.sort_values(["職員名", "スコア"], ascending=[True, False]).groupby("職員名", as_index=False).first()
    return result.sort_values(["スコア", "職員名"], ascending=[False, True]).reset_index(drop=True)


def get_shift_candidates_for_event(event_row, shift_df):
    """予定日時に対して担当候補を出す。日中は日勤、夕方以降・早朝は夜勤を優先する。"""
    if shift_df is None or shift_df.empty:
        return pd.DataFrame()
    event_date = str(event_row["event_date"])
    counts = get_staff_event_counts_for_date(event_date)
    return _get_shift_candidates_for_event(event_row, shift_df, counts)


def build_event_assignment_preview(events_df, shift_df, only_unassigned=True):
    if events_df is None or events_df.empty:
        return pd.DataFrame()
    counts_by_date = {}
    loaded_months = set()
    for raw_event_date in events_df["event_date"].tolist():
        parsed = pd.to_datetime(raw_event_date, errors="coerce")
        if pd.isna(parsed):
            continue
        year_month = (int(parsed.year), int(parsed.month))
        if year_month in loaded_months:
            continue
        loaded_months.add(year_month)
        counts_by_date.update(get_staff_event_counts_month(*year_month))
    return _build_event_assignment_preview(events_df, shift_df, counts_by_date, only_unassigned)


def _build_event_assignment_preview(events_df, shift_df, counts_by_date, only_unassigned=True):
    rows = []
    if events_df is None or events_df.empty:
        return pd.DataFrame()
    for _, ev in events_df.iterrows():
        current = str(ev["staff_name"] or "")
        if only_unassigned and current:
            continue
        event_date = str(ev["event_date"])
        cand = _get_shift_candidates_for_event(ev, shift_df, counts_by_date.get(event_date, {}))
        if cand.empty:
            rows.append({
                "予定ID": int(ev["id"]),
                "日付": ev["event_date"],
                "予定": ev["title"],
                "現在担当": current,
                "AI候補": "",
                "スコア": 0,
                "理由": "候補なし",
            })
        else:
            top = cand.iloc[0]
            rows.append({
                "予定ID": int(ev["id"]),
                "日付": ev["event_date"],
                "予定": ev["title"],
                "現在担当": current,
                "AI候補": top["職員名"],
                "スコア": int(top["スコア"]),
                "理由": top["理由"],
            })
    return pd.DataFrame(rows)



OFF_OR_BLOCKING_SHIFT_KINDS = ["休み", "希望休", "有休", "その他"]
ACTIVE_SHIFT_KINDS = WORKDAY_SHIFT_KINDS


def get_staff_day_shift_kinds(df, staff_name, target_date):
    """指定職員・指定日の勤務区分一覧を、職員名正規化込みで取得する。"""
    if df is None or df.empty:
        return []
    target_staff = normalize_staff_name(staff_name)
    tmp = df.copy()
    tmp["_staff_norm"] = tmp["staff_name"].apply(normalize_staff_name)
    day_df = tmp[
        (tmp["_staff_norm"] == target_staff) &
        (tmp["shift_date"].astype(str) == str(target_date))
    ]
    if day_df.empty:
        return []
    return [str(x) for x in day_df["shift_kind"].dropna().astype(str).tolist()]


def can_add_shift_without_overlap(existing_kinds, new_kind):
    """
    同じ日に入れてよい勤務かを判定する。
    希望休・有休・休み・その他がある日は、日勤・夜勤・明けを入れない。
    日勤・夜勤・明けがある日は、希望休・有休・休み・その他を重ねない。
    """
    existing_kinds = [str(k) for k in (existing_kinds or []) if str(k)]
    new_kind = str(new_kind or "").strip()
    if not new_kind:
        return False, "勤務区分が空です"

    if new_kind in existing_kinds:
        return False, f"{new_kind}が既に登録済み"

    existing_off = [k for k in existing_kinds if k in OFF_OR_BLOCKING_SHIFT_KINDS]
    existing_active = [k for k in existing_kinds if k in ACTIVE_SHIFT_KINDS]

    if new_kind in ACTIVE_SHIFT_KINDS and existing_off:
        return False, f"休み系（{', '.join(existing_off)}）が既にあるため{new_kind}は入れません"
    if new_kind in OFF_OR_BLOCKING_SHIFT_KINDS and existing_active:
        return False, f"勤務系（{', '.join(existing_active)}）が既にあるため{new_kind}は入れません"

    # 日勤・夜勤・明けは同日で重ねない
    if new_kind in ACTIVE_SHIFT_KINDS and existing_active:
        return False, f"勤務系（{', '.join(existing_active)}）が既にあるため{new_kind}は入れません"

    # 休み系同士も重ねない。希望休や有休が入っていれば休み扱いとして十分。
    if new_kind in OFF_OR_BLOCKING_SHIFT_KINDS and existing_off:
        return False, f"休み系（{', '.join(existing_off)}）が既にあるため重ねません"

    return True, "追加可能"


def can_add_shift_to_working_df(df, staff_name, target_date, new_kind):
    existing_kinds = get_staff_day_shift_kinds(df, staff_name, target_date)
    return can_add_shift_without_overlap(existing_kinds, new_kind)



def has_off_shift_on_day(df, staff_name, target_date):
    """希望休・有休・休み・その他が同日にあるか。"""
    kinds = get_staff_day_shift_kinds(df, staff_name, target_date)
    return any(k in OFF_OR_BLOCKING_SHIFT_KINDS for k in kinds), [k for k in kinds if k in OFF_OR_BLOCKING_SHIFT_KINDS]


def has_active_shift_on_day(df, staff_name, target_date):
    """日勤・夜勤・夜勤明けが同日にあるか。"""
    kinds = get_staff_day_shift_kinds(df, staff_name, target_date)
    return any(k in ACTIVE_SHIFT_KINDS for k in kinds), [k for k in kinds if k in ACTIVE_SHIFT_KINDS]


def staff_available_for_kind(df, staff_name, target_date, shift_kind, limit_map=None):
    """
    AIシフト案用。
    希望休・有休・その他・休みと、日勤・夜勤・夜勤明けが同日に重ならないよう厳密に見る。
    特に夜勤は、翌日が明けになるため、翌日に希望休等がある場合は候補外にする。
    """
    if df is None:
        df = pd.DataFrame()
    date_text = str(target_date)
    target_staff = normalize_staff_name(staff_name)

    existing_kinds = get_staff_day_shift_kinds(df, target_staff, date_text)
    overlap_ok, overlap_reason = can_add_shift_without_overlap(existing_kinds, shift_kind)
    if not overlap_ok:
        return False, overlap_reason

    try:
        prev_date = (pd.to_datetime(date_text).date() - timedelta(days=1)).strftime("%Y-%m-%d")
        prev_kinds = get_staff_day_shift_kinds(df, target_staff, prev_date)
        if "夜勤" in prev_kinds:
            return False, "前日夜勤のため翌日は明け"
        if "夜勤明け" in prev_kinds:
            return False, "前日明けのため翌日は休み"
    except Exception:
        pass

    # 夜勤を入れると翌日は必ず明け扱いになるため、翌日に希望休・有休・休み・その他がある場合は夜勤候補から外す。
    if shift_kind == "夜勤":
        try:
            next_date = (pd.to_datetime(date_text).date() + timedelta(days=1)).strftime("%Y-%m-%d")
            next_kinds = get_staff_day_shift_kinds(df, target_staff, next_date)
            next_off = [k for k in next_kinds if k in OFF_OR_BLOCKING_SHIFT_KINDS]
            next_active = [k for k in next_kinds if k in ACTIVE_SHIFT_KINDS]
            if next_off:
                return False, f"翌日に休み系（{', '.join(next_off)}）があるため夜勤不可"
            if next_active:
                return False, f"翌日に勤務系（{', '.join(next_active)}）があるため夜勤不可"

            rest_date = (pd.to_datetime(date_text).date() + timedelta(days=2)).strftime("%Y-%m-%d")
            rest_kinds = get_staff_day_shift_kinds(df, target_staff, rest_date)
            rest_active = [k for k in rest_kinds if k in ACTIVE_SHIFT_KINDS]
            if rest_active:
                return False, f"明け翌日に勤務系（{', '.join(rest_active)}）があるため夜勤不可"
        except Exception:
            pass

    exceed, reason = would_exceed_staff_shift_limit(df, target_staff, date_text, shift_kind, limit_map=limit_map)
    if exceed:
        return False, reason

    return True, "候補"


def summarize_ai_reject_reasons(reject_reasons):
    """AI候補から外れた理由を、現場で読みやすいカテゴリ別に要約する。"""
    buckets = {
        "上限": 0,
        "休み系": 0,
        "夜勤明け": 0,
        "明け翌日": 0,
        "重複": 0,
        "その他": 0,
    }
    samples = []
    for reason in reject_reasons or []:
        text = str(reason)
        if len(samples) < 4:
            samples.append(text)
        if "上限" in text:
            buckets["上限"] += 1
        elif "前日夜勤" in text or "翌日は明け" in text or "夜勤明け" in text:
            buckets["夜勤明け"] += 1
        elif "前日明け" in text or "明け翌日" in text:
            buckets["明け翌日"] += 1
        elif "休み系" in text or "希望休" in text or "有休" in text or "休み" in text:
            buckets["休み系"] += 1
        elif "既に" in text or "重複" in text or "勤務系" in text:
            buckets["重複"] += 1
        else:
            buckets["その他"] += 1

    parts = [f"{name}{count}名" for name, count in buckets.items() if count]
    summary = "、".join(parts) if parts else "候補理由なし"
    if samples:
        summary += "。例：" + " / ".join(samples)
    return summary


def create_ai_shift_draft(df, staff_names, year, month):
    """
    ルール型AIで不足分のシフト案を作る。
    既存シフトとAI追加分を都度数え直し、日勤上限・夜勤上限・合計上限を絶対に超えない範囲だけで候補化する。
    """
    last_day = calendar.monthrange(int(year), int(month))[1]
    # 職員名の空白違いを統一。上限表と候補職員の名前がズレても判定漏れしないようにする。
    staff_names = sorted(set([normalize_staff_name(s) for s in staff_names if normalize_staff_name(s)]))
    if not staff_names:
        return pd.DataFrame()

    limit_map = get_staff_shift_limit_map()

    working_df = df.copy() if df is not None else pd.DataFrame()
    if working_df is None or working_df.empty:
        working_df = pd.DataFrame(columns=["shift_date", "staff_name", "shift_kind", "start_time", "end_time", "next_day", "memo", "created_at", "updated_at"])
    else:
        working_df["staff_name"] = working_df["staff_name"].apply(normalize_staff_name)

    rows = []

    # まず、既存シフトが既に上限超過している職員を明示する。
    for s in staff_names:
        limits = normalize_shift_limits_for_staff(limit_map, s)
        d_count, n_count, t_count = staff_month_work_counts(working_df, s, int(year), int(month))
        if d_count > limits["day"] or n_count > limits["night"] or t_count > limits["total"]:
            rows.append({
                "日付": f"{int(year)}-{int(month):02d}-01",
                "勤務": "注意",
                "候補職員": s,
                "理由": f"既存シフトが上限超過しています（日勤 {d_count}/{limits['day']}、夜勤 {n_count}/{limits['night']}、合計 {t_count}/{limits['total']}）。AIはこの職員へ追加しません。",
                "保存対象": False,
            })

    for d in range(1, last_day + 1):
        target_date = f"{int(year)}-{int(month):02d}-{d:02d}"
        day_df = working_df[working_df["shift_date"].astype(str) == target_date] if not working_df.empty else pd.DataFrame()
        current_day = int(len(day_df[day_df["shift_kind"].astype(str) == "日勤"])) if not day_df.empty else 0
        current_night = int(len(day_df[day_df["shift_kind"].astype(str) == "夜勤"])) if not day_df.empty else 0

        need_list = ["日勤"] * max(0, 2 - current_day) + ["夜勤"] * max(0, 1 - current_night)

        for need_kind in need_list:
            candidates = []
            reject_reasons = []

            for s in staff_names:
                limits = normalize_shift_limits_for_staff(limit_map, s)
                d_count, n_count, t_count = staff_month_work_counts(working_df, s, int(year), int(month))

                # 上限を先に見る。ここで弾くので、夜勤0回の職員に夜勤が入ることはない。
                if need_kind == "日勤" and d_count >= limits["day"]:
                    reject_reasons.append(f"{s}: 日勤上限{limits['day']}回")
                    continue
                if need_kind == "夜勤" and n_count >= limits["night"]:
                    reject_reasons.append(f"{s}: 夜勤上限{limits['night']}回")
                    continue
                if t_count >= limits["total"]:
                    reject_reasons.append(f"{s}: 合計上限{limits['total']}回")
                    continue

                ok, why = staff_available_for_kind(working_df, s, target_date, need_kind, limit_map=limit_map)
                if not ok:
                    reject_reasons.append(f"{s}: {why}")
                    continue

                # 上限内の職員だけを、少ない人優先で点数化
                score = 100
                score -= t_count * 3
                if need_kind == "日勤":
                    score -= d_count * 2
                    # 日勤残り枠が少ない人は少し下げる
                    score -= max(0, 3 - (limits["day"] - d_count))
                if need_kind == "夜勤":
                    score -= n_count * 5
                    score -= max(0, 3 - (limits["night"] - n_count)) * 2

                try:
                    consecutive_before = 0
                    for back in range(1, 6):
                        prev_date = (pd.to_datetime(target_date).date() - timedelta(days=back)).strftime("%Y-%m-%d")
                        p_df = working_df[
                            (working_df["staff_name"].apply(normalize_staff_name) == s) &
                            (working_df["shift_date"].astype(str) == prev_date)
                        ] if not working_df.empty else pd.DataFrame()
                        if not p_df.empty and any(k in p_df["shift_kind"].astype(str).tolist() for k in WORKDAY_SHIFT_KINDS):
                            consecutive_before += 1
                        else:
                            break
                    if consecutive_before >= 4:
                        score -= 60
                except Exception:
                    pass

                candidates.append((score, s))

            if not candidates:
                reason_text = "日勤・夜勤・合計上限、希望休、有休、休み、その他、夜勤翌日の明け重複を考慮した結果、候補なし"
                if reject_reasons:
                    reason_text += "（" + summarize_ai_reject_reasons(reject_reasons) + "）"
                rows.append({
                    "日付": target_date,
                    "勤務": need_kind,
                    "候補職員": "",
                    "理由": reason_text,
                    "保存対象": False,
                })
                continue

            candidates.sort(reverse=True)
            selected = candidates[0][1]

            # 最終ガード：追加直前にもう一度、現在のworking_dfで上限超過を確認する
            exceed, limit_reason = would_exceed_staff_shift_limit(
                working_df, selected, target_date, need_kind, limit_map=limit_map
            )
            if exceed:
                rows.append({
                    "日付": target_date,
                    "勤務": need_kind,
                    "候補職員": "",
                    "理由": limit_reason,
                    "保存対象": False,
                })
                continue

            rows.append({
                "日付": target_date,
                "勤務": need_kind,
                "候補職員": selected,
                "理由": "日勤・夜勤・合計上限内で候補化",
                "保存対象": True,
            })

            stime, etime, nd = default_shift_times(need_kind)
            add_row = pd.DataFrame([{
                "shift_date": target_date,
                "staff_name": selected,
                "shift_kind": need_kind,
                "start_time": stime,
                "end_time": etime,
                "next_day": nd,
                "memo": "AIシフト案",
                "created_at": now_text(),
                "updated_at": now_text(),
            }])
            working_df = pd.concat([working_df, add_row], ignore_index=True)

            if need_kind == "夜勤":
                # 夜勤の翌日は「明」、明けの翌日は「休み」を自動で入れる
                try:
                    next_date_obj = pd.to_datetime(target_date).date() + timedelta(days=1)
                    rest_date_obj = pd.to_datetime(target_date).date() + timedelta(days=2)

                    if next_date_obj.month == int(month):
                        next_date = next_date_obj.strftime("%Y-%m-%d")
                        next_kinds = get_staff_day_shift_kinds(working_df, selected, next_date)
                        can_add_ake, ake_reason = can_add_shift_without_overlap(next_kinds, "夜勤明け")
                        if not can_add_ake:
                            # 既に夜勤明けがある場合も重複追加しない。休み・希望休・有休・その他がある場合は明けを重ねない。
                            rows.append({
                                "日付": next_date,
                                "勤務": "注意",
                                "候補職員": selected,
                                "理由": f"夜勤翌日の明けは追加しません：{ake_reason}",
                                "保存対象": False,
                            })
                        else:
                            rows.append({
                                "日付": next_date,
                                "勤務": "夜勤明け",
                                "候補職員": selected,
                                "理由": "夜勤翌日の明けを自動付与",
                                "保存対象": True,
                            })
                            ake_row = pd.DataFrame([{
                                "shift_date": next_date,
                                "staff_name": selected,
                                "shift_kind": "夜勤明け",
                                "start_time": "",
                                "end_time": "",
                                "next_day": 0,
                                "memo": "夜勤翌日の明け自動付与",
                                "created_at": now_text(),
                                "updated_at": now_text(),
                            }])
                            working_df = pd.concat([working_df, ake_row], ignore_index=True)
                    else:
                        rows.append({
                            "日付": next_date_obj.strftime("%Y-%m-%d"),
                            "勤務": "注意",
                            "候補職員": selected,
                            "理由": "夜勤翌日の明けが翌月になるため、翌月シフトで確認してください。",
                            "保存対象": False,
                        })

                    if rest_date_obj.month == int(month):
                        rest_date = rest_date_obj.strftime("%Y-%m-%d")
                        rest_kinds = get_staff_day_shift_kinds(working_df, selected, rest_date)
                        can_add_rest, rest_reason = can_add_shift_without_overlap(rest_kinds, "休み")
                        if not can_add_rest:
                            # 希望休・有休・その他が既にある日は、それを休み扱いとして休みを重ねない。
                            rows.append({
                                "日付": rest_date,
                                "勤務": "注意",
                                "候補職員": selected,
                                "理由": f"明け翌日の休みは追加しません：{rest_reason}",
                                "保存対象": False,
                            })
                        else:
                            rows.append({
                                "日付": rest_date,
                                "勤務": "休み",
                                "候補職員": selected,
                                "理由": "明け翌日の休みを自動付与",
                                "保存対象": True,
                            })
                            rest_row = pd.DataFrame([{
                                "shift_date": rest_date,
                                "staff_name": selected,
                                "shift_kind": "休み",
                                "start_time": "",
                                "end_time": "",
                                "next_day": 0,
                                "memo": "明け翌日の休み自動付与",
                                "created_at": now_text(),
                                "updated_at": now_text(),
                            }])
                            working_df = pd.concat([working_df, rest_row], ignore_index=True)
                    else:
                        rows.append({
                            "日付": rest_date_obj.strftime("%Y-%m-%d"),
                            "勤務": "注意",
                            "候補職員": selected,
                            "理由": "明け翌日の休みが翌月になるため、翌月シフトで確認してください。",
                            "保存対象": False,
                        })
                except Exception as e:
                    rows.append({
                        "日付": target_date,
                        "勤務": "注意",
                        "候補職員": selected,
                        "理由": f"夜勤後の明け・休み自動付与に失敗しました：{e}",
                        "保存対象": False,
                    })

    return pd.DataFrame(rows)

def apply_ai_shift_draft_to_df(base_df, draft_df):
    """AIシフト案をDB保存前に画面上で仮反映したDataFrameを作る。重複する案は仮反映しない。"""
    work = base_df.copy() if base_df is not None else pd.DataFrame()
    if work is not None and not work.empty and "staff_name" in work.columns:
        work["staff_name"] = work["staff_name"].apply(normalize_staff_name)
    if draft_df is None or draft_df.empty:
        return work
    if work is None or work.empty:
        work = pd.DataFrame(columns=["id", "shift_date", "staff_name", "shift_kind", "start_time", "end_time", "next_day", "memo", "created_at", "updated_at"])

    rows = []
    for _, r in draft_df.iterrows():
        if not bool(r.get("保存対象", False)):
            continue
        staff_name = normalize_staff_name(r.get("候補職員", ""))
        shift_kind = str(r.get("勤務", "")).strip()
        shift_date = str(r.get("日付", "")).strip()
        if not staff_name or not shift_kind or not shift_date:
            continue
        ok, _ = can_add_shift_to_working_df(work, staff_name, shift_date, shift_kind)
        if not ok:
            continue
        stime, etime, nd = default_shift_times(shift_kind)
        new_row = {
            "id": -1,
            "shift_date": shift_date,
            "staff_name": staff_name,
            "shift_kind": shift_kind,
            "start_time": stime,
            "end_time": etime,
            "next_day": nd,
            "memo": "AIシフト案（未保存）",
            "created_at": now_text(),
            "updated_at": now_text(),
        }
        rows.append(new_row)
        work = pd.concat([work, pd.DataFrame([new_row])], ignore_index=True)
    return work



def sanitize_ai_shift_draft_by_limits(draft_df, base_df, year, month):
    """
    session_stateに残った古いAI案も、現在の上限・重複条件で再検査する。
    希望休・有休・休み・その他がある日は、日勤・夜勤・明けを重ねない。
    """
    if draft_df is None or draft_df.empty:
        return draft_df

    working_df = base_df.copy() if base_df is not None else pd.DataFrame()
    if working_df is None or working_df.empty:
        working_df = pd.DataFrame(columns=["shift_date", "staff_name", "shift_kind", "start_time", "end_time", "next_day", "memo", "created_at", "updated_at"])
    else:
        working_df["staff_name"] = working_df["staff_name"].apply(normalize_staff_name)

    limit_map = get_staff_shift_limit_map()
    cleaned_rows = []

    for _, r in draft_df.iterrows():
        shift_date = str(r.get("日付", "")).strip()
        shift_kind = str(r.get("勤務", "")).strip()
        staff_name = normalize_staff_name(r.get("候補職員", ""))
        save_target = bool(r.get("保存対象", False))

        if save_target and staff_name and shift_kind:
            # 全勤務区分について同日重複・休み系との重なりを再判定
            overlap_ok, overlap_reason = can_add_shift_to_working_df(working_df, staff_name, shift_date, shift_kind)
            if not overlap_ok:
                cleaned_rows.append({
                    "日付": shift_date,
                    "勤務": "注意",
                    "候補職員": staff_name,
                    "理由": f"現在の重複条件で除外しました：{shift_kind} / {overlap_reason}",
                    "保存対象": False,
                })
                continue

            # 日勤・夜勤は上限も再判定
            if shift_kind in ["日勤", "夜勤"]:
                ok, why = staff_available_for_kind(working_df, staff_name, shift_date, shift_kind, limit_map=limit_map)
                if not ok:
                    cleaned_rows.append({
                        "日付": shift_date,
                        "勤務": "注意",
                        "候補職員": staff_name,
                        "理由": f"現在の上限条件で除外しました：{shift_kind} / {why}",
                        "保存対象": False,
                    })
                    continue

            cleaned_rows.append({
                "日付": shift_date,
                "勤務": shift_kind,
                "候補職員": staff_name,
                "理由": str(r.get("理由", "現在条件で確認済み")),
                "保存対象": True,
            })
            stime, etime, nd = default_shift_times(shift_kind)
            add_row = pd.DataFrame([{
                "shift_date": shift_date,
                "staff_name": staff_name,
                "shift_kind": shift_kind,
                "start_time": stime,
                "end_time": etime,
                "next_day": nd,
                "memo": "AIシフト案",
                "created_at": now_text(),
                "updated_at": now_text(),
            }])
            working_df = pd.concat([working_df, add_row], ignore_index=True)
            continue

        row_dict = dict(r)
        if "候補職員" in row_dict:
            row_dict["候補職員"] = staff_name
        cleaned_rows.append(row_dict)

    return pd.DataFrame(cleaned_rows, columns=["日付", "勤務", "候補職員", "理由", "保存対象"])



def save_ai_shift_draft_rows(draft_df):
    """
    AIシフト案を保存する直前にも再検査する。
    画面表示後に上限やシフトが変更されても、DB保存時に日勤・夜勤・合計上限を超えないようにする。
    """
    if draft_df is None or draft_df.empty:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="保存するAIシフト案がありません。")

    save_targets = draft_df[draft_df.get("保存対象", False).astype(bool)].copy()
    if save_targets.empty:
        return ShiftSaveResult(SHIFT_SAVE_NO_CHANGE, message="保存対象にできるAIシフト案がありません。")

    # DB更新より前に、保存対象となる全日付の確定状態を検査する。
    checked_months = set()
    for raw_date in save_targets["日付"].astype(str).tolist():
        target = pd.to_datetime(raw_date, errors="raise").date()
        key = (target.year, target.month)
        if key not in checked_months:
            ensure_shift_month_editable(*key)
            checked_months.add(key)

    # 複数月が混在する可能性は低いが、安全のため日付順に処理する
    save_targets["_sort_date"] = save_targets["日付"].astype(str)
    save_targets = save_targets.sort_values(["_sort_date", "勤務", "候補職員"])

    first_date = pd.to_datetime(str(save_targets.iloc[0]["日付"])).date()
    working_df = get_staff_shifts_month(first_date.year, first_date.month, include_prev_day=True)
    if working_df is not None and not working_df.empty and "staff_name" in working_df.columns:
        working_df["staff_name"] = working_df["staff_name"].apply(normalize_staff_name)
    limit_map = get_staff_shift_limit_map()

    params = []
    for _, r in save_targets.iterrows():
        staff_name = normalize_staff_name(r.get("候補職員", ""))
        shift_kind = str(r.get("勤務", "")).strip()
        shift_date = str(r.get("日付", "")).strip()
        if not staff_name or not shift_kind or not shift_date:
            continue

        # すべての勤務区分で、同日重複・休み系との重なりを再確認
        overlap_ok, overlap_reason = can_add_shift_to_working_df(working_df, staff_name, shift_date, shift_kind)
        if not overlap_ok:
            continue

        # 日勤・夜勤は上限、夜勤翌日等も再確認
        if shift_kind in ["日勤", "夜勤"]:
            ok, why = staff_available_for_kind(working_df, staff_name, shift_date, shift_kind, limit_map=limit_map)
            if not ok:
                continue

        stime, etime, nd = default_shift_times(shift_kind)
        params.append((shift_date, staff_name, shift_kind, stime or None, etime or None, int(nd or 0), "AIシフト案から保存", now_text(), now_text()))

        add_row = pd.DataFrame([{
            "shift_date": shift_date,
            "staff_name": staff_name,
            "shift_kind": shift_kind,
            "start_time": stime,
            "end_time": etime,
            "next_day": nd,
            "memo": "AIシフト案から保存",
            "created_at": now_text(),
            "updated_at": now_text(),
        }])
        working_df = pd.concat([working_df, add_row], ignore_index=True)

    params = dedupe_shift_insert_params(params)
    if not params:
        return ShiftSaveResult(
            SHIFT_SAVE_DUPLICATE,
            message="現在の勤務や上限と重複するため、保存できるAIシフト案はありませんでした。",
        )
    saved = execute_many("""
        INSERT INTO staff_shifts
        (shift_date, staff_name, shift_kind, start_time, end_time, next_day, memo, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, params)
    clear_shift_caches()
    return ShiftSaveResult(
        SHIFT_SAVE_SAVED,
        count=saved,
        message=f"AIシフト案を {saved} 件保存しました。",
    )

def make_staff_shift_pdf(year, month):
    return report_make_staff_shift_pdf(
        year,
        month,
        get_staff_shifts_month,
        create_shift_matrix,
        create_shift_shortage_table,
        create_shift_quality_check_table,
        create_shift_limit_check_table,
        get_shift_month_status,
    )


def make_shift_calendar_pdf(year, month, shift_df, staff_list, selected_staff_names=None, finalized=False):
    if not hasattr(report_service, "make_shift_calendar_pdf"):
        raise RuntimeError("月間シフトカレンダーPDF関数が report_service.py に見つかりません。")
    return report_service.make_shift_calendar_pdf(
        year,
        month,
        shift_df,
        staff_list=staff_list,
        selected_staff_names=selected_staff_names,
        finalized=finalized,
    )


def make_king_of_time_shift_csv(year, month, selected_staff_keys=None):
    return report_make_king_of_time_shift_csv(
        year,
        month,
        get_staff_shifts,
        get_staff_code_map,
        normalize_staff_name,
        default_shift_times,
        selected_staff_keys=selected_staff_keys,
        get_staff_key_map=get_staff_key_map,
    )


def build_king_of_time_clock_export(year, month, selected_staff_keys=None):
    return report_build_king_of_time_clock_export(
        year,
        month,
        get_staff_shifts,
        get_staff_code_map,
        normalize_staff_name,
        default_shift_times=default_shift_times,
        selected_staff_keys=selected_staff_keys,
        get_staff_key_map=get_staff_key_map,
    )


def make_staff_shift_excel(year, month):
    return report_make_staff_shift_excel(
        year,
        month,
        get_staff_shifts_month,
        create_shift_matrix,
        create_shift_shortage_table,
        create_shift_quality_check_table,
        create_shift_limit_check_table,
        get_shift_month_status,
    )


SHIFT_REPORT_SNAPSHOT_COLUMNS = (
    "id", "shift_date", "staff_name", "shift_kind", "start_time", "end_time", "next_day",
)


def _normalize_shift_snapshot_value(column, value):
    """シフト帳票スナップショットの値を型や欠損表現に依存しない文字列へそろえる。"""
    if value is None or (not isinstance(value, (list, tuple, dict, set)) and pd.isna(value)):
        return ""
    if column == "shift_date":
        parsed = pd.to_datetime(value, errors="coerce")
        return parsed.strftime("%Y-%m-%d") if not pd.isna(parsed) else str(value).strip()
    if column in ("id", "next_day"):
        try:
            return str(int(float(value)))
        except (TypeError, ValueError):
            return str(value).strip()
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    return str(value)


def _hash_snapshot_rows(rows):
    """行全体の巨大なreprを作らず、境界を含めて安定したSHA-256を計算する。"""
    digest = hashlib.sha256()
    for row in rows:
        for value in row:
            encoded = str(value).encode("utf-8")
            digest.update(len(encoded).to_bytes(8, "big"))
            digest.update(encoded)
        digest.update(b"\xff")
    return digest.hexdigest()


def _make_shift_report_snapshot(month_rows, kot_detail_rows, previous_context_rows):
    month_rows = tuple(sorted(month_rows))
    kot_detail_rows = tuple(sorted(kot_detail_rows))
    previous_context_rows = tuple(sorted(previous_context_rows))
    return {
        "month_rows": month_rows,
        "base_hash": _hash_snapshot_rows(month_rows),
        "kot_detail_hash": _hash_snapshot_rows(kot_detail_rows),
        "previous_context_hash": _hash_snapshot_rows(previous_context_rows),
    }


def _build_shift_report_snapshot(shift_df, year, month):
    """画面用の読取インデックスと同じ1回の走査で帳票署名も作る。"""
    return build_shift_read_index(shift_df, year, month)["report_snapshot"]


def _stable_rows_for_hash(df, columns):
    if df is None or df.empty:
        return ()
    available = [column for column in columns if column in df.columns]
    rows = [
        tuple(_normalize_shift_snapshot_value(column, value) for column, value in zip(available, values))
        for values in df[available].itertuples(index=False, name=None)
    ]
    return tuple(sorted(rows))


def _hash_cache_parts(*parts):
    rows = []
    for index, part in enumerate(parts):
        if isinstance(part, tuple):
            rows.extend((str(index),) + tuple(str(value) for value in row) if isinstance(row, tuple) else (str(index), str(row)) for row in part)
        else:
            rows.append((str(index), str(part)))
    return _hash_snapshot_rows(rows)


def _shift_report_signature(shift_snapshot, month_status, limits_df):
    limit_rows = _stable_rows_for_hash(limits_df, ["職員名", "日勤上限", "夜勤上限", "合計上限", "メモ"])
    status_key = tuple(sorted((str(key), str(value or "")) for key, value in (month_status or {}).items()))
    staffing_rules = (
        tuple(DAY_STAFFING_SHIFT_KINDS), tuple(DAY_LIMIT_SHIFT_KINDS),
        tuple(NIGHT_LIMIT_SHIFT_KINDS), tuple(WORKDAY_SHIFT_KINDS), AI_SHIFT_RULE_VERSION,
    )
    return _hash_cache_parts(
        shift_snapshot["base_hash"], shift_snapshot["previous_context_hash"],
        _hash_snapshot_rows(limit_rows), status_key, staffing_rules,
    )


def _king_of_time_report_signature(shift_base_hash, kot_detail_hash):
    code_rows = tuple(sorted((str(key), str(value or "")) for key, value in get_staff_code_map(active_only=False).items()))
    staff_key_rows = tuple(sorted((str(key), int(value)) for key, value in get_staff_key_map(active_only=False).items()))
    time_rules = tuple((kind, default_shift_times(kind)) for kind in SHIFT_KINDS)
    return _hash_cache_parts(shift_base_hash, kot_detail_hash, code_rows, staff_key_rows, time_rules)


@st.cache_data(ttl=60, show_spinner=False)
def _cached_staff_shift_pdf(year, month, report_signature, cache_version):
    return make_staff_shift_pdf(int(year), int(month))


@st.cache_data(ttl=60, show_spinner=False)
def _cached_staff_shift_excel(year, month, report_signature, cache_version):
    return make_staff_shift_excel(int(year), int(month))


@st.cache_data(ttl=60, show_spinner=False)
def _cached_shift_calendar_pdf(
    year, month, shift_base_hash, cache_version, _shift_rows, staff_list, selected_staff_names, finalized
):
    shift_df = pd.DataFrame(_shift_rows, columns=["shift_date", "staff_name", "shift_kind"])
    return make_shift_calendar_pdf(
        int(year),
        int(month),
        shift_df,
        list(staff_list),
        selected_staff_names=list(selected_staff_names) or None,
        finalized=bool(finalized),
    )


@st.cache_data(ttl=60, show_spinner=False)
def _cached_king_of_time_clock_export(year, month, selected_staff_keys, report_signature, cache_version):
    return build_king_of_time_clock_export(
        int(year), int(month),
        selected_staff_keys=None if selected_staff_keys is None else list(selected_staff_keys),
    )




def format_shift_label(shift_value):
    mapping = {
        "日勤": "日",
        "管": "管",
        "管理業務": "管",
        "夜勤": "夜",
        "夜勤明け": "明",
        "休み": "休",
        "希望休": "希",
        "有休": "有",
        "その他": "他",
    }
    return mapping.get(str(shift_value or ""), str(shift_value or ""))


def _weekday_column_label(year, month, day):
    d = date(int(year), int(month), int(day))
    return f"{day}\n({get_weekday_label(d)})"


def _status_column(df):
    if df is None or df.empty:
        return None
    for col in df.columns:
        if "状" in str(col) or str(col) in ("状態", "���"):
            return col
    return df.columns[-2] if len(df.columns) >= 2 else None


def render_shift_calendar(year, month, shift_df, read_index=None, staff_filter="全職員"):
    st.markdown("### 月間シフトカレンダー")
    if read_index is None:
        read_index = build_shift_read_index(shift_df)
    first_weekday, last_day = calendar.monthrange(int(year), int(month))
    start_col = (first_weekday + 1) % 7
    days = [None] * start_col
    days.extend(date(int(year), int(month), d) for d in range(1, last_day + 1))
    while len(days) % 7 != 0:
        days.append(None)

    header_cols = st.columns(7)
    for idx, label in enumerate(["日", "月", "火", "水", "木", "金", "土"]):
        cls = "hm-cal-head hm-sunday-text" if idx == 0 else "hm-cal-head hm-saturday-text" if idx == 6 else "hm-cal-head"
        header_cols[idx].markdown(f'<div class="{cls}">{label}</div>', unsafe_allow_html=True)

    for week_start in range(0, len(days), 7):
        cols = st.columns(7)
        for idx, target_date in enumerate(days[week_start:week_start + 7]):
            with cols[idx]:
                if target_date is None:
                    st.markdown('<div class="hm-shift-day hm-blank"></div>', unsafe_allow_html=True)
                    continue
                date_text = target_date.strftime("%Y-%m-%d")
                date_entries = read_index["entries_by_date_kind"].get(date_text, {})
                classes = ["hm-shift-day"]
                if target_date.weekday() == 6:
                    classes.append("hm-sunday")
                if target_date.weekday() == 5:
                    classes.append("hm-saturday")
                if target_date == today_jst():
                    classes.append("hm-today")
                lines = []
                if date_entries:
                    for kind in SHIFT_KINDS:
                        if kind == "管":
                            entries = date_entries.get(MANAGEMENT_INDEX_KIND, ())
                        else:
                            entries = date_entries.get(kind, ())
                        members = _unique_index_staff_names(entries, staff_filter)
                        if members:
                            names = "、".join(members[:3])
                            if len(members) > 3:
                                names += f" 他{len(members) - 3}"
                            lines.append(f'<div class="hm-shift-line"><b>{format_shift_label(kind)}</b> {html_escape(names)}</div>')
                html = (
                    f'<div class="{" ".join(classes)}">'
                    f'<div class="hm-day-top"><span class="hm-day-number">{target_date.day}</span>'
                    f'<span class="hm-weekday">{get_weekday_label(target_date)}</span></div>'
                    f'{"".join(lines) if lines else "<div class=\"hm-shift-empty\">未入力</div>"}'
                    f'</div>'
                )
                st.markdown(html, unsafe_allow_html=True)
                if st.button("編集", key=f"edit_shift_day_{target_date.isoformat()}", use_container_width=True):
                    st.session_state["selected_shift_date"] = target_date.strftime("%Y-%m-%d")


def render_selected_shift_day_editor(year, month, shift_df, month_status=None, read_index=None):
    st.markdown("### 選択日のシフト編集")
    st.caption("カレンダーの日付ごとの編集ボタンから、その日の職員シフトをまとめて修正できます。")
    st.caption("管は管理業務です。日勤人数には含めませんが、本人の勤務日数には含まれます。")

    selected_text = st.session_state.get("selected_shift_date")
    if not selected_text:
        st.info("月間シフトカレンダーの日付カードにある「編集」ボタンを押すと、この場所に日別編集フォームが表示されます。")
        return

    try:
        selected_date = datetime.strptime(str(selected_text), "%Y-%m-%d").date()
    except Exception:
        st.session_state.pop("selected_shift_date", None)
        st.warning("選択日付を読み取れませんでした。もう一度カレンダーから日付を選んでください。")
        return

    if selected_date.year != int(year) or selected_date.month != int(month):
        st.info("選択中の日付は表示中の月ではありません。カレンダーから編集したい日付を選んでください。")
        return

    staff_names = get_active_staff()
    if not staff_names:
        st.info("職員マスタに職員が登録されていません。先に職員マスタで登録してください。")
        return

    date_label = f"{selected_date.year}年{selected_date.month}月{selected_date.day}日（{get_weekday_label(selected_date)}）"
    st.markdown(f"#### {date_label} のシフト編集")

    message = st.session_state.pop("last_shift_day_save_message", None)
    if message:
        st.success(message)

    if month_status and month_status.get("status_error"):
        st.error("確定状態を確認できないため、日別編集を停止しています。")
    elif month_status and month_status.get("is_confirmed"):
        st.warning("この月は確定済みです。日別編集を保存する場合は、先に確定を解除してください。")

    with st.form(f"shift_day_editor_form_{selected_date.strftime('%Y%m%d')}"):
        assignments = {}
        cols = st.columns(3)
        for idx, staff_name in enumerate(staff_names):
            current_label = shift_day_actual_label_for_staff(
                shift_df, staff_name, selected_date.strftime("%Y-%m-%d"), read_index
            )
            current_index = SHIFT_EDITOR_OPTIONS.index(current_label) if current_label in SHIFT_EDITOR_OPTIONS else 0
            with cols[idx % 3]:
                assignments[staff_name] = st.selectbox(
                    staff_name,
                    SHIFT_EDITOR_OPTIONS,
                    index=current_index,
                    key=f"shift_day_select_{selected_date.strftime('%Y%m%d')}_{idx}_{staff_name}",
                    disabled=bool(month_status and shift_month_is_read_only(month_status)),
                )
        submitted = st.form_submit_button(
            "この日のシフトを保存",
            type="primary",
            use_container_width=True,
            disabled=bool(month_status and shift_month_is_read_only(month_status)),
        )

    if not submitted:
        return

    try:
        result = save_day_shift_assignments(selected_date.strftime("%Y-%m-%d"), assignments)
        st.session_state["shift_editor_reset_counter"] = int(st.session_state.get("shift_editor_reset_counter", 0) or 0) + 1
        st.session_state["last_shift_day_save_message"] = result.message
        st.rerun()
    except ShiftUpdateBlockedError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"この日のシフトを保存できませんでした：{e}")


def render_shift_editor(year, month, shift_df, staff_filter="全職員", month_status=None, read_index=None):
    st.markdown("### 月間シフト表")
    staff_names = get_active_staff()
    if staff_filter and staff_filter != "全職員":
        staff_names = [staff_filter]
    editable_matrix = create_editable_shift_matrix(
        staff_names, shift_df, int(year), int(month), read_index
    )
    if editable_matrix.empty:
        st.info("職員マスタに職員が登録されていません。先に職員マスタで登録してください。")
        return

    last_day = calendar.monthrange(int(year), int(month))[1]
    first_col = editable_matrix.columns[0]
    column_config = {
        first_col: st.column_config.TextColumn("職員名", disabled=True, width="medium"),
    }
    for d in range(1, last_day + 1):
        column_config[str(d)] = st.column_config.SelectboxColumn(
            _weekday_column_label(year, month, d),
            options=SHIFT_EDITOR_OPTIONS,
            required=False,
            width="small",
            help="日・管・夜・明・休・希・有・他を選択。空欄にすると削除扱いです。",
        )

    st.caption("各セルをクリックして、日・管・夜・明・休・希・有・他を直接入力できます。列見出しには曜日を表示しています。")
    edited_matrix = st.data_editor(
        editable_matrix,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config=column_config,
        key=f"shift_matrix_editor_{int(year)}_{int(month)}_{staff_filter}_{st.session_state.get('shift_editor_reset_counter', 0)}",
        disabled=bool(month_status and shift_month_is_read_only(month_status)),
    )

    if month_status and shift_month_is_read_only(month_status):
        if month_status.get("status_error"):
            st.error("確定状態を確認できないため、月間シフト表の保存を停止しています。")
        else:
            st.warning("この月は確定済みです。修正する場合は、先に確定を解除してください。")
        return
    if st.button("月間シフト表の入力内容を保存", use_container_width=True, type="primary"):
        try:
            result = save_shift_matrix_from_editor(int(year), int(month), edited_matrix)
            if result.status == SHIFT_SAVE_SAVED:
                queue_shift_save_result(result)
                st.rerun()
            show_shift_save_result(result)
        except ShiftUpdateBlockedError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"月間シフト表を保存できませんでした：{e}")


def page_shift_manager():
    st.subheader("シフト管理・AI割当")
    today = today_jst()
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        shift_year = st.number_input("対象年", min_value=2020, max_value=2100, value=today.year, step=1, key="shift_year")
    with c2:
        shift_month = st.number_input("対象月", min_value=1, max_value=12, value=today.month, step=1, key="shift_month")
    with c3:
        staff_filter = st.selectbox("表示する職員", ["全職員"] + get_active_staff(), key=f"shift_staff_filter_{int(shift_year)}_{int(shift_month)}")

    month_status = get_shift_month_status(int(shift_year), int(shift_month))
    month_read_only = shift_month_is_read_only(month_status)
    if month_status.get("status_error"):
        st.error("この月の確定状態を確認できないため、勤務データの更新と確定操作を停止しています。")
    elif month_status.get("is_confirmed"):
        st.success(f"この月のシフトは確定済みです。確定日時：{month_status.get('confirmed_at', '')}")
    else:
        st.info("この月のシフトは作成中です。")
    render_queued_shift_save_result()

    staff_options = [""] + get_active_staff()
    shift_df = get_staff_shifts_month(int(shift_year), int(shift_month), include_prev_day=True)
    shift_read_index = build_shift_read_index(
        shift_df, int(shift_year), int(shift_month)
    )

    render_shift_calendar(
        int(shift_year), int(shift_month), shift_df, shift_read_index, staff_filter
    )
    render_selected_shift_day_editor(
        int(shift_year), int(shift_month), shift_df, month_status, shift_read_index
    )

    st.markdown("### 入力・編集フォーム")
    with st.expander("希望休・有休・休みを登録", expanded=True):
        with st.form("hope_shift_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                hope_date = st.date_input("日付", value=today, key="hope_date", disabled=month_read_only)
            with c2:
                hope_staff = st.selectbox("職員", staff_options, key="hope_staff", disabled=month_read_only)
            with c3:
                hope_kind = st.selectbox("区分", ["希望休", "有休", "休み"], key="hope_kind", disabled=month_read_only)
            hope_memo = st.text_input("メモ", placeholder="本人希望、通院、家庭都合など", disabled=month_read_only)
            submit_hope = st.form_submit_button("保存", disabled=month_read_only)
        if submit_hope:
            if not hope_staff:
                st.error("職員を選択してください。")
            else:
                try:
                    result = save_single_shift(hope_date.strftime("%Y-%m-%d"), hope_staff, hope_kind, None, None, 0, hope_memo)
                    if result.status == SHIFT_SAVE_SAVED:
                        queue_shift_save_result(result)
                        st.rerun()
                    show_shift_save_result(result)
                except ShiftUpdateBlockedError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"休暇・休みを保存できませんでした：{e}")

    with st.expander("1日分の基本シフト入力"):
        with st.form("basic_shift_form", clear_on_submit=True):
            shift_date = st.date_input("シフト日", value=today, disabled=month_read_only)
            c1, c2, c3 = st.columns(3)
            with c1:
                day_staff_1 = st.selectbox("日勤1", staff_options, key="day_staff_1", disabled=month_read_only)
            with c2:
                day_staff_2 = st.selectbox("日勤2", staff_options, key="day_staff_2", disabled=month_read_only)
            with c3:
                night_staff = st.selectbox("夜勤", staff_options, key="night_staff", disabled=month_read_only)
            shift_memo = st.text_input("シフトメモ", disabled=month_read_only)
            submit_basic = st.form_submit_button("この日の基本シフトを保存", disabled=month_read_only)
        if submit_basic:
            if not day_staff_1 and not day_staff_2 and not night_staff:
                st.error("少なくとも1名を選択してください。")
            else:
                try:
                    result = save_basic_day_shift(shift_date.strftime("%Y-%m-%d"), day_staff_1, day_staff_2, night_staff, shift_memo)
                    if result.status == SHIFT_SAVE_SAVED:
                        queue_shift_save_result(result)
                        st.rerun()
                    show_shift_save_result(result)
                except ShiftUpdateBlockedError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"基本シフトを保存できませんでした：{e}")

    with st.expander("個別シフトを追加・調整する"):
        with st.form("single_shift_form", clear_on_submit=True):
            s_date = st.date_input("日付", value=today, key="single_shift_date", disabled=month_read_only)
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                s_staff = st.selectbox("職員", staff_options, key="single_shift_staff", disabled=month_read_only)
            with c2:
                s_kind = st.selectbox("勤務区分", SHIFT_KINDS, key="single_shift_kind", disabled=month_read_only)
            default_start, default_end, default_next = default_shift_times("日勤")
            with c3:
                s_start = st.text_input("開始", value=default_start, key="single_shift_start", disabled=month_read_only)
            with c4:
                s_end = st.text_input("終了", value=default_end, key="single_shift_end", disabled=month_read_only)
            s_next = st.checkbox("終了は翌日", value=False, disabled=month_read_only)
            s_memo = st.text_input("メモ", key="single_shift_memo", disabled=month_read_only)
            add_single = st.form_submit_button("個別シフトを追加", disabled=month_read_only)
        if add_single:
            if not s_staff:
                st.error("職員を選択してください。")
            else:
                try:
                    result = save_single_shift(s_date.strftime("%Y-%m-%d"), s_staff, s_kind, s_start, s_end, 1 if s_next else 0, s_memo)
                    if result.status == SHIFT_SAVE_SAVED:
                        queue_shift_save_result(result)
                        st.rerun()
                    show_shift_save_result(result)
                except ShiftUpdateBlockedError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"個別シフトを保存できませんでした：{e}")

    with st.expander("月間シフトを全部クリアして再入力する"):
        st.warning("表示中の月のシフト入力内容をすべて削除します。職員別勤務回数上限は残ります。")
        clear_confirm = st.checkbox(
            f"{int(shift_year)}年{int(shift_month)}月のシフトを全クリアすることを確認しました",
            key=f"clear_month_shift_confirm_{int(shift_year)}_{int(shift_month)}",
            disabled=month_read_only,
        )
        if st.button("この月のシフトを全クリアする", use_container_width=True, disabled=not clear_confirm or month_read_only):
            try:
                result = clear_month_staff_shifts(int(shift_year), int(shift_month))
                queue_shift_save_result(result)
                st.rerun()
            except ShiftUpdateBlockedError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"この月のシフトを全クリアできませんでした：{e}")

    render_shift_editor(
        int(shift_year), int(shift_month), shift_df, staff_filter, month_status,
        shift_read_index,
    )

    st.markdown("### 職員別勤務回数上限")
    limits_df = get_staff_shift_limits()
    if limits_df.empty:
        st.info("職員マスタに職員が登録されていません。")
    else:
        edited_limits = st.data_editor(
            limits_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key=f"shift_limit_editor_{int(shift_year)}_{int(shift_month)}",
        )
        if st.button("職員別勤務回数上限を保存", use_container_width=True):
            saved = save_staff_shift_limits_from_editor(edited_limits)
            clear_shift_caches()
            st.success(f"職員別勤務回数上限を保存しました。登録件数：{saved}件")
            st.rerun()

    st.markdown("### AIシフト案作成")
    with st.expander("AIが参照する職員別上限を確認"):
        st.dataframe(debug_shift_limit_summary_for_ai(), use_container_width=True, hide_index=True)
    if st.session_state.get("ai_shift_rule_version") != AI_SHIFT_RULE_VERSION:
        st.session_state.pop("ai_shift_draft", None)
        st.session_state["ai_shift_rule_version"] = AI_SHIFT_RULE_VERSION
    if st.button("不足分のAIシフト案を作成", use_container_width=True):
        st.session_state.pop("ai_shift_draft", None)
        st.session_state["ai_shift_rule_version"] = AI_SHIFT_RULE_VERSION
        st.session_state["ai_shift_draft"] = create_ai_shift_draft(
            shift_df,
            get_active_staff(),
            int(shift_year),
            int(shift_month),
        )
    draft = st.session_state.get("ai_shift_draft")
    if isinstance(draft, pd.DataFrame) and not draft.empty:
        draft = sanitize_ai_shift_draft_by_limits(draft, shift_df, int(shift_year), int(shift_month))
        st.session_state["ai_shift_draft"] = draft
        st.dataframe(draft, use_container_width=True, hide_index=True)
        preview_shift_df = apply_ai_shift_draft_to_df(shift_df, draft)
        st.markdown("#### AI案を反映した仮シフト表")
        st.dataframe(create_shift_matrix(preview_shift_df, int(shift_year), int(shift_month)), use_container_width=True, hide_index=True)
        c_save, c_clear = st.columns(2)
        with c_save:
            if st.button("確認したAIシフト案を保存する", use_container_width=True, disabled=month_read_only):
                try:
                    result = save_ai_shift_draft_rows(draft)
                    if result.status == SHIFT_SAVE_SAVED:
                        st.session_state.pop("ai_shift_draft", None)
                        queue_shift_save_result(result)
                        st.rerun()
                    show_shift_save_result(result)
                except ShiftUpdateBlockedError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"AIシフト案を保存できませんでした：{e}")
        with c_clear:
            if st.button("AIシフト案を破棄する", use_container_width=True):
                st.session_state.pop("ai_shift_draft", None)
                st.warning("AIシフト案を破棄しました。")
                st.rerun()

    st.markdown("### チェック結果・警告")
    matrix = create_shift_matrix(
        shift_df, int(shift_year), int(shift_month), shift_read_index
    )
    shortage = create_shift_shortage_table(
        shift_df, int(shift_year), int(shift_month), shift_read_index
    )
    checks = create_shift_quality_check_table(
        shift_df, int(shift_year), int(shift_month), shift_read_index, matrix
    )
    limit_checks = create_shift_limit_check_table(matrix)
    month_start = f"{int(shift_year)}-{int(shift_month):02d}-01"
    month_end = f"{int(shift_year)}-{int(shift_month):02d}-{calendar.monthrange(int(shift_year), int(shift_month))[1]:02d}"
    raw_month_shifts = fetch_staff_shifts_raw(month_start, month_end)
    duplicate_checks = create_shift_duplicate_check_table(
        raw_month_shifts, int(shift_year), int(shift_month)
    )
    status_col = _status_column(shortage)
    shortage_ng = shortage[shortage[status_col].astype(str) != "OK"] if status_col and not shortage.empty else pd.DataFrame()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("人員不足日", len(shortage_ng))
    with c2:
        st.metric("重複・連勤警告", 0 if checks is None or checks.empty else len(checks))
    with c3:
        st.metric("上限超過", 0 if limit_checks.empty else len(limit_checks))
    with c4:
        st.metric("状態", "確定" if month_status.get("is_confirmed") else "作成中")
    if not matrix.empty:
        st.dataframe(matrix, use_container_width=True, hide_index=True)
    if shortage_ng.empty:
        st.success("日勤・夜勤の必要人数チェックはOKです。")
    else:
        st.warning("人員不足の日があります。")
        st.dataframe(shortage_ng, use_container_width=True, hide_index=True)
    if checks is None or checks.empty:
        st.success("夜勤明け、希望休、有休、休みとの大きな重複警告はありません。")
    else:
        st.warning("シフト確認事項があります。")
        st.dataframe(checks, use_container_width=True, hide_index=True)
    if limit_checks.empty:
        st.success("職員別の勤務回数上限内です。")
    else:
        st.warning("職員別の勤務回数上限を超えている箇所があります。")
        st.dataframe(limit_checks, use_container_width=True, hide_index=True)
    if not duplicate_checks.empty:
        st.error("同一職員・同一日に複数の勤務登録があります。")
        st.dataframe(duplicate_checks, use_container_width=True, hide_index=True)

    st.markdown("### PDF/Excel/KING OF TIME CSV出力")
    st.caption("月間シフトカレンダーPDFは、保存済みの現在データから作成します。未保存の表編集はPDFに反映されません。先に保存してください。")
    # staff_filterは画面カレンダー専用。帳票は従来どおり保存済み全職員データから作る。
    shift_snapshot = shift_read_index["report_snapshot"]
    report_signature = _shift_report_signature(shift_snapshot, month_status, limits_df)

    if REPORTLAB_AVAILABLE:
        st.markdown("#### 月間シフトカレンダーPDF出力")
        pdf_target_mode = st.radio(
            "出力対象職員",
            ["全員", "職員を選択"],
            horizontal=True,
            key=f"shift_calendar_pdf_target_mode_{int(shift_year)}_{int(shift_month)}",
        )
        selected_pdf_staff = []
        if pdf_target_mode == "職員を選択":
            selected_pdf_staff = st.multiselect(
                "PDFに出力する職員",
                get_active_staff(),
                key=f"shift_calendar_pdf_staff_{int(shift_year)}_{int(shift_month)}",
            )
            if not selected_pdf_staff:
                st.info("職員を1名以上選択すると、選択した職員だけのシフトカレンダーPDFを作成できます。")

        if pdf_target_mode == "全員" or selected_pdf_staff:
            try:
                calendar_pdf_bytes = _cached_shift_calendar_pdf(
                    int(shift_year),
                    int(shift_month),
                    shift_snapshot["base_hash"],
                    SHIFT_CALENDAR_PDF_CACHE_VERSION,
                    _shift_rows=shift_snapshot["month_rows"],
                    staff_list=tuple(get_active_staff()),
                    selected_staff_names=tuple(selected_pdf_staff if pdf_target_mode == "職員を選択" else []),
                    finalized=bool(month_status.get("is_confirmed")),
                )
                if pdf_target_mode == "全員":
                    pdf_suffix = "all"
                else:
                    pdf_suffix = "_".join(re.sub(r"[^0-9A-Za-z_-]+", "", name) or f"staff{idx + 1}" for idx, name in enumerate(selected_pdf_staff))
                st.download_button(
                    "月間シフトカレンダーPDFをダウンロード",
                    data=calendar_pdf_bytes,
                    file_name=f"hidamari_shift_calendar_{int(shift_year)}_{int(shift_month):02d}_{pdf_suffix}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"月間シフトカレンダーPDFを作成できませんでした：{e}")
    else:
        st.warning("reportlab が未導入のため、月間シフトカレンダーPDFは出力できません。")

    kot_staff_options = get_king_of_time_export_staff(int(shift_year), int(shift_month))
    kot_staff_labels = {staff_id: staff_name for staff_id, staff_name in kot_staff_options}
    kot_available_staff_keys = list(kot_staff_labels)
    kot_selection_key = f"king_of_time_selected_staff_{int(shift_year)}_{int(shift_month)}"
    kot_options_key = f"{kot_selection_key}_options"
    previous_options = set(st.session_state.get(kot_options_key, []))
    available_options = set(kot_available_staff_keys)
    if kot_selection_key not in st.session_state:
        st.session_state[kot_selection_key] = kot_available_staff_keys.copy()
    elif previous_options != available_options:
        current_selection = set(st.session_state.get(kot_selection_key, []))
        newly_available = available_options - previous_options
        st.session_state[kot_selection_key] = [
            staff_id
            for staff_id in kot_available_staff_keys
            if staff_id in current_selection or staff_id in newly_available
        ]
    st.session_state[kot_options_key] = kot_available_staff_keys.copy()

    selected_kot_staff_keys = st.multiselect(
        "KING OF TIME打刻CSVに出力する職員",
        options=kot_available_staff_keys,
        format_func=lambda staff_id: kot_staff_labels.get(staff_id, str(staff_id)),
        key=kot_selection_key,
        disabled=not bool(kot_available_staff_keys),
    )
    if month_status.get("is_confirmed") and not selected_kot_staff_keys:
        st.warning("職員が1名も選択されていません。CSVは生成・ダウンロードできません。")

    selected_kot_staff_names = [kot_staff_labels[key] for key in selected_kot_staff_keys]
    missing_codes = get_missing_staff_code_names(selected_kot_staff_names)
    kot_preview_df = pd.DataFrame()
    kot_error_df = pd.DataFrame()
    kot_csv_bytes = None
    kot_generation_error = None
    kot_report_signature = None
    try:
        kot_report_signature = _king_of_time_report_signature(
            shift_snapshot["base_hash"], shift_snapshot["kot_detail_hash"]
        )
        kot_preview_df, kot_error_df, kot_csv_bytes = _cached_king_of_time_clock_export(
            int(shift_year),
            int(shift_month),
            tuple(selected_kot_staff_keys),
            kot_report_signature,
            KING_OF_TIME_CACHE_VERSION,
        )
    except Exception as e:
        kot_generation_error = e

    try:
        if kot_report_signature is None:
            raise RuntimeError("KING OF TIME検査用の帳票署名を作成できませんでした。")
        _, confirmation_kot_error_df, _ = _cached_king_of_time_clock_export(
            int(shift_year), int(shift_month), None,
            kot_report_signature, KING_OF_TIME_CACHE_VERSION,
        )
        confirmation_kot_generation_error = None
    except Exception as e:
        confirmation_kot_error_df = pd.DataFrame()
        confirmation_kot_generation_error = e

    confirmation_errors = build_shift_confirmation_errors(
        shortage_ng,
        checks,
        limit_checks,
        duplicate_checks,
        confirmation_kot_error_df,
        confirmation_kot_generation_error,
    )
    if not confirmation_errors.empty:
        st.error("重大な問題があるため、この勤務表は確定できません。")
        st.dataframe(confirmation_errors, use_container_width=True, hide_index=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if not month_status.get("is_confirmed"):
            if st.button(
                "この月のシフトを確定する",
                use_container_width=True,
                disabled=bool(month_status.get("status_error") or not confirmation_errors.empty),
            ):
                try:
                    result = confirm_shift_month_if_valid(
                        int(shift_year), int(shift_month), confirmation_errors,
                        current_login_user_for_shift(),
                    )
                    if result.status == SHIFT_SAVE_SAVED:
                        queue_shift_save_result(result)
                        st.rerun()
                    show_shift_save_result(result)
                except ShiftUpdateBlockedError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"勤務表を確定できませんでした：{e}")
        else:
            if st.button("確定を解除する", use_container_width=True):
                try:
                    result = set_shift_month_status(int(shift_year), int(shift_month), False, current_login_user_for_shift())
                    queue_shift_save_result(result)
                    st.rerun()
                except Exception as e:
                    st.error(f"確定を解除できませんでした：{e}")
    with c2:
        if REPORTLAB_AVAILABLE:
            try:
                pdf_bytes = _cached_staff_shift_pdf(
                    int(shift_year), int(shift_month), report_signature, STAFF_SHIFT_PDF_CACHE_VERSION
                )
                st.download_button("勤務表PDF", data=pdf_bytes, file_name=f"hidamari_shift_{int(shift_year)}_{int(shift_month):02d}.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"シフトPDFを作成できませんでした：{e}")
    with c3:
        if OPENPYXL_AVAILABLE:
            try:
                excel_bytes = _cached_staff_shift_excel(
                    int(shift_year), int(shift_month), report_signature, STAFF_SHIFT_EXCEL_CACHE_VERSION
                )
                st.download_button(
                    "勤務表Excel",
                    data=excel_bytes,
                    file_name=f"hidamari_shift_{int(shift_year)}_{int(shift_month):02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"シフトExcelを作成できませんでした：{e}")
    with c4:
        if missing_codes:
            st.warning("KING OF TIME従業員コード未登録：" + "、".join(missing_codes[:8]) + ("..." if len(missing_codes) > 8 else ""))
        if kot_generation_error is None:
            if not month_status.get("is_confirmed"):
                st.warning("KING OF TIME打刻CSVは、月間シフト確定後にダウンロードできます。")
            elif not selected_kot_staff_keys:
                st.warning("出力する職員を選択してください。")
            elif kot_error_df is not None and not kot_error_df.empty:
                st.error("KING OF TIME打刻CSVは、エラーがあるためダウンロードできません。")
            else:
                st.download_button(
                    "KING OF TIME打刻CSV",
                    data=kot_csv_bytes,
                    file_name=f"king_of_time_clock_{int(shift_year)}_{int(shift_month):02d}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
        else:
            st.error(f"KING OF TIME用CSVを作成できませんでした：{kot_generation_error}")

    if kot_generation_error is None:
        st.markdown("#### KING OF TIME打刻CSVプレビュー")
        if kot_preview_df is None or kot_preview_df.empty:
            st.info("打刻CSVに出力する日勤・夜勤シフトはありません。")
        else:
            def highlight_kot_error(row):
                if str(row.get("エラー有無", "")) == "あり":
                    return ["background-color: #fde2e2; color: #8a1f1f"] * len(row)
                return [""] * len(row)

            st.dataframe(
                kot_preview_df.style.apply(highlight_kot_error, axis=1),
                use_container_width=True,
                hide_index=True,
            )
        if kot_error_df is not None and not kot_error_df.empty:
            st.markdown("#### KING OF TIME打刻CSVエラー一覧")
            st.dataframe(kot_error_df, use_container_width=True, hide_index=True)
    else:
        st.error(f"KING OF TIME打刻CSVプレビューを作成できませんでした：{kot_generation_error}")

    st.markdown("### 過去シフト検索・更新・削除")
    c1, c2, c3 = st.columns(3)
    with c1:
        search_start = st.date_input("検索開始", value=date(int(shift_year), int(shift_month), 1))
    with c2:
        search_end = st.date_input("検索終了", value=date(int(shift_year), int(shift_month), calendar.monthrange(int(shift_year), int(shift_month))[1]))
    with c3:
        shift_keyword = st.text_input("職員名・勤務区分・メモ検索")

    search_df = get_staff_shifts(search_start.strftime("%Y-%m-%d"), search_end.strftime("%Y-%m-%d"), shift_keyword)
    if search_df.empty:
        st.info("該当するシフトはありません。")
    else:
        st.dataframe(search_df[["id", "shift_date", "staff_name", "shift_kind", "start_time", "end_time", "next_day", "memo"]], use_container_width=True, hide_index=True)
        selected_shift_id = st.selectbox("更新・削除するシフトID", search_df["id"].tolist())
        target = search_df[search_df["id"] == selected_shift_id].iloc[0]
        target_date_for_status = pd.to_datetime(str(target["shift_date"]), errors="coerce")
        target_month_status = (
            get_shift_month_status(target_date_for_status.year, target_date_for_status.month)
            if not pd.isna(target_date_for_status)
            else {"is_confirmed": 0, "status_error": "勤務日を読み取れません。"}
        )
        target_read_only = shift_month_is_read_only(target_month_status)
        if target_month_status.get("status_error"):
            st.error("選択したシフトの確定状態を確認できないため、更新・削除を停止しています。")
        elif target_month_status.get("is_confirmed"):
            st.warning("選択したシフトの月は確定済みです。更新・削除する場合は先に確定を解除してください。")
        with st.form("shift_update_form"):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                u_date = st.date_input("日付", value=datetime.strptime(target["shift_date"], "%Y-%m-%d").date(), disabled=target_read_only)
            with c2:
                current_staff = target["staff_name"] if target["staff_name"] in staff_options else ""
                u_staff = st.selectbox("職員", staff_options, index=staff_options.index(current_staff) if current_staff in staff_options else 0, disabled=target_read_only)
            with c3:
                current_kind = shift_kind_for_selectbox(target["shift_kind"])
                u_kind = st.selectbox("勤務区分", SHIFT_KINDS, index=SHIFT_KINDS.index(current_kind) if current_kind in SHIFT_KINDS else 0, disabled=target_read_only)
            with c4:
                u_next = st.checkbox("終了は翌日", value=bool(target["next_day"]), disabled=target_read_only)
            u_start = st.text_input("開始時刻", value=format_time_for_display(target.get("start_time")), disabled=target_read_only)
            u_end = st.text_input("終了時刻", value=format_time_for_display(target.get("end_time")), disabled=target_read_only)
            u_memo = st.text_input("メモ", value=target["memo"] or "", disabled=target_read_only)
            c_update, c_delete = st.columns(2)
            with c_update:
                update_shift = st.form_submit_button("シフトを更新", disabled=target_read_only)
            with c_delete:
                delete_shift = st.form_submit_button("シフトを削除", disabled=target_read_only)
        if update_shift:
            if not u_staff:
                st.error("職員を選択してください。")
            else:
                try:
                    result = update_staff_shift(
                        int(selected_shift_id), u_date.strftime("%Y-%m-%d"), u_staff, u_kind,
                        u_start, u_end, 1 if u_next else 0, u_memo,
                    )
                    if result.status == SHIFT_SAVE_SAVED:
                        queue_shift_save_result(result)
                        st.rerun()
                    show_shift_save_result(result)
                except ShiftUpdateBlockedError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"シフトを更新できませんでした：{e}")
        if delete_shift:
            try:
                result = delete_staff_shift(int(selected_shift_id))
                if result.status == SHIFT_SAVE_SAVED:
                    queue_shift_save_result(result)
                    st.rerun()
                show_shift_save_result(result)
            except ShiftUpdateBlockedError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"シフトを削除できませんでした：{e}")

    st.markdown("### カレンダー予定へのAI担当割当")
    events = monthly_events(int(shift_year), int(shift_month))
    event_rows = []
    for items in events.values():
        for ev in items:
            event_rows.append(dict(ev))
    events_df = pd.DataFrame(event_rows)
    if events_df.empty:
        st.info("この月の予定はありません。")
        return

    event_counts_by_date = get_staff_event_counts_month(int(shift_year), int(shift_month))

    options = {}
    for _, ev in events_df.iterrows():
        current_staff = ev["staff_name"] if ev.get("staff_name") else "未担当"
        start_text = format_time_for_display(ev.get("start_time"))
        label = f"ID:{ev['id']}｜{ev['event_date']}｜{start_text}｜{ev['title']}｜担当:{current_staff}"
        options[label] = int(ev["id"])

    if options:
        selected_event_label = st.selectbox(
            "担当候補を見る予定",
            list(options.keys()),
            key=f"event_assignment_select_{int(shift_year)}_{int(shift_month)}",
        )
        selected_event_id = options[selected_event_label]
        event_row = events_df[events_df["id"] == selected_event_id].iloc[0]

        candidates = _get_shift_candidates_for_event(
            event_row,
            shift_df,
            event_counts_by_date.get(str(event_row["event_date"]), {}),
        )
        if candidates.empty:
            st.warning("勤務表から担当候補が見つかりません。")
        else:
            st.dataframe(candidates, use_container_width=True, hide_index=True)
            c1, c2 = st.columns(2)
            with c1:
                ai_staff = st.selectbox(
                    "AI候補から選ぶ",
                    candidates["職員名"].tolist(),
                    key=f"ai_assign_staff_{selected_event_id}",
                )
                if st.button("AI候補を担当に反映", use_container_width=True, key=f"apply_ai_staff_{selected_event_id}"):
                    execute("UPDATE events SET staff_name=?, updated_at=? WHERE id=?", (ai_staff, now_text(), int(selected_event_id)))
                    clear_event_caches()
                    st.success(f"予定ID:{selected_event_id} の担当を {ai_staff} さんにしました。")
                    st.rerun()
            with c2:
                manual_staff = st.selectbox(
                    "自分で担当を選ぶ",
                    staff_options,
                    key=f"manual_assign_staff_{selected_event_id}",
                )
                if st.button("自分で選んだ担当を反映", use_container_width=True, key=f"apply_manual_staff_{selected_event_id}"):
                    execute("UPDATE events SET staff_name=?, updated_at=? WHERE id=?", (manual_staff or None, now_text(), int(selected_event_id)))
                    clear_event_caches()
                    st.success("担当を更新しました。")
                    st.rerun()

    st.markdown("#### 未担当予定の一括AI割当")
    preview_df = _build_event_assignment_preview(events_df, shift_df, event_counts_by_date, only_unassigned=True)
    if preview_df.empty:
        st.info("未担当予定はありません。")
    else:
        st.dataframe(preview_df, use_container_width=True, hide_index=True)
        assignable = preview_df[preview_df["AI候補"].fillna("").astype(str) != ""]
        if not assignable.empty and st.button(
            "未担当予定へ第1候補を一括反映",
            use_container_width=True,
            key=f"apply_bulk_ai_staff_{int(shift_year)}_{int(shift_month)}",
        ):
            params = [(r["AI候補"], now_text(), int(r["予定ID"])) for _, r in assignable.iterrows()]
            updated = execute_many("UPDATE events SET staff_name=?, updated_at=? WHERE id=?", params)
            clear_event_caches()
            st.success(f"{updated}件の予定へ担当候補を反映しました。")
            st.rerun()



def current_login_user_for_shift():
    """シフト確定者表示用。ログイン機構がない場合も止めない。"""
    for key in ["username", "user_id", "login_user", "user"]:
        try:
            value = st.session_state.get(key)
            if value:
                return str(value)
        except Exception:
            pass
    return "system"


def get_kot_auto_schedule_patterns():
    df = fetch_df("""
        SELECT shift_kind, pattern_code, pattern_name, day_type_code,
               day_type_name, leave_name, is_active
        FROM kot_auto_schedule_patterns
        ORDER BY shift_kind
    """)
    result = {}
    if df is not None and not df.empty:
        for _, row in df.iterrows():
            result[str(row["shift_kind"])] = {
                "pattern_code": str(row.get("pattern_code") or ""),
                "pattern_name": str(row.get("pattern_name") or ""),
                "day_type_code": str(row.get("day_type_code") or "1"),
                "day_type_name": str(row.get("day_type_name") or "平日"),
                "leave_name": str(row.get("leave_name") or ""),
                "is_active": int(row.get("is_active") or 0),
            }
    return result


def get_kot_auto_schedule_settings():
    df = fetch_df("SELECT setting_key, setting_value FROM kot_auto_schedule_settings ORDER BY setting_key")
    if df is None or df.empty:
        return {}
    return {str(row["setting_key"]): str(row.get("setting_value") or "") for _, row in df.iterrows()}


def save_kot_auto_schedule_settings(pattern_values, setting_values):
    for shift_kind, values in pattern_values.items():
        execute("""
            UPDATE kot_auto_schedule_patterns
            SET pattern_code=?, pattern_name=?, day_type_code=?, day_type_name=?,
                leave_name=?, is_active=1, updated_at=?
            WHERE shift_kind=?
        """, (
            values["pattern_code"].strip(), values["pattern_name"].strip(),
            values["day_type_code"], values["day_type_name"],
            values.get("leave_name", "").strip(), now_text(), shift_kind,
        ))
    for key, value in setting_values.items():
        execute("""
            UPDATE kot_auto_schedule_settings SET setting_value=?, updated_at=?
            WHERE setting_key=?
        """, (str(value), now_text(), key))


def page_kot_auto_schedule():
    st.subheader("KING OF TIME 自動スケジュールCSV")
    st.caption("確定済みの月間シフトから、従業員別自動スケジュールデータCSVを作成します。")
    st.warning(
        "このCSVは対象月だけの日別シフトを登録するものではありません。"
        "KING OF TIMEの従業員別自動スケジュール（第1～第6週の固定ルール）を更新し、"
        "今後の自動割り当てに利用されます。"
    )

    today = today_jst()
    c1, c2 = st.columns(2)
    with c1:
        target_year = st.number_input("対象年", min_value=2000, max_value=2099, value=int(today.year), step=1)
    with c2:
        target_month = st.selectbox("対象月", list(range(1, 13)), index=int(today.month) - 1)

    shift_df = get_staff_shifts_month(int(target_year), int(target_month))
    month_staff_names = set()
    if shift_df is not None and not shift_df.empty:
        month_staff_names = {
            normalize_staff_name(value)
            for value in shift_df["staff_name"].dropna().astype(str).tolist()
        }
    staff_df = fetch_df("""
        SELECT id, staff_name, staff_code, is_active
        FROM staff ORDER BY is_active DESC, staff_name
    """)
    staff_records = []
    staff_labels = {}
    if staff_df is not None and not staff_df.empty:
        for _, row in staff_df.iterrows():
            record = {
                "id": int(row["id"]),
                "staff_name": normalize_staff_name(row.get("staff_name", "")),
                "staff_code": row.get("staff_code"),
            }
            if month_staff_names:
                if record["staff_name"] not in month_staff_names:
                    continue
            elif int(row.get("is_active") or 0) != 1:
                continue
            staff_records.append(record)
            staff_labels[record["id"]] = record["staff_name"]
    all_staff_ids = list(staff_labels)
    selected_staff_ids = st.multiselect(
        "対象職員（複数選択）", all_staff_ids, default=all_staff_ids,
        format_func=lambda staff_id: staff_labels.get(staff_id, str(staff_id)),
    )
    holiday_mode = st.radio(
        "祝日の曜日番号",
        ["祝日は「8：祝」として出力", "祝日も通常の曜日番号で出力"],
        horizontal=True,
    )

    patterns = get_kot_auto_schedule_patterns()
    settings = get_kot_auto_schedule_settings()
    with st.expander("勤務パターン・休日設定", expanded=False):
        st.caption("パターンコードはKING OF TIME側の設定と同じ3～10文字の半角英数字を入力してください。")
        with st.form("kot_auto_schedule_settings_form"):
            pattern_values = {}
            for shift_kind in ("日勤", "夜勤", "管"):
                current = patterns.get(shift_kind, {})
                st.markdown(f"**{shift_kind}**")
                p1, p2 = st.columns(2)
                with p1:
                    pattern_code = st.text_input(
                        f"{shift_kind} パターンコード", value=current.get("pattern_code", ""),
                        key=f"kot_auto_pattern_code_{shift_kind}",
                    )
                with p2:
                    pattern_name = st.text_input(
                        f"{shift_kind} パターン名", value=current.get("pattern_name", ""),
                        key=f"kot_auto_pattern_name_{shift_kind}",
                    )
                pattern_values[shift_kind] = {
                    "pattern_code": pattern_code, "pattern_name": pattern_name,
                    "day_type_code": "1", "day_type_name": "平日", "leave_name": "",
                }

            day_type_options = {"1": "平日", "2": "法定休日", "3": "法定外休日"}
            s1, s2 = st.columns(2)
            with s1:
                rest_code = st.selectbox(
                    "通常の休みの勤務日種別", list(day_type_options),
                    index=list(day_type_options).index(settings.get("rest_day_type_code", "3")),
                    format_func=lambda value: f"{value}：{day_type_options[value]}",
                )
                rest_leave = st.text_input("通常の休みの休暇区分名", value=settings.get("rest_leave_name", "公休"))
                paid_leave = st.text_input("有休の休暇区分名", value=settings.get("paid_leave_name", "有休"))
            with s2:
                paid_code = st.selectbox(
                    "有休日の勤務日種別", list(day_type_options),
                    index=list(day_type_options).index(settings.get("paid_day_type_code", "1")),
                    format_func=lambda value: f"{value}：{day_type_options[value]}",
                )
                statutory = st.selectbox(
                    "法定休日として扱う曜日", ["", "1", "2", "3", "4", "5", "6", "7"],
                    index=["", "1", "2", "3", "4", "5", "6", "7"].index(settings.get("statutory_weekday", "")),
                    format_func=lambda value: "指定なし" if not value else f"{value}：{['日','月','火','水','木','金','土'][int(value)-1]}",
                )
                holiday_code = st.selectbox(
                    "祝日の勤務日種別", list(day_type_options),
                    index=list(day_type_options).index(settings.get("holiday_day_type_code", "1")),
                    format_func=lambda value: f"{value}：{day_type_options[value]}",
                )
            save_settings = st.form_submit_button("設定を保存")
        if save_settings:
            setting_values = {
                "rest_day_type_code": rest_code, "rest_day_type_name": day_type_options[rest_code],
                "rest_leave_name": rest_leave, "paid_day_type_code": paid_code,
                "paid_day_type_name": day_type_options[paid_code], "paid_leave_name": paid_leave,
                "statutory_weekday": statutory, "holiday_day_type_code": holiday_code,
            }
            save_kot_auto_schedule_settings(pattern_values, setting_values)
            st.success("勤務パターン・休日設定を保存しました。")
            st.rerun()

    preview_df, error_df, warning_df, output_bytes = build_auto_schedule_export(
        int(target_year), int(target_month), shift_df, staff_records, selected_staff_ids,
        patterns, settings,
        is_confirmed=shift_month_is_confirmed(int(target_year), int(target_month)),
        holiday_as_eight=holiday_mode.startswith("祝日は"),
    )
    if not selected_staff_ids:
        st.warning("職員が1名も選択されていません。CSVは生成・ダウンロードできません。")
    st.markdown("### 出力内容のプレビュー")
    if preview_df.empty:
        st.info("出力予定データはありません。")
    else:
        st.dataframe(preview_df, use_container_width=True, hide_index=True)
    if not error_df.empty:
        st.markdown("### 重大エラー")
        st.error("重大エラーがあるためCSVをダウンロードできません。エラーのある職員・日付は出力対象から除外されます。")
        st.dataframe(error_df, use_container_width=True, hide_index=True)
    if not warning_df.empty:
        st.markdown("### 警告")
        st.warning("警告対象のシフトはCSVへ出力されません。")
        st.dataframe(warning_df, use_container_width=True, hide_index=True)
    confirmed_rule_import = st.checkbox(
        "対象月限定の登録ではなく、KING OF TIMEの自動スケジュールルールを更新するCSVであることを確認しました",
        value=False,
    )
    if output_bytes is not None and confirmed_rule_import:
        st.download_button(
            "CSVをダウンロード", data=output_bytes,
            file_name=f"king_of_time_auto_schedule_{int(target_year)}_{int(target_month):02d}.csv",
            mime="text/csv", use_container_width=True,
        )
    elif output_bytes is not None:
        st.info("上の確認欄をチェックするとCSVをダウンロードできます。")

def page_storage_check():
    st.subheader("保存状態チェック")
    st.caption("PostgreSQL・Supabase Storage・DB上のファイル紐づけを確認します。")

    # -----------------------------
    # 接続・設定状態
    # -----------------------------
    st.markdown("### 1. Storage設定の状態")

    storage_ok = storage_is_configured()
    bucket_ok, bucket_msg = check_storage_bucket_access()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("PostgreSQL", "接続中")
    with c2:
        st.metric("Storage設定", "OK" if storage_ok else "未設定")
    with c3:
        st.metric("Storageバケット", "OK" if bucket_ok else "要確認")
    with c4:
        st.metric("requests", "OK" if REQUESTS_AVAILABLE else "未導入")

    st.write(f"**SUPABASE_URL**：`{get_supabase_url() or '未設定'}`")
    st.write(f"**SUPABASE_STORAGE_BUCKET**：`{get_supabase_storage_bucket() or '未設定'}`")
    st.write(f"**SUPABASE_SERVICE_ROLE_KEY**：`{'設定あり' if bool(get_supabase_storage_key()) else '未設定'}`")

    if bucket_ok:
        st.success(bucket_msg)
    else:
        st.warning(bucket_msg)

    # -----------------------------
    # 件数サマリー
    # -----------------------------
    st.markdown("---")
    st.markdown("### 2. 保存件数")

    event_count = count_query("SELECT COUNT(*) FROM events")
    photo_count = count_query("SELECT COUNT(*) FROM event_photos")
    file_count = count_query("SELECT COUNT(*) FROM event_files")
    storage_photo_count = count_query("SELECT COUNT(*) FROM event_photos WHERE file_path LIKE ?", ("storage://%",))
    storage_file_count = count_query("SELECT COUNT(*) FROM event_files WHERE file_path LIKE ?", ("storage://%",))
    local_photo_count = photo_count - storage_photo_count
    local_file_count = file_count - storage_file_count

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("予定件数", event_count)
    with c2:
        st.metric("写真メモ件数", photo_count)
    with c3:
        st.metric("添付ファイル件数", file_count)
    with c4:
        st.metric("Storage保存パス件数", storage_photo_count + storage_file_count)

    c5, c6 = st.columns(2)
    with c5:
        st.info(f"写真メモ：Storage {storage_photo_count}件 / 旧ローカル {local_photo_count}件")
    with c6:
        st.info(f"添付ファイル：Storage {storage_file_count}件 / 旧ローカル {local_file_count}件")

    # -----------------------------
    # 最新登録10件
    # -----------------------------
    st.markdown("---")
    st.markdown("### 3. 最新登録10件")

    latest_df = fetch_df("""
        SELECT
            e.id,
            e.event_date,
            e.category,
            e.title,
            e.user_name,
            e.staff_name,
            e.created_at,
            (SELECT COUNT(*) FROM event_photos p WHERE p.event_id = e.id) AS photo_count,
            (SELECT COUNT(*) FROM event_files f WHERE f.event_id = e.id) AS file_count
        FROM events e
        ORDER BY e.id DESC
        LIMIT 10
    """)
    if latest_df.empty:
        st.info("予定はまだ登録されていません。")
    else:
        st.dataframe(latest_df, use_container_width=True, hide_index=True)

    # -----------------------------
    # DBにはあるが読めないファイルの確認
    # -----------------------------
    st.markdown("---")
    st.markdown("### 4. DBにはあるがStorage・ローカルから読めないファイル")

    check_limit = st.number_input("確認件数上限", min_value=1, max_value=500, value=100, step=10)
    st.caption("Storage確認は通信が発生するため、件数が多い場合は少し時間がかかります。")

    if st.button("ファイル紐づけを確認する", use_container_width=True):
        targets = []

        photos = fetch_df("""
            SELECT
                '写真メモ' AS 種別,
                p.id AS id,
                p.event_id,
                p.file_name,
                p.file_path,
                p.created_at,
                e.event_date,
                e.title
            FROM event_photos p
            LEFT JOIN events e ON p.event_id = e.id
            ORDER BY p.id DESC
            LIMIT ?
        """, (int(check_limit),))

        files = fetch_df("""
            SELECT
                '添付ファイル' AS 種別,
                f.id AS id,
                f.event_id,
                f.file_name,
                f.file_path,
                f.created_at,
                e.event_date,
                e.title
            FROM event_files f
            LEFT JOIN events e ON f.event_id = e.id
            ORDER BY f.id DESC
            LIMIT ?
        """, (int(check_limit),))

        if not photos.empty:
            targets.extend(photos.to_dict("records"))
        if not files.empty:
            targets.extend(files.to_dict("records"))

        if not targets:
            st.info("確認対象の写真・添付ファイルはありません。")
        else:
            results = []
            progress = st.progress(0)
            for i, item in enumerate(targets, start=1):
                ok, message = storage_object_exists(item.get("file_path"))
                results.append({
                    "状態": "OK" if ok else "NG",
                    "種別": item.get("種別"),
                    "ID": item.get("id"),
                    "予定ID": item.get("event_id"),
                    "予定日": item.get("event_date"),
                    "予定タイトル": item.get("title"),
                    "ファイル名": item.get("file_name"),
                    "file_path": item.get("file_path"),
                    "確認結果": message,
                    "登録日時": item.get("created_at"),
                })
                progress.progress(i / len(targets))

            result_df = pd.DataFrame(results)
            ng_df = result_df[result_df["状態"] == "NG"].copy()

            st.metric("確認件数", len(result_df))
            st.metric("読めないファイル件数", len(ng_df))

            if ng_df.empty:
                st.success("確認した範囲では、DBに紐づいたファイルはすべて読み取り可能でした。")
            else:
                st.warning("DBには登録されていますが、Storage・ローカルから読めないファイルがあります。")
                st.dataframe(ng_df, use_container_width=True, hide_index=True)

            with st.expander("確認結果の全件を見る"):
                st.dataframe(result_df, use_container_width=True, hide_index=True)

    # -----------------------------
    # Storageパス一覧
    # -----------------------------
    st.markdown("---")
    st.markdown("### 5. Storage保存パス一覧")

    path_df = fetch_df("""
        SELECT '写真メモ' AS 種別, id, event_id, file_name, file_path, created_at
        FROM event_photos
        WHERE file_path LIKE ?
        UNION ALL
        SELECT '添付ファイル' AS 種別, id, event_id, file_name, file_path, created_at
        FROM event_files
        WHERE file_path LIKE ?
        ORDER BY created_at DESC
        LIMIT 50
    """, ("storage://%", "storage://%"))
    if path_df.empty:
        st.info("Storage保存パスはまだ登録されていません。")
    else:
        st.dataframe(path_df, use_container_width=True, hide_index=True)




def page_export():
    st.subheader("Excel・PDF出力")

    st.markdown("### Excel出力")
    started = time.perf_counter()
    events = fetch_df(f"SELECT {EVENT_DETAIL_COLUMNS} FROM events ORDER BY event_date, start_time, id")
    photos = fetch_df("SELECT id, event_id, file_name, file_path, photo_memo, created_at FROM event_photos ORDER BY event_id, id")
    files = fetch_df("SELECT id, event_id, file_name, file_path, file_type, file_memo, created_at FROM event_files ORDER BY event_id, id")
    categories = fetch_df(f"SELECT {CATEGORY_COLUMNS} FROM categories ORDER BY sort_order, category_name")
    users = fetch_df(f"SELECT {USER_COLUMNS} FROM users ORDER BY user_name")
    staff = fetch_df(f"SELECT {STAFF_COLUMNS} FROM staff ORDER BY staff_name")
    note_perf("Excel出力データ取得", started)

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

    st.caption("予定・利用者マスタ・職員マスタをまとめてExcel出力します。")

    st.markdown("---")
    st.markdown("### PDFカレンダー出力")
    st.caption("A4横の月間カレンダーPDFと、予定詳細一覧を出力します。掲示・申し送り・印刷用に使えます。")

    today = today_jst()
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        pdf_year = st.number_input("PDF出力 年", min_value=2020, max_value=2100, value=today.year, step=1)
    with c2:
        pdf_month = st.number_input("PDF出力 月", min_value=1, max_value=12, value=today.month, step=1)
    with c3:
        include_detail = st.checkbox("予定詳細一覧も付ける", value=True)

    if not REPORTLAB_AVAILABLE:
        st.error("PDF出力には reportlab が必要です。requirements.txt に reportlab を追加してください。")
    else:
        try:
            pdf_bytes = make_calendar_pdf(int(pdf_year), int(pdf_month), include_detail=include_detail)
            st.download_button(
                label="PDFカレンダーをダウンロード",
                data=pdf_bytes,
                file_name=f"hidamari_calendar_{int(pdf_year)}_{int(pdf_month):02d}.pdf",
                mime="application/pdf",
            )
        except Exception as e:
            st.error(f"PDFを作成できませんでした：{e}")




def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📅", layout="wide")
    init_db_once()
    add_css()

    st.title("📅 ひだまり帳 Ver1.4.7 PostgreSQL版")
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
            "シフト管理・AI割当",
            "KING OF TIME 自動スケジュールCSV",
            "保存状態チェック",
            "予定カテゴリ設定",
            "利用者マスタ",
            "職員マスタ",
            "Excel・PDF出力",
        ],
    )

    page_started = time.perf_counter()
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
    elif menu == "シフト管理・AI割当":
        page_shift_manager()
    elif menu == "KING OF TIME 自動スケジュールCSV":
        page_kot_auto_schedule()
    elif menu == "保存状態チェック":
        page_storage_check()
    elif menu == "予定カテゴリ設定":
        page_category_master()
    elif menu == "利用者マスタ":
        page_master_users()
    elif menu == "職員マスタ":
        page_master_staff()
    elif menu == "Excel・PDF出力":
        page_export()

    note_perf(menu, page_started)
    show_perf_log()


if __name__ == "__main__":
    main()
