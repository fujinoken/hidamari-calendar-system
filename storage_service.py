# -*- coding: utf-8 -*-
"""Supabase Storage and attachment helpers for ひだまり帳."""
import hashlib
import mimetypes
import os
import urllib.parse
from datetime import datetime
from pathlib import Path

import streamlit as st

try:
    import requests
except ImportError:
    requests = None

from config import (
    DEFAULT_SUPABASE_STORAGE_BUCKET,
    STORAGE_PATH_PREFIX,
    SUPABASE_STORAGE_BUCKET_KEY,
    SUPABASE_STORAGE_KEY_KEYS,
    SUPABASE_URL_KEYS,
)
from db import JST, execute, fetch_df, now_text

REQUESTS_AVAILABLE = requests is not None


def get_secret_or_env(*keys, default=None):
    """Streamlit secrets または環境変数から設定値を取得する。"""
    for key in keys:
        try:
            if key in st.secrets:
                value = st.secrets[key]
                if value:
                    return str(value).strip()
        except Exception:
            pass
        value = os.getenv(key)
        if value:
            return str(value).strip()
    return default


def get_supabase_url():
    """
    Supabase Project URLを取得する。
    /rest/v1 や /storage/v1 まで入れてしまった場合も、プロジェクトURLへ補正する。
    """
    url = get_secret_or_env(*SUPABASE_URL_KEYS, default="")
    if not url:
        return ""
    url = str(url).strip().rstrip("/")
    for suffix in ["/rest/v1", "/storage/v1", "/storage/v1/object"]:
        if url.endswith(suffix):
            url = url[: -len(suffix)]
    return url.rstrip("/")


def get_supabase_storage_key():
    """
    Supabase Storage操作用キーを取得する。
    非公開バケットをサーバー側から扱うため、Streamlit secretsには SERVICE_ROLE_KEY を入れる。
    """
    return get_secret_or_env(*SUPABASE_STORAGE_KEY_KEYS, default="")


def get_supabase_storage_bucket():
    """Storageバケット名。未指定時は hidamari-calendar-files を使う。"""
    return get_secret_or_env(SUPABASE_STORAGE_BUCKET_KEY, default=DEFAULT_SUPABASE_STORAGE_BUCKET)


def storage_is_configured():
    """Supabase Storage保存に必要な設定が揃っているか確認する。"""
    return bool(get_supabase_url() and get_supabase_storage_key() and get_supabase_storage_bucket())


def require_storage_ready():
    """Storage未設定やrequests未導入の場合に分かりやすいエラーを出す。"""
    if requests is None:
        raise RuntimeError("requests がインストールされていません。requirements.txt に requests を追加してください。")
    if not get_supabase_url():
        raise RuntimeError("SUPABASE_URL が未設定です。Streamlit secrets に設定してください。")
    if not get_supabase_storage_key():
        raise RuntimeError("SUPABASE_SERVICE_ROLE_KEY が未設定です。Streamlit secrets に設定してください。")
    if not get_supabase_storage_bucket():
        raise RuntimeError("SUPABASE_STORAGE_BUCKET が未設定です。")


def guess_content_type(file_name):
    content_type, _ = mimetypes.guess_type(str(file_name or ""))
    return content_type or "application/octet-stream"


def make_storage_object_path(event_id, original_name, kind):
    """
    Storage内の保存パスを作る。
    元ファイル名はDBのfile_nameに保存し、Storage側は衝突しにくい安全な名前にする。
    """
    suffix = Path(str(original_name or "")).suffix.lower()
    if not suffix:
        suffix = ".bin"
    timestamp = datetime.now(JST).strftime("%Y%m%d%H%M%S%f")
    digest = hashlib.md5(f"{event_id}-{original_name}-{timestamp}".encode("utf-8")).hexdigest()[:12]
    safe_kind = "photos" if kind == "photos" else "files"
    return f"events/{int(event_id)}/{safe_kind}/{timestamp}_{digest}{suffix}"


def make_storage_db_path(bucket, object_path):
    """DBには storage://bucket/path の形式で保存する。"""
    return f"{STORAGE_PATH_PREFIX}{bucket}/{object_path}"


def parse_storage_db_path(file_path):
    """storage://bucket/path を (bucket, path) に分解する。"""
    value = str(file_path or "").strip()
    if not value.startswith(STORAGE_PATH_PREFIX):
        return None, None
    rest = value[len(STORAGE_PATH_PREFIX):]
    if "/" not in rest:
        return None, None
    bucket, object_path = rest.split("/", 1)
    return bucket, object_path


def is_storage_file(file_path):
    return str(file_path or "").strip().startswith(STORAGE_PATH_PREFIX)


def storage_url_for_object(bucket, object_path, authenticated=False):
    """
    Supabase Storage REST URLを作る。
    upload/delete は /object、非公開ファイルの取得は /object/authenticated を使う。
    """
    base_url = get_supabase_url()
    quoted_bucket = urllib.parse.quote(str(bucket), safe="")
    quoted_path = urllib.parse.quote(str(object_path), safe="/")
    prefix = "object/authenticated" if authenticated else "object"
    return f"{base_url}/storage/v1/{prefix}/{quoted_bucket}/{quoted_path}"


def storage_headers(content_type=None, upsert=False):
    key = get_supabase_storage_key()
    headers = {
        "Authorization": f"Bearer {key}",
        "apikey": key,
    }
    if content_type:
        headers["Content-Type"] = content_type
    if upsert:
        headers["x-upsert"] = "true"
    return headers


def upload_bytes_to_storage(file_bytes, object_path, content_type):
    """Supabase Storageへファイル本体をアップロードする。"""
    require_storage_ready()
    bucket = get_supabase_storage_bucket()
    url = storage_url_for_object(bucket, object_path)
    response = requests.post(
        url,
        headers=storage_headers(content_type=content_type, upsert=True),
        data=file_bytes,
        timeout=90,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"Supabase Storageへのアップロードに失敗しました: {response.status_code} {response.text[:300]}")
    return make_storage_db_path(bucket, object_path)


def download_bytes_from_storage(file_path):
    """storage:// のDBパスからStorage上のファイル本体を取得する。"""
    require_storage_ready()
    bucket, object_path = parse_storage_db_path(file_path)
    if not bucket or not object_path:
        raise RuntimeError("Storageパスの形式が不正です。")
    url = storage_url_for_object(bucket, object_path, authenticated=True)
    response = requests.get(url, headers=storage_headers(), timeout=90)
    if response.status_code >= 400:
        raise RuntimeError(f"Supabase Storageからの取得に失敗しました: {response.status_code} {response.text[:300]}")
    return response.content


def delete_storage_object(file_path):
    """
    Storage上のファイル削除。
    環境やAPI仕様差に備え、失敗しても画面操作を止めないためFalseを返す。
    """
    if not is_storage_file(file_path):
        return False
    try:
        require_storage_ready()
        bucket, object_path = parse_storage_db_path(file_path)
        if not bucket or not object_path:
            return False
        url = storage_url_for_object(bucket, object_path)
        response = requests.delete(url, headers=storage_headers(), timeout=60)
        return response.status_code < 400
    except Exception:
        return False


@st.cache_data(ttl=300, show_spinner=False)
def _read_saved_file_bytes_cached(file_path):
    """
    Storage保存・旧ローカル保存の両方に対応してファイル本体を読む。
    画像や添付の再表示を速くするため、短時間キャッシュする。
    """
    if is_storage_file(file_path):
        return download_bytes_from_storage(file_path)

    local_path = Path(str(file_path or ""))
    if local_path.exists():
        with open(local_path, "rb") as f:
            return f.read()
    return None


def read_saved_file_bytes(file_path):
    return _read_saved_file_bytes_cached(str(file_path or ""))


def render_saved_image(file_path, caption=None, use_container_width=True):
    """Storage保存・旧ローカル保存の両方に対応して画像を表示する。"""
    try:
        data = read_saved_file_bytes(file_path)
        if data:
            st.image(data, caption=caption, use_container_width=use_container_width)
            return True
    except Exception as e:
        st.warning(f"画像を取得できません：{e}")
        return False

    st.warning("画像ファイルが見つかりません。")
    return False


def render_saved_download_button(label, file_path, file_name, key):
    """Storage保存・旧ローカル保存の両方に対応してダウンロードボタンを表示する。"""
    try:
        data = read_saved_file_bytes(file_path)
        if data:
            st.download_button(
                label,
                data=data,
                file_name=file_name,
                key=key,
            )
            return True
    except Exception as e:
        st.warning(f"ファイルを取得できません：{e}")
        return False

    st.warning("ファイルが見つかりません。")
    return False


def delete_saved_file(file_path):
    """Storage保存・旧ローカル保存の両方に対応してファイル本体を削除する。"""
    if is_storage_file(file_path):
        return delete_storage_object(file_path)

    local_path = Path(str(file_path or ""))
    if local_path.exists():
        try:
            local_path.unlink()
            return True
        except Exception:
            return False
    return False


def storage_object_exists(file_path):
    """
    DBに保存されたfile_pathが実際に読めるか確認する。
    Storage保存はHEAD、失敗時はGETで確認。旧ローカル保存はファイル存在で確認。
    """
    value = str(file_path or "").strip()
    if not value:
        return False, "file_path空欄"

    if is_storage_file(value):
        try:
            require_storage_ready()
            bucket, object_path = parse_storage_db_path(value)
            if not bucket or not object_path:
                return False, "Storageパス形式不正"

            url = storage_url_for_object(bucket, object_path, authenticated=True)
            response = requests.head(url, headers=storage_headers(), timeout=30)
            if response.status_code < 400:
                return True, "OK"

            # 環境によってHEADが許可されない場合の保険
            response = requests.get(url, headers=storage_headers(), timeout=30)
            if response.status_code < 400:
                return True, "OK"

            return False, f"Storage取得不可 {response.status_code}: {response.text[:120]}"
        except Exception as e:
            return False, f"Storage確認エラー: {e}"

    local_path = Path(value)
    if local_path.exists():
        return True, "OK（旧ローカル）"
    return False, "旧ローカルファイルなし"


def check_storage_bucket_access():
    """Storageバケットにアクセスできるか簡易確認する。"""
    if requests is None:
        return False, "requests が未導入です。"
    if not storage_is_configured():
        return False, "Storage設定が未完了です。"

    try:
        bucket = get_supabase_storage_bucket()
        base_url = get_supabase_url()
        quoted_bucket = urllib.parse.quote(str(bucket), safe="")
        url = f"{base_url}/storage/v1/object/list/{quoted_bucket}"
        response = requests.post(
            url,
            headers=storage_headers(content_type="application/json"),
            json={"prefix": "", "limit": 1, "offset": 0},
            timeout=30,
        )
        if response.status_code < 400:
            return True, "Storageバケットへ接続できました。"
        return False, f"Storageバケット確認失敗 {response.status_code}: {response.text[:180]}"
    except Exception as e:
        return False, f"Storageバケット確認エラー: {e}"



def save_uploaded_photos(event_id, uploaded_files, photo_memo=""):
    """
    アップロード写真をSupabase Storageへ保存し、DBへ紐づける。
    戻り値: (保存成功数, 失敗数)
    """
    if not uploaded_files:
        return 0, 0

    saved_count = 0
    failed_count = 0

    if not storage_is_configured():
        st.error("Supabase Storage設定が未完了のため、写真メモは保存されませんでした。")
        return 0, len(uploaded_files)

    for uploaded in uploaded_files:
        try:
            object_path = make_storage_object_path(event_id, uploaded.name, kind="photos")
            file_bytes = bytes(uploaded.getbuffer())
            storage_path = upload_bytes_to_storage(
                file_bytes,
                object_path,
                guess_content_type(uploaded.name),
            )

            photo_id = execute("""
                INSERT INTO event_photos
                (event_id, file_name, file_path, photo_memo, created_at)
                VALUES (?, ?, ?, ?, ?)
            """, (
                int(event_id),
                uploaded.name,
                storage_path,
                photo_memo.strip() or None,
                now_text(),
            ))
            if photo_id:
                saved_count += 1
        except Exception as e:
            failed_count += 1
            st.error(f"写真メモのStorage保存に失敗しました：{uploaded.name}｜{e}")

    return saved_count, failed_count


@st.cache_data(ttl=20, show_spinner=False)
def get_event_photos(event_id):
    return fetch_df(
        "SELECT * FROM event_photos WHERE event_id=? ORDER BY id",
        (int(event_id),)
    )


def save_uploaded_files(event_id, uploaded_files, file_memo=""):
    """
    Excel等の添付ファイルをSupabase Storageへ保存し、DBへ紐づける。
    戻り値: (保存成功数, 失敗数)
    """
    if not uploaded_files:
        return 0, 0

    saved_count = 0
    failed_count = 0

    if not storage_is_configured():
        st.error("Supabase Storage設定が未完了のため、Excel・書類ファイルは保存されませんでした。")
        return 0, len(uploaded_files)

    for uploaded in uploaded_files:
        try:
            suffix = Path(uploaded.name).suffix.lower()
            object_path = make_storage_object_path(event_id, uploaded.name, kind="files")
            file_bytes = bytes(uploaded.getbuffer())
            storage_path = upload_bytes_to_storage(
                file_bytes,
                object_path,
                guess_content_type(uploaded.name),
            )

            file_id = execute("""
                INSERT INTO event_files
                (event_id, file_name, file_path, file_type, file_memo, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                int(event_id),
                uploaded.name,
                storage_path,
                suffix.replace(".", ""),
                file_memo.strip() or None,
                now_text(),
            ))
            if file_id:
                saved_count += 1
        except Exception as e:
            failed_count += 1
            st.error(f"Excel・書類ファイルのStorage保存に失敗しました：{uploaded.name}｜{e}")

    return saved_count, failed_count


@st.cache_data(ttl=20, show_spinner=False)
def get_event_files(event_id):
    return fetch_df(
        "SELECT * FROM event_files WHERE event_id=? ORDER BY id",
        (int(event_id),)
    )


