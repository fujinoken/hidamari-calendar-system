# -*- coding: utf-8 -*-
"""Application-wide constants for ひだまり帳."""
from pathlib import Path

APP_TITLE = "ひだまり帳 Ver1.4.7 PostgreSQL版"
AI_SHIFT_RULE_VERSION = "shift_overlap_strict_v6_staff_deduplicate_excel_export"

UPLOAD_DIR = Path("uploads")
FILE_DIR = Path("attached_files")

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

SHIFT_KINDS = ["日勤", "管理業務", "夜勤", "夜勤明け", "休み", "希望休", "有休", "その他"]
SHIFT_EDITOR_OPTIONS = ["", "日", "管", "夜", "明", "希", "有", "他"]

PDF_FONT_GOTHIC = "HeiseiKakuGo-W5"
PDF_FONT_MINCHO = "HeiseiMin-W3"

STORAGE_PATH_PREFIX = "storage://"
SUPABASE_URL_KEYS = ("SUPABASE_URL", "SUPABASE_PROJECT_URL")
SUPABASE_STORAGE_KEY_KEYS = (
    "SUPABASE_SERVICE_ROLE_KEY",
    "SUPABASE_SERVICE_KEY",
    "SUPABASE_STORAGE_KEY",
    "SUPABASE_KEY",
)
SUPABASE_STORAGE_BUCKET_KEYS = ("SUPABASE_STORAGE_BUCKET", "SUPABASE_BUCKET")
SUPABASE_STORAGE_BUCKET_KEY = SUPABASE_STORAGE_BUCKET_KEYS[0]
DEFAULT_SUPABASE_STORAGE_BUCKET = "hidamari-calendar-files"
KOT_DAY_PATTERN_CODE_KEY = "KOT_DAY_PATTERN_CODE"
KOT_NIGHT_PATTERN_CODE_KEY = "KOT_NIGHT_PATTERN_CODE"
KOT_REST_LEAVE_NAME_KEY = "KOT_REST_LEAVE_NAME"
KOT_PAID_LEAVE_NAME_KEY = "KOT_PAID_LEAVE_NAME"
KOT_HOPE_REST_LEAVE_NAME_KEY = "KOT_HOPE_REST_LEAVE_NAME"
