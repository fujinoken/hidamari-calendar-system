# -*- coding: utf-8 -*-
"""KING OF TIME 従業員別自動スケジュールCSV生成ロジック。"""

from __future__ import annotations

import csv
import io
import math
import re
import unicodedata
from datetime import date, timedelta

import pandas as pd


CSV_COLUMNS = ["勤務日", "従業員コード", "パターンコード"]
PREVIEW_COLUMNS = ["元の日付", "元のシフト区分", "職員名"] + CSV_COLUMNS
ISSUE_COLUMNS = ["種別", "職員名", "日付", "内容"]
WEEKDAY_LABELS = {1: "日", 2: "月", 3: "火", 4: "水", 5: "木", 6: "金", 7: "土", 8: "祝"}
WORK_SHIFT_KINDS = {"日勤", "夜勤", "管", "管理", "管理勤務", "管理業務"}
NIGHT_AFTER_KINDS = {"明", "明け", "夜勤明け"}
EXCLUDED_SHIFT_KINDS = {"休", "休み", "有休", "希望休"} | NIGHT_AFTER_KINDS


def _blank(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip().lower() in {"", "none", "nan"}


def normalize_employee_code(value):
    if _blank(value):
        return ""
    return unicodedata.normalize("NFKC", str(value)).strip()


def validate_employee_code(value):
    code = normalize_employee_code(value)
    if not code:
        return "", "KING OF TIME従業員コードが登録されていません"
    if not re.fullmatch(r"[A-Za-z0-9]{3,10}", code):
        return code, "従業員コードは3～10文字の半角英数字で設定してください"
    return code, ""


def validate_pattern(pattern):
    code = normalize_employee_code(pattern.get("pattern_code"))
    if not re.fullmatch(r"[A-Za-z0-9]{3,10}", code):
        return "KING OF TIMEパターンコードは3～10文字の半角英数字で設定してください"
    return ""


def calendar_week_number(target):
    """日曜始まりの月間カレンダー上の行を1～6で返す。"""
    first_sunday_index = (target.replace(day=1).weekday() + 1) % 7
    return ((first_sunday_index + target.day - 1) // 7) + 1


def weekday_number(target, holiday_as_eight=False, is_holiday=False):
    if holiday_as_eight and is_holiday:
        return 8
    return ((target.weekday() + 1) % 7) + 1


def _vernal_equinox_day(year):
    return int(20.8431 + 0.242194 * (year - 1980) - math.floor((year - 1980) / 4))


def _autumn_equinox_day(year):
    return int(23.2488 + 0.242194 * (year - 1980) - math.floor((year - 1980) / 4))


def japan_holiday_dates(year):
    """2000～2099年向けに日本の祝日・振替休日等を返す。"""
    year = int(year)
    if not 2000 <= year <= 2099:
        return set()

    def nth_weekday(month, weekday, nth):
        first = date(year, month, 1)
        return first + timedelta(days=(weekday - first.weekday()) % 7 + (nth - 1) * 7)

    holidays = {
        date(year, 1, 1), nth_weekday(1, 0, 2), date(year, 2, 11),
        date(year, 3, _vernal_equinox_day(year)), date(year, 4, 29),
        date(year, 5, 3), date(year, 5, 4), date(year, 5, 5),
        nth_weekday(7, 0, 3), date(year, 8, 11), nth_weekday(9, 0, 3),
        date(year, 9, _autumn_equinox_day(year)), nth_weekday(10, 0, 2),
        date(year, 11, 3), date(year, 11, 23),
    }
    if year >= 2020:
        holidays.add(date(year, 2, 23))
    if year == 2020:
        holidays -= {nth_weekday(7, 0, 3), date(year, 8, 11)}
        holidays |= {date(year, 7, 23), date(year, 7, 24), date(year, 8, 10)}
    if year == 2021:
        holidays -= {nth_weekday(7, 0, 3), nth_weekday(10, 0, 2), date(year, 8, 11)}
        holidays |= {date(year, 7, 22), date(year, 7, 23), date(year, 8, 8)}
    current = date(year, 1, 2)
    while current < date(year, 12, 31):
        if current not in holidays and current - timedelta(days=1) in holidays and current + timedelta(days=1) in holidays:
            holidays.add(current)
        current += timedelta(days=1)
    for holiday in sorted(list(holidays)):
        if holiday.weekday() == 6:
            substitute = holiday + timedelta(days=1)
            while substitute in holidays:
                substitute += timedelta(days=1)
            holidays.add(substitute)
    return holidays


def csv_bytes(rows, columns=CSV_COLUMNS, encoding="cp932"):
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=list(columns), extrasaction="ignore", lineterminator="\r\n")
    for row in rows:
        writer.writerow({column: "" if _blank(row.get(column)) else row.get(column) for column in columns})
    return output.getvalue().encode(encoding)


def _issue(kind, staff_name, shift_date, message):
    return {"種別": kind, "職員名": staff_name, "日付": shift_date, "内容": message}


def _result(preview_rows, errors, warnings):
    preview = pd.DataFrame(preview_rows, columns=PREVIEW_COLUMNS)
    for row in preview_rows:
        for column in CSV_COLUMNS:
            value = "" if _blank(row.get(column)) else str(row.get(column))
            try:
                value.encode("cp932")
            except UnicodeEncodeError:
                errors.append(_issue(
                    "重大エラー", row.get("職員名", ""), row.get("元の日付", ""),
                    f"{column}「{value}」にcp932で表現できない文字が含まれています",
                ))
    error_df = pd.DataFrame(errors, columns=ISSUE_COLUMNS)
    warning_df = pd.DataFrame(warnings, columns=ISSUE_COLUMNS)
    output = None
    if error_df.empty:
        try:
            output = csv_bytes(preview_rows)
        except UnicodeEncodeError as exc:
            errors.append(_issue("重大エラー", "", "", f"cp932で表現できない文字が含まれています：{exc}"))
            error_df = pd.DataFrame(errors, columns=ISSUE_COLUMNS)
    return preview, error_df, warning_df, output


def build_auto_schedule_export(
    year, month, shifts, staff_records, selected_staff_ids, patterns, settings,
    *, is_confirmed, holiday_as_eight=True, holiday_dates=None,
):
    """プレビュー、重大エラー、警告、CSVバイトを決定的な順序で生成する。"""
    year, month = int(year), int(month)
    selected_ids = {int(value) for value in (selected_staff_ids or [])}
    errors, warnings, preview_rows = [], [], []
    if not selected_ids:
        errors.append(_issue("重大エラー", "", "", "対象職員を1名以上選択してください"))

    staff_by_name = {}
    for staff in staff_records or []:
        name = str(staff.get("staff_name") or "").strip()
        if name and int(staff.get("id")) in selected_ids:
            staff_by_name[name] = staff

    frame = shifts.copy() if shifts is not None else pd.DataFrame()
    if not is_confirmed:
        errors.append(_issue("重大エラー", "", f"{year}-{month:02d}", "対象年月の確定シフトが存在しません"))
        if not frame.empty:
            for _, row in frame.iterrows():
                name = str(row.get("staff_name") or "").strip()
                if name in staff_by_name:
                    warnings.append(_issue("警告", name, str(row.get("shift_date") or ""), "未確定シフトのため出力しません"))
        return _result(preview_rows, errors, warnings)

    valid_staff = {}
    for name, staff in sorted(staff_by_name.items()):
        code, message = validate_employee_code(staff.get("staff_code"))
        if message:
            errors.append(_issue("重大エラー", name, "", message))
        else:
            valid_staff[name] = code

    if frame.empty:
        errors.append(_issue("重大エラー", "", f"{year}-{month:02d}", "対象年月の確定シフトが存在しません"))
        return _result(preview_rows, errors, warnings)
    for column in ("shift_date", "staff_name", "shift_kind"):
        if column not in frame:
            frame[column] = ""
    sort_columns = [column for column in ("shift_date", "staff_name", "shift_kind", "id") if column in frame]
    frame = frame.sort_values(sort_columns, kind="stable")
    for _, row in frame.iterrows():
        staff_name = str(row.get("staff_name") or "").strip()
        if staff_name not in valid_staff:
            continue
        raw_date = str(row.get("shift_date") or "").strip()
        shift_kind = str(row.get("shift_kind") or "").strip()
        try:
            target = pd.to_datetime(raw_date).date()
        except Exception:
            errors.append(_issue("重大エラー", staff_name, raw_date, "日付が不正です"))
            continue
        if target.year != year or target.month != month:
            continue
        if shift_kind in EXCLUDED_SHIFT_KINDS:
            warnings.append(_issue("警告", staff_name, raw_date, f"出力対象外の「{shift_kind}」です"))
            continue
        if not shift_kind:
            warnings.append(_issue("警告", staff_name, raw_date, "未確定シフトのため出力しません"))
            continue

        pattern_key = "管" if shift_kind in {"管理", "管理勤務", "管理業務"} else shift_kind
        pattern = patterns.get(pattern_key) or {}
        if shift_kind in WORK_SHIFT_KINDS:
            pattern_error = validate_pattern(pattern)
            if pattern_error:
                errors.append(_issue("重大エラー", staff_name, raw_date, f"{shift_kind}の{pattern_error}"))
                continue
            pattern_code = normalize_employee_code(pattern.get("pattern_code"))
        else:
            warnings.append(_issue("警告", staff_name, raw_date, f"勤務扱いか休日扱いか未設定の区分「{shift_kind}」です"))
            continue

        preview_rows.append({
            "元の日付": target.isoformat(), "元のシフト区分": shift_kind,
            "職員名": staff_name, "勤務日": target.strftime("%Y%m%d"),
            "従業員コード": valid_staff[staff_name], "パターンコード": pattern_code,
        })

    preview_rows.sort(key=lambda item: (
        item["勤務日"], item["従業員コード"], item["パターンコード"],
    ))
    seen = set()
    for row in preview_rows:
        key = (row["従業員コード"], row["勤務日"])
        if key in seen:
            errors.append(_issue(
                "重大エラー", row["職員名"], row["元の日付"],
                f"同一職員・同一勤務日の行が重複しています（{key[1]}）",
            ))
        seen.add(key)
    return _result(preview_rows, errors, warnings)
