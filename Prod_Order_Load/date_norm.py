# -*- coding: utf-8 -*-
"""
날짜 문자열 정규화(TEXT 정렬·필터와 일치시키기 위해 YYYY-MM-DD로 통일).
"""

from __future__ import annotations

import re
from datetime import date, datetime


def clean_date_text(raw: str) -> str:
    """날짜 문자열에서 불필요한 부분 제거 및 오타 보정."""
    if not raw:
        return ""
    text = raw
    text = re.sub(r"\([^)]*\)", "", text)
    text = text.strip()
    m = re.match(r"^(\d{4})-26-(\d{1,2})-(\d{1,2})$", text)
    if m:
        year, month, day = m.groups()
        text = f"{year}-{int(month)}-{int(day)}"
    return text


def parse_to_datetime(text: str) -> datetime:
    """비교·정렬용 파싱. 실패 시 datetime.min."""
    cleaned = clean_date_text(text)
    if not cleaned:
        return datetime.min
    try:
        parts = [int(p) for p in cleaned.split("-") if p]
        if len(parts) >= 3:
            year, month, day = parts[0], parts[1], parts[2]
            # 2자리 연도: 00~49 → 20xx, 50~99 → 19xx (엑셀·관례와 유사)
            if year < 100:
                year += 2000 if year < 50 else 1900
            return datetime(year, month, day)
    except (ValueError, OSError):
        pass
    return datetime.min


def normalize_date_to_iso(v) -> str | None:
    """SQLite·표시용 YYYY-MM-DD. 파싱 불가면 None."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    if not s:
        return None
    dt = parse_to_datetime(s)
    if dt == datetime.min:
        return None
    return dt.date().isoformat()
