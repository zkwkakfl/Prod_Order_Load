# -*- coding: utf-8 -*-
"""
셀 문자열에서 '한글머리글명 + 변경' 직결(공백 없음) 이후를 제거한다.
예: '디지트론 고객사변경: 위드솔루텍' -> '디지트론' (고객사 + 변경)
"""

from __future__ import annotations

import re
from functools import lru_cache

from config import SHEET_HEADER_ALIASES_PER_COL, STANDARD_HEADERS


@lru_cache(maxsize=1)
def _korean_aliases_longest_first() -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for group in SHEET_HEADER_ALIASES_PER_COL:
        for a in group:
            t = (a or "").strip()
            if t and t not in seen:
                seen.add(t)
                out.append(t)
    out.sort(key=len, reverse=True)
    return out


@lru_cache(maxsize=1)
def _change_marker_pattern() -> re.Pattern[str]:
    """가장 긴 한글 별칭부터 매칭(짧은 이름이 긴 이름의 접두로 잘못 잡히는 것 완화)."""
    parts = [re.escape(a) for a in _korean_aliases_longest_first()]
    if not parts:
        return re.compile(r"$^")
    inner = "|".join(parts)
    # 한글필드명과 '변경' 사이 공백 없음. 그 뒤 (옵션) ':'·공백은 제거 구간에 포함
    return re.compile(rf"(?:{inner})변경\s*[:：]?\s*", re.UNICODE)


def strip_field_change_suffix(text: str | None) -> tuple[str | None, str | None]:
    """
    반환: (정리된 값, 잘려 나간 뒷부분 전체 또는 None).
    값이 비문자·빈 문자열이면 그대로 둔다.
    """
    if text is None:
        return None, None
    if not isinstance(text, str):
        text = str(text)
    s = text
    if not s.strip():
        return s, None

    pat = _change_marker_pattern()
    m = pat.search(s)
    if not m:
        return s, None

    head = s[: m.start()].strip()
    tail = s[m.start() :].strip()
    return head, tail if tail else None
