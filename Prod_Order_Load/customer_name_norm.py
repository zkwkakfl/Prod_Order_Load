# -*- coding: utf-8 -*-
"""
고객사명 정규화(공백/괄호) + alias→canonical 적용용 공용 모듈.

정규화 규칙(요약):
- 줄바꿈/탭/연속 공백을 1칸으로 정리
- 괄호 안에 한글이 없으면(영문 블록): 괄호 안 텍스트만 추출(여러 개면 공백으로 연결)
- 그런 괄호가 없으면: 괄호 블록을 제거한 문자열을 반환
"""

from __future__ import annotations

import re

# 한글 음절 + 호환 자모 (괄호 안에 하나라도 있으면 '비영문'으로 간주)
_HANGUL = re.compile(r"[\u3131-\u318F\uAC00-\uD7A3]")

# 반각 ( ) — 중첩 없이 한 단계만 처리
_PAREN_HALF = re.compile(r"\(([^()]*)\)")
# 전각 （ ）
_PAREN_FULL = re.compile(r"（([^（）]*)）")


def collapse_ws(value: str) -> str:
    s = str(value).replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[\n\t]+", " ", s)
    s = re.sub(r" +", " ", s).strip()
    return s


def _inner_is_latinish(inner: str) -> bool:
    """한글이 없으면 True (괄호 안을 '영문 블록'으로 본다)."""
    t = inner.strip()
    if not t:
        return False
    return _HANGUL.search(t) is None


def _next_paren_match(s: str) -> re.Match[str] | None:
    """문자열에서 가장 앞에 나오는 반각/전각 괄호 매치 하나."""
    m1 = _PAREN_HALF.search(s)
    m2 = _PAREN_FULL.search(s)
    candidates = [m for m in (m1, m2) if m]
    if not candidates:
        return None
    return min(candidates, key=lambda m: m.start())


def normalize_customer_name(value: str | None) -> str | None:
    """
    - 영문만 있는 괄호가 하나라도 있으면: 그 괄호 안 문자열만(여러 개면 공백 구분) 반환.
    - 그런 괄호가 없으면: 괄호 블록을 제거한 뒤 공백 정리.
    """
    if value is None:
        return None
    s = collapse_ws(str(value))
    if not s:
        return None

    english_parts: list[str] = []
    rest = s
    while True:
        m = _next_paren_match(rest)
        if not m:
            break
        inner = m.group(1).strip()
        before = rest[: m.start()]
        after = rest[m.end() :]
        if _inner_is_latinish(inner):
            english_parts.append(inner)
        rest = before + " " + after
    rest = collapse_ws(rest)

    if english_parts:
        out = collapse_ws(" ".join(english_parts))
        return out if out else None
    return rest if rest else None


def apply_alias(normalized: str | None, alias_map: dict[str, str]) -> str | None:
    if normalized is None:
        return None
    key = normalized.strip()
    if not key:
        return None
    return alias_map.get(key, key)


def canonicalize_customer_name(raw: str | None, alias_map: dict[str, str]) -> str | None:
    """raw → normalize_customer_name → alias 적용."""
    norm = normalize_customer_name(raw)
    return apply_alias(norm, alias_map)

