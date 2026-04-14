# -*- coding: utf-8 -*-
"""
발주사양 셀(통합 전 한 컬럼) → 유형(order_spec) + 상세(order_spec_detail).

유형 판별:
- 문자열에 「신규」가 포함되면 → 신규제작
- 그렇지 않고 「변경」이 포함되면 → 사양변경
- 둘 다 아니면 → 중복제작

상세(order_spec_detail)에는 원문(trim) 전체를 둔다.
"""

from __future__ import annotations


def classify_order_spec_kind(text: str) -> str:
    """trim된 발주사양 문자열에 대한 유형(항상 세 가지 중 하나)."""
    if "신규" in text:
        return "신규제작"
    if "변경" in text:
        return "사양변경"
    return "중복제작"


def split_order_spec_cell(raw) -> tuple[str | None, str | None]:
    """
    반환: (order_spec 유형, order_spec_detail 원문).
    빈 셀이면 (None, None).
    """
    if raw is None:
        return None, None
    s = str(raw).strip()
    if not s:
        return None, None

    kind = classify_order_spec_kind(s)
    return kind, s
