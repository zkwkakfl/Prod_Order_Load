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

import re


_TARGETS = ("BOM", "PCB", "메탈마스크")
_SAME_WORDS = ("동일", "동일함")
_CHANGE_WORDS = ("변경", "수정", "개정")


def _norm_text(s: str) -> str:
    # 개행/탭/연속 공백 정리
    t = str(s).replace("\r\n", "\n").replace("\r", "\n")
    t = re.sub(r"[\n\t]+", " ", t)
    t = re.sub(r" +", " ", t).strip()
    return t


def _strip_non_change_phrase(s: str) -> str:
    """
    '변경시 생산관리팀 공유'처럼 안내 문구의 '변경' 때문에 오탐이 나는 것을 방지.
    - '변경시'는 변경 키워드로 취급하지 않는다.
    """
    return s.replace("변경시", "")


def _segment_for_target(text: str, target: str) -> str:
    """
    target이 등장한 위치부터 다음 target 등장 전까지를 구간으로 본다.
    (콤마/하이픈 등 표기 흔들림이 있어도 비교적 견고)
    """
    idx = text.find(target)
    if idx < 0:
        return ""
    tail = text[idx:]
    next_positions = []
    for t in _TARGETS:
        if t == target:
            continue
        j = tail.find(t)
        if j >= 0:
            next_positions.append(j)
    end = min(next_positions) if next_positions else len(tail)
    return tail[:end]


def _target_status(text: str, target: str) -> str | None:
    """
    반환:
    - 'change': 해당 타겟이 변경으로 표시됨
    - 'same': 동일로 표시됨
    - None: 판단 불가(타겟 없음/표기 없음)
    """
    seg = _segment_for_target(text, target)
    if not seg:
        return None
    seg2 = _strip_non_change_phrase(seg)
    if any(w in seg2 for w in _CHANGE_WORDS):
        return "change"
    if any(w in seg2 for w in _SAME_WORDS):
        return "same"
    return None


def classify_order_spec_kind(text: str) -> str:
    """trim된 발주사양 문자열에 대한 유형(항상 세 가지 중 하나)."""
    t = _norm_text(text)
    if "신규" in t:
        return "신규제작"

    # 핵심 규칙: BOM/PCB/메탈마스크 3개 항목의 상태로 판정
    statuses = {tg: _target_status(t, tg) for tg in _TARGETS}
    if any(statuses[tg] is not None for tg in _TARGETS):
        if any(statuses[tg] == "change" for tg in _TARGETS):
            return "사양변경"
        # 하나라도 타겟이 등장했는데 change가 없다면, 모두 동일로 간주 가능한 경우 중복제작
        # (예: 'BOM,PCB,메탈마스크-동일함' + '변경시 생산관리팀 공유' → 중복제작)
        if all(statuses[tg] in ("same", None) for tg in _TARGETS):
            return "중복제작"

    # 폴백: 타겟 표기가 없는 경우 기존 규칙을 더 안전하게 적용
    # '변경시' 같은 안내 문구는 제거 후 판단
    t2 = _strip_non_change_phrase(t)
    if "변경" in t2:
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
