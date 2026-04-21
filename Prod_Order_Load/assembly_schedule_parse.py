# -*- coding: utf-8 -*-
"""
조립공정일정 셀(한 칸·여러 줄) → 자재검수 / SMT / IMT / 검사 날짜.

형식 예:
  자재검수 : 2025-03-15
  SMT : 2025-03-16
라벨 순서 무관. 값 없음·파싱 불가는 None.
"""

from __future__ import annotations

import re

from date_norm import clean_date_text, normalize_date_to_iso

_LABEL_PATTERN = re.compile(
    r"^\s*(자재검수|SMT|IMT|검사)\s*[:：]\s*(.*)$",
)


def parse_assembly_schedule_cell(raw) -> tuple[str | None, str | None, str | None, str | None]:
    """
    반환: (자재검수, SMT, IMT, 검사) — 각각 YYYY-MM-DD 또는 None.
    """
    if raw is None:
        return (None, None, None, None)
    s = str(raw).strip()
    if not s:
        return (None, None, None, None)

    slots: dict[str, str | None] = {
        "자재검수": None,
        "SMT": None,
        "IMT": None,
        "검사": None,
    }
    for line in re.split(r"\r\n|\n|\r", s):
        line = line.strip()
        if not line:
            continue
        m = _LABEL_PATTERN.match(line)
        if not m:
            continue
        label, rest = m.group(1), (m.group(2) or "").strip()
        if not rest:
            slots[label] = None
            continue
        iso = normalize_date_to_iso(clean_date_text(rest))
        slots[label] = iso if iso else None

    return (
        slots["자재검수"],
        slots["SMT"],
        slots["IMT"],
        slots["검사"],
    )
