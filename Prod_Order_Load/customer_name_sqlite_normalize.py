# -*- coding: utf-8 -*-
"""
consolidated_data.customer_name 정규화:
- 반각 () / 전각 （） 괄호 처리
  · 괄호 안에 한 글도 한글이 없으면(영문 블록): 그 안의 문자열만 남기고 나머지 전부 삭제.
    여러 개면 공백 한 칸으로 이어 붙임. 예) "위드 (WithTech)" -> "WithTech"
  · 괄호 안에 한글이 하나라도 있으면: 괄호와 안쪽만 삭제(바깥 텍스트는 유지).
- 줄바꿈·탭·연속 공백 정리
"""

from __future__ import annotations

import argparse
import random
import re
import sqlite3
import sys
from pathlib import Path

from sqlite_export import CONSOLIDATED_TABLE, _quoted_ident

# 한글 음절 + 호환 자모 (괄호 안에 하나라도 있으면 '비영문'으로 간주)
_HANGUL = re.compile(r"[\u3131-\u318F\uAC00-\uD7A3]")

# 반각 ( ) — 중첩 없이 한 단계만 처리
_PAREN_HALF = re.compile(r"\(([^()]*)\)")
# 전각 （ ）
_PAREN_FULL = re.compile(r"（([^（）]*)）")


def _collapse_ws(s: str) -> str:
    s = s.replace("\r\n", "\n").replace("\r", "\n")
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
    영문만 있는 괄호가 하나라도 있으면: 그 괄호 안 문자열만(여러 개면 공백 구분) 반환.
    그런 괄호가 없으면: 한글 포함 괄호만 제거한 뒤 공백 정리.
    """
    if value is None:
        return None
    s = _collapse_ws(str(value))
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
    rest = _collapse_ws(rest)

    if english_parts:
        return _collapse_ws(" ".join(english_parts))
    return rest if rest else None


def _print_samples(
    updates: list[tuple[int, str | None, str | None]],
    *,
    limit: int,
    random_sample: bool,
) -> None:
    if not updates:
        print("샘플 없음: 변경 대상이 없습니다.")
        return

    sample_rows = list(updates)
    if random_sample:
        sample_rows = random.sample(sample_rows, k=min(limit, len(sample_rows)))
    else:
        sample_rows = sample_rows[:limit]

    print(f"샘플 {len(sample_rows)}건")
    for row_id, raw, new_v in sample_rows:
        print(f"[id={row_id}]")
        print(f"  원본: {raw!r}")
        print(f"  결과: {new_v!r}")


def _run(db_path: Path, *, apply: bool, samples: int, random_sample: bool) -> int:
    if not db_path.is_file():
        print(f"[오류] DB 파일 없음: {db_path}", file=sys.stderr)
        return 1

    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute(
            f"SELECT id, {_quoted_ident('customer_name')} "
            f"FROM {_quoted_ident(CONSOLIDATED_TABLE)} "
            f"ORDER BY id"
        )
        rows = cur.fetchall()
    finally:
        conn.close()

    updates: list[tuple[int, str | None, str | None]] = []
    for row_id, raw in rows:
        new_v = normalize_customer_name(raw)
        if (raw or "") != (new_v or ""):
            updates.append((row_id, raw, new_v))

    print(f"총 행: {len(rows)}, 변경 대상: {len(updates)}")
    _print_samples(updates, limit=samples, random_sample=random_sample)

    if not apply:
        print("(적용 안 함: --apply 로 실제 UPDATE)")
        return 0

    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        sql = (
            f"UPDATE {_quoted_ident(CONSOLIDATED_TABLE)} "
            f"SET {_quoted_ident('customer_name')} = ?, "
            f"{_quoted_ident('update_at')} = datetime('now', 'localtime') "
            f"WHERE id = ?"
        )
        for row_id, _raw, new_v in updates:
            cur.execute(sql, (new_v, row_id))
        conn.commit()
    finally:
        conn.close()

    print(f"[완료] {len(updates)}건 UPDATE 반영")
    return 0


def main() -> int:
    base = Path(__file__).resolve().parent
    default_db = base / "공정발주내역.sqlite"

    p = argparse.ArgumentParser(description="SQLite 고객사명(customer_name) 괄호·공백 정규화")
    p.add_argument(
        "--db",
        type=Path,
        default=default_db,
        help=f"SQLite 경로 (기본: {default_db.name})",
    )
    p.add_argument("--apply", action="store_true", help="실제 UPDATE (없으면 미리보기만)")
    p.add_argument(
        "--samples",
        type=int,
        default=10,
        help="샘플 출력 개수 (기본: 10)",
    )
    p.add_argument(
        "--random-sample",
        action="store_true",
        help="앞에서부터가 아니라 변경 대상 중 랜덤 샘플 출력",
    )
    args = p.parse_args()
    return _run(
        args.db.resolve(),
        apply=args.apply,
        samples=max(0, args.samples),
        random_sample=args.random_sample,
    )


if __name__ == "__main__":
    raise SystemExit(main())
