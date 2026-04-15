# -*- coding: utf-8 -*-
"""
consolidated_data.customer_name 정규화:
1) 괄호/공백 정규화(customer_name_norm.normalize_customer_name)
2) customer_name_aliases(alias→canonical) 적용
3) --apply 시 consolidated_data.customer_name을 canonical로 UPDATE
4) 치환 전 값(정규화된 alias)은 customer_name_aliases에 upsert(재통합 시 재사용)
"""

from __future__ import annotations

import argparse
import random
import sqlite3
import sys
from pathlib import Path

from customer_name_norm import canonicalize_customer_name, collapse_ws, normalize_customer_name
from sqlite_export import (
    CONSOLIDATED_TABLE,
    CUSTOMER_NAME_ALIASES_TABLE,
    _quoted_ident,
    ensure_customer_name_aliases_table,
    load_customer_name_alias_map,
)


DEFAULT_ALIASES: dict[str, str] = {
    # 스크린샷 기반(필요 시 여기 추가)
    "KB테크": "KB-TECH",
    "LIG 정밀기술": "LIG정밀기술",
    "글랜에어테크놀리지": "글랜에어테크놀로지",
    "글린에어테크놀로지": "글랜에어테크놀로지",
    "시그윅스": "시그웍스",
    "아이스펙": "아이스팩",
    "웨이브": "웨이브일렉트로닉스",
    "웨이브 일렉트로닉스": "웨이브일렉트로닉스",
}


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


def _upsert_aliases(cur: sqlite3.Cursor, alias_to_canon: dict[str, str]) -> int:
    ensure_customer_name_aliases_table(cur)
    t = _quoted_ident(CUSTOMER_NAME_ALIASES_TABLE)
    sql = (
        f"INSERT INTO {t} (alias, canonical, updated_at) "
        f"VALUES (?, ?, datetime('now','localtime')) "
        f"ON CONFLICT(alias) DO UPDATE SET "
        f"canonical=excluded.canonical, updated_at=excluded.updated_at"
    )
    n = 0
    for a, c in alias_to_canon.items():
        sa = collapse_ws(str(a))
        sc = collapse_ws(str(c))
        if not sa or not sc:
            continue
        cur.execute(sql, (sa, sc))
        n += 1
    return n


def _run(
    db_path: Path,
    *,
    apply: bool,
    samples: int,
    random_sample: bool,
    sync_default_aliases: bool,
) -> int:
    if not db_path.is_file():
        print(f"[오류] DB 파일 없음: {db_path}", file=sys.stderr)
        return 1

    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        if sync_default_aliases:
            n = _upsert_aliases(cur, DEFAULT_ALIASES)
            conn.commit()
            print(f"[별칭] 기본 매핑 {n}건 upsert")
        alias_map = load_customer_name_alias_map(cur)
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
        new_v = canonicalize_customer_name(raw, alias_map)
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
        ensure_customer_name_aliases_table(cur)
        sql = (
            f"UPDATE {_quoted_ident(CONSOLIDATED_TABLE)} "
            f"SET {_quoted_ident('customer_name')} = ?, "
            f"{_quoted_ident('update_at')} = datetime('now', 'localtime') "
            f"WHERE id = ?"
        )
        # 치환 전 값은 alias로 upsert하여 재통합 시에도 적용되게 한다.
        alias_upserts: dict[str, str] = {}
        for row_id, raw, new_v in updates:
            raw_norm = normalize_customer_name(raw)
            if raw_norm and new_v and raw_norm != new_v:
                alias_upserts[raw_norm] = new_v
            cur.execute(sql, (new_v, row_id))
        if alias_upserts:
            n2 = _upsert_aliases(cur, alias_upserts)
            print(f"[별칭] 변경분 alias {n2}건 upsert")
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
    p.add_argument(
        "--sync-default-aliases",
        action="store_true",
        help="DEFAULT_ALIASES(코드 내 매핑)를 customer_name_aliases에 upsert",
    )
    args = p.parse_args()
    return _run(
        args.db.resolve(),
        apply=args.apply,
        samples=max(0, args.samples),
        random_sample=args.random_sample,
        sync_default_aliases=bool(args.sync_default_aliases),
    )


if __name__ == "__main__":
    raise SystemExit(main())
