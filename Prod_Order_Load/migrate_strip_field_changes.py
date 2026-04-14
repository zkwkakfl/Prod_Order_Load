# -*- coding: utf-8 -*-
"""
기존 공정발주 SQLite에서 '한글머리글+변경' 접미를 제거하고, 잘린 내용은 field_change_log에 적재한다.

사용:
  python migrate_strip_field_changes.py [경로.sqlite]
기본 경로: 스크립트 폴더의 공정발주내역.sqlite
"""

from __future__ import annotations

import argparse
import sqlite3
from datetime import datetime
from pathlib import Path

from config import STANDARD_HEADERS
from date_norm import normalize_date_to_iso
from field_change_strip import strip_field_change_suffix
from sqlite_export import (
    CONSOLIDATED_TABLE,
    FIELD_CHANGE_LOG_TABLE,
    ensure_field_change_log_table,
    parse_material_receipt_qty,
    rebuild_consolidated_material_qty_integer,
    _quoted_ident,
    _stringify_for_strip,
)


def _qi(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def migrate(db_path: Path) -> tuple[int, int, bool]:
    """
    반환: (갱신된 셀 수, 감사 로그 삽입 행 수, material_receipt_note INTEGER 재구성 여부)
    """
    if not db_path.is_file():
        raise FileNotFoundError(str(db_path))

    logged_at = datetime.now().isoformat(timespec="seconds")
    cells_updated = 0
    audits_inserted = 0
    rebuilt_int = False

    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (CONSOLIDATED_TABLE,),
        )
        if not cur.fetchone():
            raise RuntimeError(f"테이블 {CONSOLIDATED_TABLE} 가 없습니다.")

        cur.execute(f"PRAGMA table_info({_qi(CONSOLIDATED_TABLE)})")
        existing = {row[1] for row in cur.fetchall()}
        data_cols = [h for h in STANDARD_HEADERS if h in existing]
        if not data_cols:
            raise RuntimeError("STANDARD_HEADERS 에 해당하는 열이 DB에 없습니다.")

        ensure_field_change_log_table(cur)

        cur.execute(
            f"SELECT id, {_qi('work_order_no')}, "
            + ", ".join(_qi(h) for h in data_cols)
            + f" FROM {_qi(CONSOLIDATED_TABLE)} ORDER BY id"
        )
        rows = cur.fetchall()
        id_idx = 0
        wo_idx = 1
        col_offsets = {h: 2 + i for i, h in enumerate(data_cols)}

        insert_audit = (
            f"INSERT INTO {_quoted_ident(FIELD_CHANGE_LOG_TABLE)} "
            f"(consolidated_row_id, work_order_no, field_name, removed_text, logged_at) "
            f"VALUES (?, ?, ?, ?, ?)"
        )

        for row in rows:
            rid = row[id_idx]
            wo = row[wo_idx]
            updates: dict[str, str | None] = {}
            audits: list[tuple[str, str]] = []

            for h in data_cols:
                raw = row[col_offsets[h]]
                if raw is None:
                    continue
                sv = _stringify_for_strip(raw)
                if not sv:
                    continue

                if h == "material_receipt_note":
                    head, tail = strip_field_change_suffix(sv)
                    if not tail:
                        continue
                    new_val = parse_material_receipt_qty(head)
                    updates[h] = new_val
                    audits.append((h, tail))
                    continue

                head, tail = strip_field_change_suffix(sv)
                if not tail:
                    continue

                if h == "created_date":
                    new_val = normalize_date_to_iso(head) if head else None
                else:
                    new_val = head if head is not None else ""

                updates[h] = new_val
                audits.append((h, tail))

            if updates:
                set_parts = ", ".join(f"{_qi(c)} = ?" for c in updates)
                params = list(updates.values()) + [rid]
                cur.execute(
                    f"UPDATE {_qi(CONSOLIDATED_TABLE)} SET {set_parts} WHERE id = ?",
                    params,
                )
                cells_updated += len(updates)

            for fn, removed in audits:
                cur.execute(insert_audit, (rid, wo, fn, removed, logged_at))
                audits_inserted += 1

        rebuilt_int = rebuild_consolidated_material_qty_integer(cur)

        conn.commit()
    finally:
        conn.close()

    return cells_updated, audits_inserted, rebuilt_int


def main() -> None:
    base = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description="필드변경 접미 제거 + field_change_log 적재")
    ap.add_argument(
        "db",
        nargs="?",
        default=str(base / "공정발주내역.sqlite"),
        help="SQLite 파일 경로",
    )
    args = ap.parse_args()
    p = Path(args.db)
    u, a, rebuilt = migrate(p)
    print(f"완료: {p}")
    print(f"  갱신된 셀: {u}개")
    print(f"  감사 로그 행: {a}개")
    print(f"  material_receipt_note INTEGER 재구성: {'예' if rebuilt else '아니오(이미 INTEGER)'}")


if __name__ == "__main__":
    main()
