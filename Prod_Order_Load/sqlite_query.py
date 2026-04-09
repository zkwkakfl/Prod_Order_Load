# -*- coding: utf-8 -*-
"""
통합 SQLite(consolidated_data) 조회·필터.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

CONSOLIDATED_TABLE = "consolidated_data"


def _qi(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def list_columns(conn: sqlite3.Connection) -> list[str]:
    """테이블 컬럼명 순서(물리 순)."""
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({CONSOLIDATED_TABLE})")
    return [row[1] for row in cur.fetchall()]


def query_consolidated(
    db_path: Path,
    *,
    job_contains: str = "",
    customer_contains: str = "",
    business_contains: str = "",
    product_contains: str = "",
    code_contains: str = "",
    spec_contains: str = "",
    date_from: str = "",
    date_to: str = "",
    limit: int = 50_000,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """
    필터는 부분 일치(LIKE, 대소문자 구분은 SQLite 기본).
    날짜는 TEXT(YYYY-MM-DD 권장) 기준 문자열 비교.
    반환: (컬럼명 리스트, 행 튜플 리스트).
    """
    if not db_path.is_file():
        return [], []

    conds: list[str] = ["1=1"]
    params: list[Any] = []

    def add_like(col_kr: str, needle: str) -> None:
        s = (needle or "").strip()
        if not s or s == "(전체)":
            return
        conds.append(f"{_qi(col_kr)} LIKE ? ESCAPE '\\'")
        esc = s.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
        params.append(f"%{esc}%")

    add_like("작업지시번호", job_contains)
    add_like("고객사", customer_contains)
    add_like("사업명", business_contains)
    add_like("품명", product_contains)
    add_like("품번", code_contains)
    add_like("발주사양", spec_contains)

    df = (date_from or "").strip()
    dt = (date_to or "").strip()
    if df and df != "(전체)":
        conds.append(f'{_qi("날짜")} >= ?')
        params.append(df)
    if dt and dt != "(전체)":
        conds.append(f'{_qi("날짜")} <= ?')
        params.append(dt)

    lim = max(1, min(int(limit), 200_000))

    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (CONSOLIDATED_TABLE,),
        )
        if not cur.fetchone():
            return [], []

        cols = list_columns(conn)
        if not cols:
            return [], []

        where_sql = " AND ".join(conds)
        sql = f"SELECT * FROM {CONSOLIDATED_TABLE} WHERE {where_sql} ORDER BY id DESC LIMIT ?"
        params.append(lim)
        cur.execute(sql, params)
        rows = cur.fetchall()
        return cols, rows
    finally:
        conn.close()


def get_last_exported_at(db_path: Path) -> str | None:
    """가장 최근 통합 시각(exported_at 최댓값). 없으면 None."""
    if not db_path.is_file():
        return None
    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (CONSOLIDATED_TABLE,),
        )
        if not cur.fetchone():
            return None
        cur.execute(f"SELECT MAX({_qi('exported_at')}) FROM {CONSOLIDATED_TABLE}")
        row = cur.fetchone()
        if not row or row[0] is None or str(row[0]).strip() == "":
            return None
        return str(row[0]).strip()
    finally:
        conn.close()


def fetch_distinct_column(
    db_path: Path,
    column_kr: str,
    *,
    limit: int = 400,
) -> list[str]:
    """열의 고유값 목록(빈 값 제외, 정렬)."""
    if not db_path.is_file():
        return []
    lim = max(1, min(int(limit), 2_000))
    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (CONSOLIDATED_TABLE,),
        )
        if not cur.fetchone():
            return []
        cols = list_columns(conn)
        if column_kr not in cols:
            return []
        qcol = _qi(column_kr)
        cur.execute(
            f"""
            SELECT DISTINCT {qcol} FROM {CONSOLIDATED_TABLE}
            WHERE {qcol} IS NOT NULL AND TRIM(CAST({qcol} AS TEXT)) != ''
            ORDER BY {qcol} COLLATE NOCASE
            LIMIT ?
            """,
            (lim,),
        )
        out: list[str] = []
        for (v,) in cur.fetchall():
            if v is None:
                continue
            s = str(v).strip()
            if s:
                out.append(s)
        return out
    finally:
        conn.close()
