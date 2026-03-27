# -*- coding: utf-8 -*-
"""
통합 결과 행 데이터를 SQLite에 저장한다.
엑셀 수식 열(폴더명, BOM파일명, 발행리스트)은 동일 규칙으로 계산한 문자열을 저장한다.
"""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Optional

from config import STANDARD_HEADERS


def _get_cell(row_data: list, col_name: str):
    try:
        idx = STANDARD_HEADERS.index(col_name) + 1
        if 0 <= idx < len(row_data):
            return row_data[idx]
    except ValueError:
        pass
    return None


def _computed_folder_bom_issue(row_data: list) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """엑셀 수식과 동일한 조합 규칙."""
    name = _get_cell(row_data, "품명")
    code = _get_cell(row_data, "품번")
    job = _get_cell(row_data, "작업지시번호")
    cust = _get_cell(row_data, "고객사")
    proj = _get_cell(row_data, "사업명")

    folder = f"{name}({code})" if name and code else None
    bom = f"{job} {cust}_{name}({code})" if job and cust and name and code else None
    issue = f"{proj}-{name}({code})" if proj and name and code else None
    return folder, bom, issue


def _value_for_sql(v) -> Optional[str]:
    """SQLite TEXT 컬럼용 문자열. None은 NULL."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    return str(v).strip() if isinstance(v, str) else str(v)


def _quoted_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def save_consolidated_to_sqlite(
    rows: list[list],
    db_path: Path,
    log: Callable[[str], None],
) -> bool:
    """
    통합된 row 리스트(행당 1-based 인덱스, STANDARD_HEADERS 순)를 db_path에 저장한다.
    기존 동일 파일이 있으면 덮어쓴다(테이블 재생성).
    """
    if not rows:
        log("[SQLite] 저장할 행이 없어 DB 파일을 만들지 않습니다.")
        return True

    table = "consolidated_data"
    col_names = list(STANDARD_HEADERS)
    exported_iso = datetime.now().isoformat(timespec="seconds")

    folder_v, bom_v, issue_v = [], [], []
    for row_data in rows:
        f, b, i = _computed_folder_bom_issue(row_data)
        folder_v.append(f)
        bom_v.append(b)
        issue_v.append(i)

    cols_sql = [
        "id INTEGER PRIMARY KEY AUTOINCREMENT",
        f'{_quoted_ident("exported_at")} TEXT NOT NULL',
    ]
    for h in col_names:
        cols_sql.append(f"{_quoted_ident(h)} TEXT")
    create_sql = f"CREATE TABLE {table} ({', '.join(cols_sql)})"

    placeholders_row = ", ".join(["?"] * (1 + len(col_names)))
    insert_cols = ", ".join(
        [_quoted_ident("exported_at")]
        + [_quoted_ident(h) for h in col_names]
    )
    insert_sql = f"INSERT INTO {table} ({insert_cols}) VALUES ({placeholders_row})"

    try:
        db_path.parent.mkdir(parents=True, exist_ok=True)
        if db_path.exists():
            db_path.unlink()
    except OSError as e:
        log(f"[SQLite 오류] 파일 준비 실패: {db_path} - {e}")
        return False

    try:
        conn = sqlite3.connect(str(db_path))
        try:
            cur = conn.cursor()
            cur.execute(create_sql)
            for i, row_data in enumerate(rows):
                vals = [exported_iso]
                for j, h in enumerate(col_names, start=1):
                    if h == "폴더명":
                        vals.append(_value_for_sql(folder_v[i]))
                    elif h == "BOM파일명":
                        vals.append(_value_for_sql(bom_v[i]))
                    elif h == "발행리스트":
                        vals.append(_value_for_sql(issue_v[i]))
                    else:
                        v = row_data[j] if j < len(row_data) else None
                        vals.append(_value_for_sql(v))
                cur.execute(insert_sql, vals)
            conn.commit()
        finally:
            conn.close()
    except sqlite3.Error as e:
        log(f"[SQLite 오류] 저장 실패: {db_path} - {e}")
        try:
            if db_path.exists():
                db_path.unlink()
        except OSError:
            pass
        return False

    log(f"[SQLite] 저장 완료: {db_path} (총 {len(rows)}행)")
    return True
