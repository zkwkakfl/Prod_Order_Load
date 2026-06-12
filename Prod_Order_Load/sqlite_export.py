# -*- coding: utf-8 -*-
"""
통합 결과 행 데이터를 SQLite에 저장한다.
엑셀 수식 열(폴더명, BOM파일명, 발행리스트)은 동일 규칙으로 계산한 문자열을 저장한다.
"""

from __future__ import annotations

import re
import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Optional

from config import STANDARD_HEADERS
from customer_name_norm import canonicalize_customer_name
from date_norm import normalize_date_to_iso
from field_change_strip import strip_field_change_suffix

CONSOLIDATED_TABLE = "consolidated_data"
FIELD_CHANGE_LOG_TABLE = "field_change_log"
CUSTOMER_NAME_ALIASES_TABLE = "customer_name_aliases"
_COMPUTED_LABEL_COLS = frozenset({"folder_label", "bom_file_label", "release_list_label"})


def parse_material_receipt_qty(head: str | None) -> int | None:
    """
    자재입고 수량(문자열 앞부분)을 정수로 변환. 빈 값·파싱 불가는 None.
    """
    if head is None:
        return None
    t = str(head).strip().replace(",", "").replace(" ", "")
    if not t:
        return None
    try:
        x = float(t)
        if x != x:  # NaN
            return None
        return int(round(x))
    except (ValueError, OverflowError):
        pass
    m = re.search(r"-?\d+", t)
    return int(m.group(0)) if m else None


def _stringify_for_strip(v) -> str | None:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return str(v)


def _get_cell(row_data: list, col_name: str):
    try:
        idx = STANDARD_HEADERS.index(col_name) + 1
        if 0 <= idx < len(row_data):
            return row_data[idx]
    except ValueError:
        pass
    return None


def _clean_label_value(v) -> str | None:
    """폴더명·BOM파일명·발행리스트 조합용: 변경 꼬리 제거 후 문자열."""
    if v is None:
        return None
    sv = _stringify_for_strip(v)
    if not sv:
        return None
    head, _ = strip_field_change_suffix(sv)
    if head is None:
        return None
    s = str(head).strip()
    return s if s else None


def build_folder_bom_issue_labels(
    *,
    product_name=None,
    part_no=None,
    work_order_no=None,
    customer_name=None,
    project_name=None,
) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """
    폴더명·BOM파일명·발행리스트 문자열 생성.
    - folder_label: 품명(품번)
    - bom_file_label: 작업지시번호 고객사_품명(품번)
    - release_list_label: 사업명-품명(품번) (사업명 없으면 folder와 동일)
    """
    name = _clean_label_value(product_name)
    code = _clean_label_value(part_no)
    job = _clean_label_value(work_order_no)
    cust = _clean_label_value(customer_name)
    proj = _clean_label_value(project_name)

    folder = f"{name}({code})" if name and code else None
    bom = f"{job} {cust}_{name}({code})" if job and cust and name and code else None
    issue = (f"{proj}-{name}({code})" if proj else f"{name}({code})") if name and code else None
    return folder, bom, issue


def _computed_folder_bom_issue(row_data: list) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """엑셀 수식과 동일한 조합 규칙."""
    return build_folder_bom_issue_labels(
        product_name=_get_cell(row_data, "product_name"),
        part_no=_get_cell(row_data, "part_no"),
        work_order_no=_get_cell(row_data, "work_order_no"),
        customer_name=_get_cell(row_data, "customer_name"),
        project_name=_get_cell(row_data, "project_name"),
    )


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

def ensure_customer_name_aliases_table(cur: sqlite3.Cursor) -> None:
    """
    고객사명 별칭(alias) → 표준명(canonical) 테이블.
    - alias: 원문/별칭 (PRIMARY KEY)
    - canonical: 표준 고객사명
    """
    t = _quoted_ident(CUSTOMER_NAME_ALIASES_TABLE)
    cur.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {t} (
            alias TEXT PRIMARY KEY,
            canonical TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )


def load_customer_name_alias_map(cur: sqlite3.Cursor) -> dict[str, str]:
    """alias→canonical 매핑을 메모리로 로드."""
    ensure_customer_name_aliases_table(cur)
    t = _quoted_ident(CUSTOMER_NAME_ALIASES_TABLE)
    cur.execute(f"SELECT alias, canonical FROM {t}")
    out: dict[str, str] = {}
    for a, c in cur.fetchall():
        if a is None or c is None:
            continue
        sa = str(a).strip()
        sc = str(c).strip()
        if sa and sc:
            out[sa] = sc
    return out


def _drop_table_if_exists(cur: sqlite3.Cursor, table_name: str) -> None:
    cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    )
    if cur.fetchone():
        cur.execute(f"DROP TABLE IF EXISTS {_quoted_ident(table_name)}")


def material_receipt_cell_int(v) -> int | None:
    """엑셀/DB 값에서 수량변경 접미 제거 후 정수(통합 출력·검증용)."""
    if v is None:
        return None
    sv = _stringify_for_strip(v)
    if not sv:
        return None
    head, _ = strip_field_change_suffix(sv)
    return parse_material_receipt_qty(head)


def rebuild_consolidated_material_qty_integer(cur: sqlite3.Cursor) -> bool:
    """
    material_receipt_note 가 TEXT인 기존 DB를 INTEGER 열로 재생성한다.
    이미 INTEGER이면 아무 것도 하지 않는다.
    """
    cur.execute(f"PRAGMA table_info({_quoted_ident(CONSOLIDATED_TABLE)})")
    info = cur.fetchall()
    if not info:
        return False
    types = {row[1]: row[2] for row in info}
    if types.get("material_receipt_note", "").upper() == "INTEGER":
        return False

    old_cols = [row[1] for row in info]
    cur.execute(f"SELECT * FROM {_quoted_ident(CONSOLIDATED_TABLE)} ORDER BY id")
    all_rows = cur.fetchall()

    tmp = CONSOLIDATED_TABLE + "__new"
    cur.execute(f"DROP TABLE IF EXISTS {_quoted_ident(tmp)}")

    col_names = list(STANDARD_HEADERS)
    cols_sql = [
        "id INTEGER PRIMARY KEY AUTOINCREMENT",
        f'{_quoted_ident("update_at")} TEXT NOT NULL',
    ]
    for h in col_names:
        typ = "INTEGER" if h == "material_receipt_note" else "TEXT"
        cols_sql.append(f"{_quoted_ident(h)} {typ}")
    cur.execute(f"CREATE TABLE {_quoted_ident(tmp)} ({', '.join(cols_sql)})")

    insert_cols = ", ".join(["id", _quoted_ident("update_at")] + [_quoted_ident(h) for h in col_names])
    ph = ", ".join(["?"] * (2 + len(col_names)))
    ins = f"INSERT INTO {_quoted_ident(tmp)} ({insert_cols}) VALUES ({ph})"

    for tup in all_rows:
        d = {old_cols[i]: tup[i] for i in range(len(old_cols))}
        out: list = [d["id"], d["update_at"]]
        for h in col_names:
            if h == "material_receipt_note":
                raw = d.get(h)
                out.append(material_receipt_cell_int(raw))
            else:
                out.append(d.get(h))
        cur.execute(ins, out)

    cur.execute(f"DROP TABLE {_quoted_ident(CONSOLIDATED_TABLE)}")
    cur.execute(f"ALTER TABLE {_quoted_ident(tmp)} RENAME TO {_quoted_ident(CONSOLIDATED_TABLE)}")

    cur.execute(f"SELECT MAX(id) FROM {_quoted_ident(CONSOLIDATED_TABLE)}")
    mx_row = cur.fetchone()
    mx = mx_row[0] if mx_row and mx_row[0] is not None else None
    if mx is not None:
        try:
            cur.execute("DELETE FROM sqlite_sequence WHERE name = ?", (CONSOLIDATED_TABLE,))
            cur.execute(
                "INSERT INTO sqlite_sequence (name, seq) VALUES (?, ?)",
                (CONSOLIDATED_TABLE, int(mx)),
            )
        except sqlite3.OperationalError:
            pass
    return True


def ensure_field_change_log_table(cur: sqlite3.Cursor) -> None:
    """필드변경 제거 이력 테이블(없으면 생성)."""
    t = _quoted_ident(FIELD_CHANGE_LOG_TABLE)
    cur.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {t} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            consolidated_row_id INTEGER NOT NULL,
            work_order_no TEXT,
            field_name TEXT NOT NULL,
            removed_text TEXT NOT NULL,
            logged_at TEXT NOT NULL
        )
        """
    )


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

    cols_sql = [
        "id INTEGER PRIMARY KEY AUTOINCREMENT",
        f'{_quoted_ident("update_at")} TEXT NOT NULL',
    ]
    for h in col_names:
        typ = "INTEGER" if h == "material_receipt_note" else "TEXT"
        cols_sql.append(f"{_quoted_ident(h)} {typ}")
    create_sql = f"CREATE TABLE {table} ({', '.join(cols_sql)})"

    placeholders_row = ", ".join(["?"] * (1 + len(col_names)))
    insert_cols = ", ".join(
        [_quoted_ident("update_at")]
        + [_quoted_ident(h) for h in col_names]
    )
    insert_sql = f"INSERT INTO {table} ({insert_cols}) VALUES ({placeholders_row})"
    log_tbl = _quoted_ident(FIELD_CHANGE_LOG_TABLE)
    insert_audit_sql = (
        f"INSERT INTO {log_tbl} "
        f"(consolidated_row_id, work_order_no, field_name, removed_text, logged_at) "
        f"VALUES (?, ?, ?, ?, ?)"
    )
    wo_idx_in_vals = 1 + col_names.index("work_order_no")

    try:
        db_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        log(f"[SQLite 오류] 폴더 준비 실패: {db_path.parent} - {e}")
        return False

    def _normalize_row_field(
        h: str, v, alias_map: dict[str, str]
    ) -> tuple[object | None, list[tuple[str, str]]]:
        audits: list[tuple[str, str]] = []
        if v is None:
            return None, audits
        sv = _stringify_for_strip(v)
        if not sv:
            return None, audits
        head, tail = strip_field_change_suffix(sv)
        if tail:
            audits.append((h, tail))
        if h == "created_date":
            iso = normalize_date_to_iso(head) if head else None
            return iso if iso is not None else head, audits
        if h == "material_receipt_note":
            return parse_material_receipt_qty(head), audits
        if h == "customer_name":
            return canonicalize_customer_name(head, alias_map), audits
        return head if head is not None else None, audits

    try:
        conn = sqlite3.connect(str(db_path))
        try:
            cur = conn.cursor()
            # DB 파일은 유지하고, 통합 결과 테이블만 재생성한다.
            # (별칭 테이블 등 누적 자산은 보존)
            ensure_customer_name_aliases_table(cur)
            alias_map = load_customer_name_alias_map(cur)
            _drop_table_if_exists(cur, FIELD_CHANGE_LOG_TABLE)
            _drop_table_if_exists(cur, CONSOLIDATED_TABLE)
            cur.execute(create_sql)
            ensure_field_change_log_table(cur)
            for i, row_data in enumerate(rows):
                audits: list[tuple[str, str]] = []
                norm: dict[str, object | None] = {}
                for j, h in enumerate(col_names, start=1):
                    if h in _COMPUTED_LABEL_COLS:
                        continue
                    v = row_data[j] if j < len(row_data) else None
                    norm[h], field_audits = _normalize_row_field(h, v, alias_map)
                    audits.extend(field_audits)

                folder, bom, issue = build_folder_bom_issue_labels(
                    product_name=norm.get("product_name"),
                    part_no=norm.get("part_no"),
                    work_order_no=norm.get("work_order_no"),
                    customer_name=norm.get("customer_name"),
                    project_name=norm.get("project_name"),
                )
                norm["folder_label"] = folder
                norm["bom_file_label"] = bom
                norm["release_list_label"] = issue

                vals: list = [exported_iso]
                for h in col_names:
                    val = norm.get(h)
                    if h == "material_receipt_note":
                        vals.append(val)
                    else:
                        vals.append(_value_for_sql(val))
                cur.execute(insert_sql, vals)
                rid = int(cur.lastrowid)
                wo_val = vals[wo_idx_in_vals]
                for fn, removed in audits:
                    cur.execute(insert_audit_sql, (rid, wo_val, fn, removed, exported_iso))
            conn.commit()
        finally:
            conn.close()
    except sqlite3.Error as e:
        log(f"[SQLite 오류] 저장 실패: {db_path} - {e}")
        return False

    log(f"[SQLite] 저장 완료: {db_path} (총 {len(rows)}행)")
    return True
