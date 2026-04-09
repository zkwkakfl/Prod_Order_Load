# -*- coding: utf-8 -*-
"""
폴더 구조 생성 (VBA CreatedFolders.bas 이식).
- 기본 경로 아래에 '폴더명'별 상위 폴더를 만들고, 설정된 하위 폴더 이름마다 하위 디렉터리를 만든다.
- 엑셀 복사·워크북 생성은 하지 않는다.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Callable

from config import DEFAULT_OUTPUT_DIR, FOLDER_CREATE_SETTINGS_FILE

_ILLEGAL = re.compile(r'[\\/:*?"<>|]')


def replace_illegal_chars(name: str) -> str:
    """Windows 파일명에 쓸 수 없는 문자 제거 (VBA ReplaceIllegalChars 대응)."""
    if not name:
        return ""
    s = _ILLEGAL.sub("", str(name).strip())
    if s.endswith("."):
        s = s.rstrip(".")
    return s.strip()


def load_folder_create_settings() -> dict:
    """기본 경로(str), 하위 폴더 이름 목록(list[str])."""
    default = {"base_path": str(DEFAULT_OUTPUT_DIR), "subfolders": []}
    try:
        if not FOLDER_CREATE_SETTINGS_FILE.is_file():
            return default
        with FOLDER_CREATE_SETTINGS_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
        base = (data.get("base_path") or "").strip()
        subs = data.get("subfolders") or data.get("subfolder_names") or []
        if isinstance(subs, str):
            subs = [ln.strip() for ln in subs.splitlines() if ln.strip()]
        elif isinstance(subs, list):
            subs = [str(x).strip() for x in subs if str(x).strip()]
        else:
            subs = []
        if base:
            default["base_path"] = base
        default["subfolders"] = subs
        return default
    except (OSError, json.JSONDecodeError):
        return default


def save_folder_create_settings(base_path: str, subfolders: list[str]) -> None:
    data = {
        "base_path": (base_path or "").strip(),
        "subfolders": [s.strip() for s in subfolders if s.strip()],
    }
    FOLDER_CREATE_SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with FOLDER_CREATE_SETTINGS_FILE.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def ensure_base_directory(base: Path, log: Callable[[str], None] | None = None) -> bool:
    """기본 경로가 없으면 생성 시도."""
    try:
        base.mkdir(parents=True, exist_ok=True)
        return base.is_dir()
    except OSError as e:
        if log:
            log(f"[폴더] 기본 경로 생성 실패: {base} — {e}")
        return False


def create_folder_structure(
    base_path: Path,
    parent_folder_names: list[str],
    subfolder_names: list[str],
    *,
    log: Callable[[str], None] | None = None,
) -> tuple[int, int, list[str]]:
    """
    각 parent 이름마다 base/parent/ 를 만들고, subfolder_names 각각 base/parent/sub/ 생성.
    반환: (성공한 상위 폴더 수, 건너뜀 수, 오류 메시지 목록)
    """
    errs: list[str] = []
    if not ensure_base_directory(base_path, log):
        return 0, len(parent_folder_names), [f"기본 경로를 사용할 수 없습니다: {base_path}"]

    ok_parents = 0
    skipped = 0
    subs_clean = [replace_illegal_chars(s) for s in subfolder_names]
    subs_clean = [s for s in subs_clean if s]

    seen: set[str] = set()
    for raw in parent_folder_names:
        name = replace_illegal_chars(raw)
        if not name:
            skipped += 1
            if log:
                log(f"[폴더] 빈 이름 건너뜀: {raw!r}")
            continue
        key = name.casefold()
        if key in seen:
            skipped += 1
            continue
        seen.add(key)

        parent = base_path / name
        try:
            parent.mkdir(parents=True, exist_ok=True)
            for sub in subs_clean:
                (parent / sub).mkdir(parents=True, exist_ok=True)
            ok_parents += 1
            if log:
                msg = f"[폴더] 생성: {parent}"
                if subs_clean:
                    msg += f" (+ 하위 {len(subs_clean)}개)"
                log(msg)
        except OSError as e:
            errs.append(f"{name}: {e}")
            if log:
                log(f"[폴더] 오류 {name}: {e}")

    return ok_parents, skipped, errs
