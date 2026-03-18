# -*- coding: utf-8 -*-
"""SemVer 2.0.0 기준 버전. 릴리스 시 VERSION 파일과 동기화한다."""
from pathlib import Path

_VERSION_FILE = Path(__file__).resolve().parent / "VERSION"


def get_version() -> str:
    try:
        text = _VERSION_FILE.read_text(encoding="utf-8").strip()
        return text or "0.0.0"
    except OSError:
        return "0.0.0"
