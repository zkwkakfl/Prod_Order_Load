# -*- coding: utf-8 -*-
"""
호환용 실행 스크립트.

기존: release_manager.py 단일 파일에 모든 로직 포함
현재: 재사용을 위해 release_tool(ReleaseManager 클래��)로 코어 로직을 분리

사용 예)
  python release_manager.py
  python release_manager.py --remote prod-order
"""

from __future__ import annotations

import sys
from pathlib import Path

from release_tool.cli import main as _main


def main(argv: list[str]) -> int:
    # 기존 기본 동작: 이 파일이 있는 폴더(= 프로젝트 폴더) 기준
    project_dir = Path(__file__).resolve().parent
    return _main(["--project-dir", str(project_dir), *argv])


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

