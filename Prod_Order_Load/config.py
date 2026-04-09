# -*- coding: utf-8 -*-
"""기본 경로 및 시트/파일 설정."""

import sys
from pathlib import Path


def _app_base_dir() -> Path:
    """소스 실행 시 스크립트 폴더, PyInstaller 빌드 시 exe 폴더."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# 기본 저장·설정 파일 기준 폴더
DEFAULT_OUTPUT_DIR = _app_base_dir()
DEFAULT_OUTPUT_FILENAME = "공정발주내역.xlsx"

# 기본 소스 데이터를 읽어올 폴더 경로 목록 (네트워크 경로)
DEFAULT_SOURCE_FOLDER_PATHS = [
    r"\\192.168.0.205\생산관리\2025\1.김한식BJ\3.공정발주",
    r"\\192.168.0.205\생산관리\2025\2.김준성SW\3.공정발주",
    r"\\192.168.0.205\생산관리\2026\1.김한식BJ\2.공정발주",
    r"\\192.168.0.205\생산관리\2026\2.김준성SW\3.공정발주",
]

# 설정 파일 이름 (소스 경로를 동적으로 관리)
SOURCE_PATHS_FILE = DEFAULT_OUTPUT_DIR / "source_paths.json"
# 폴더 생성: 기본 경로 + 하위 폴더 이름 목록 (CreatedFolders.bas 이식, 파일 복사 없음)
FOLDER_CREATE_SETTINGS_FILE = DEFAULT_OUTPUT_DIR / "folder_create_settings.json"

# 통합 시트 이름 (출력 워크북 내)
DEST_SHEET_NAME = "공정발주내역"

# 소스 시트 필터: 이름에 포함되어야 함
SOURCE_SHEET_NAME_CONTAINS = "작업 발주"
# 무시할 시트: 이름에 포함되면 스킵
IGNORE_SHEET_NAME_CONTAINS = "누락주의"

# 소스 시트에서 머리글 행(1-based), 데이터 시작 행(1-based)
SOURCE_HEADER_ROW = 3
SOURCE_DATA_START_ROW = 4
# 소스 데이터 시작 열 (C = 3)
SOURCE_FIRST_COL = 3

# 통합 시트 1행 기준 헤더 순서 (열 인덱스 = 순서). 수식 컬럼은 여기 포함.
# 요구사항: 날짜, 작업지시번호, 고객사, 사업명, 품명, 품번, 공정, 자재입고수량, 고객사납품, 발주사양, 폴더명, BOM파일명, 발행리스트
STANDARD_HEADERS = [
    "날짜",
    "작업지시번호",
    "고객사",
    "사업명",
    "품명",
    "품번",
    "공정",
    "고객사납품",
    "자재입고수량",
    "발주사양",
    "폴더명",      # 수식
    "BOM파일명",   # 수식
    "발행리스트",  # 수식
    # 필요 시 추가
]
