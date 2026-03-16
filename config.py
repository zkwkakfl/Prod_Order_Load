# -*- coding: utf-8 -*-
"""기본 경로 및 시트/파일 설정."""

from pathlib import Path

# 스크립트 기준 기본 출력 경로 (실행 파일과 같은 폴더)
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_FILENAME = "공정발주내역.xlsx"

# 소스 데이터를 읽어올 폴더 경로 목록 (네트워크 경로)
SOURCE_FOLDER_PATHS = [
    r"\\192.168.0.205\생산관리\2025\1.김한식BJ\3.공정발주",
    r"\\192.168.0.205\생산관리\2025\2.김준성SW\3.공정발주",
    r"\\192.168.0.205\생산관리\2026\1.김한식BJ\2.공정발주",
    r"\\192.168.0.205\생산관리\2026\2.김준성SW\3.공정발주",
]

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
STANDARD_HEADERS = [
    "날짜",
    "폴더명",      # 수식
    "BOM파일명",   # 수식
    "발행리스트",  # 수식
    "품명",
    "품번",
    "작업지시번호",
    "고객사",
    "사업명",
    "자재입고\n수량",
    "고객사\n납품",
    "발주사양",
    # 필요 시 추가
]
