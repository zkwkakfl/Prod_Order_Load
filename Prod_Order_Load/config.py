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

# 통합 시트 1행·SQLite 컬럼명(영문 snake_case). 순서 = 데이터 열 순서.
STANDARD_HEADERS = [
    "created_date",
    "work_order_no",
    "customer_name",
    "project_name",
    "product_name",
    "part_no",
    "process_code",
    "cust_delivery_date",
    "material_receipt_note",
    "order_spec",
    "order_spec_detail",
    "folder_label",       # 수식·계산
    "bom_file_label",     # 수식·계산
    "release_list_label", # 수식·계산
]

# 소스 엑셀 시트 머리글(한글 등) → 위 STANDARD_HEADERS와 같은 순서.
# consolidation._build_header_map 에서 열 번호로 매핑한다.
SHEET_HEADER_ALIASES_PER_COL: list[list[str]] = [
    ["날짜"],
    ["작업지시번호"],
    ["고객사"],
    ["사업명"],
    ["품명"],
    ["품번"],
    ["공정"],
    ["고객사납품", "고객사\n납품"],
    ["자재입고수량", "자재입고\n수량", "수량"],
    ["발주사양", "발주사양(생산기술검토)"],
    [],
    ["폴더명"],
    ["BOM파일명"],
    ["발행리스트"],
]
