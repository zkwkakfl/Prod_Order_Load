## Python port: BOM/TOP/BOT 매칭

엑셀(`.xlsx/.xlsm`)에서 BOM/TOP/BOT 시트를 읽고,
각 시트의 **자재이름/좌표명/소요량** 3개 범위를 기반으로

- 좌표명 셀 값을 `,` 기준으로 분리(Trim)
- 키 = `(좌표명, 자재이름)` 단위로 수량 합산
- **BOM 기준**으로 TOP/BOT 각각 존재·수량 일치 여부 판정
- 시트별 좌표 중복(동일 좌표가 2회 이상 등장) 검출
- 결과를 새 엑셀 파일에 시트로 기록

을 수행합니다.

### 설치

```bash
python -m pip install -r requirements.txt
```

### 실행

**1) GUI (폼) 실행** — 파일·시트·범위를 화면에서 선택

```bash
python match_bom_top_bot.py
```

- BOM / TOP / BOT 파일 각각 선택, 저장 경로(선택), 시트·자재/좌표/수량 범위 입력 후 **매칭 실행**

**2) config 파일로 실행** — 단일 파일 또는 3개 파일 모드

- 단일 파일: `config_example.json` 참고 (`excel_path` 사용)
- 3개 파일: `config_example_3files.json` 참고 (`bom_file`, `top_file`, `bot_file` 사용)

```bash
python match_bom_top_bot.py --config config.json
```

### 출력 시트

| 시트 | 설명 |
|------|------|
| `Match_Result` | 전체 (coord, material)별 BOM/TOP/BOT 수량 및 상태 |
| `Unmatched` | BOM vs (TOP+BOT) 불일치 전체 |
| `Unmatched_TOP` | **BOM 기준** TOP만 비교 시 불일치 |
| `Unmatched_BOT` | **BOM 기준** BOT만 비교 시 불일치 |
| `Coord_Duplicates` | 시트별 좌표 중복 목록 |
| `Summary` | 매칭됨(OK) 건수, 불일치(TOP/BOT) 건수, 중복 좌표 수 |

### 주의

- 3개 범위는 **모두 1열**이어야 하며, **행 개수도 동일**해야 합니다.
- 좌표가 `A1, A2, A3`처럼 한 셀에 여러 개면 각 좌표로 확장되어 키가 생성됩니다.

