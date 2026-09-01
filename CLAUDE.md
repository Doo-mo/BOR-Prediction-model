# CLAUDE.md

이 파일은 Claude Code가 이 저장소에서 작업할 때 참고하는 프로젝트 지침입니다.
Copilot에서 Claude로 인수인계하기 위해 작성되었습니다.

## 1. 프로젝트 목적

- **LNG(액화천연가스) 극저온 저장 탱크의 BOR(Boil-Off Rate, 증발률, 단위 `%/day`) 예측용 엑셀 모델**을
  Python 스크립트로 자동 생성하는 저장소입니다.
- **BOR은 병상이용률(Bed Occupancy Rate)이 아닙니다.** (`README.md`, `generate_bor_model.py` docstring에 명시)
- 산출물은 수식이 살아 있는 엑셀 워크북 `LNG_BOR_Model.xlsx`이며, 사용자는 `Input` 시트 값만 바꾸면
  `BOR_Calc` 시트의 결과가 엑셀 수식으로 재계산됩니다.

## 2. 아키텍처 / 모델 요약

```
generate_bor_model.py  ──(openpyxl + matplotlib)──▶  LNG_BOR_Model.xlsx
                                                          │
validate_model.py  ──(재생성 후 구조·수식·수치 검증)──────┘
```

- 워크북 시트 4개(탭 이름은 ASCII): `Input`, `PropertyDB`, `BOR_Calc`, `Measure`.
- 시트 내부 라벨·설명은 한글, 탭 이름만 영문입니다.
- 계산은 정상상태(steady-state) 열전달 기반이며 핵심 관계식은 `README.md`에 정리되어 있습니다.
  (요약: `Q = ΔT / R_total`, `R_total = 두께 / (keff × A)`, `BOR [%/day] = BOG일일증발량 / LNG총질량 × 100`)
- 시트 간 행 참조 불일치를 막기 위해 `Input` 시트 레이아웃 행 번호는
  `generate_bor_model.py`의 `INPUT_ROWS` 딕셔너리 한 곳에서 관리합니다.
- 탱크 도면 PNG 2장(수직형/수평형)은 `render_tank_images()`가 `tempfile.TemporaryDirectory()`에
  생성한 뒤 `Input` 시트에 삽입하며, 저장소에는 남기지 않습니다.

## 3. 환경 / 실행 / 검증 명령 (실제 실행으로 확인됨)

```bash
pip install -r requirements.txt   # openpyxl>=3.1.0, matplotlib>=3.7.0
python generate_bor_model.py      # LNG_BOR_Model.xlsx 생성
python validate_model.py          # 구조·수식·수치 검증, 통과 시 exit code 0
```

- CI(`.github/workflows/validate.yml`)는 Python `3.11`, `fonts-nanum` 설치 후 위 두 스크립트를 실행하고
  `LNG_BOR_Model.xlsx`를 아티팩트로 업로드합니다. (로컬 확인은 Python 3.12에서도 성공)
- 한글 글꼴(`fonts-nanum` 등)이 없으면 도면 라벨에 `Glyph ... missing from font` 경고가 나고
  한글이 깨집니다. 스크립트 자체는 실패하지 않습니다.
- **린터/포매터/유닛 테스트 프레임워크는 저장소에 없습니다.** 검증 수단은 `validate_model.py` 뿐입니다.
  없는 도구를 새로 도입하기 전에 반드시 사용자에게 확인하세요.
- 주의: `validate_model.py`는 검증 과정에서 `LNG_BOR_Model.xlsx`를 **재생성**합니다.
  코드 변경 없이 실행만 했는데 워크북이 변경 상태로 남으면 `git checkout LNG_BOR_Model.xlsx`로 되돌리세요.

## 4. 주요 파일 / 엔트리 포인트

| 경로 | 역할 |
|---|---|
| `generate_bor_model.py` | 엔트리 포인트. `main()`이 `build_sheet1~4`를 호출해 워크북 생성 |
| `validate_model.py` | 검증 스크립트. `generate_bor_model`의 `INPUT_ROWS`, `main`을 import |
| `LNG_BOR_Model.xlsx` | 생성 산출물이지만 저장소에 커밋되어 있음(바이너리) |
| `requirements.txt` | 런타임 의존성 2개 |
| `.github/workflows/validate.yml` | 생성 + 검증 CI |
| `README.md` | 모델 물리 배경, 시트 구성, 물성 DB 표, 사용법 |
| `__pycache__/*.pyc` | 실수로 커밋된 바이트코드(정리 후보, `docs/handoff/TODO.md` 참고) |

## 5. 코딩 / 데이터 규칙 (저장소에서 관찰된 관습)

- Python 표준 라이브러리 + `openpyxl` + `matplotlib`만 사용합니다. 의존성 추가는 신중히.
- docstring과 주석은 한글, 식별자·함수명은 영문 snake_case입니다.
- 셀 스타일은 `_hdr` / `_lbl` / `_inp` / `_calc` / `_result` / `_sec` 같은 소문자 헬퍼로 통일합니다.
- 색상은 `CLR_*` 상수, 임계값·보정계수는 모듈 상단 상수로 선언합니다(하드코딩 금지).
- 워크북은 값이 아니라 **엑셀 수식**으로 계산되어야 합니다. openpyxl로 계산된 값만 써넣지 마세요.
- 시트 탭 이름은 ASCII로 유지합니다(한글 탭명이 수식 손실을 유발한 이력 있음 — `docs/handoff/DECISIONS.md`).
- 물성값(단열재 k, LNG 성분 물성, 배관 계수)은 `README.md` 표와 `PropertyDB` 시트가 일치해야 합니다.
  값을 바꿀 때는 양쪽 모두 갱신하세요.
- 수식/행 번호를 바꿀 때는 `INPUT_ROWS`와 `validate_model.py`의 상수(`INSULATION_ROWS`,
  `COMPONENT_ROWS`, `PIPE_FACTOR_ROWS` 등)를 함께 확인하세요.

## 6. 보안 / 데이터 취급 지침

- API 키, 토큰, 비밀번호, 사내 서버·NAS 주소, 사설 엔드포인트, 개인정보를 **코드·문서·커밋 메시지에
  절대 넣지 마세요.** 존재하지 않는 자격증명이나 운영 데이터를 지어내지도 마세요.
- 실제 현장 운전 데이터·측정치를 문서에 붙여넣지 말고, 필요하면 사용자에게 로컬 파일로 다루도록 안내하세요.
- 대용량 바이너리(스프레드시트, 데이터셋)를 문서나 저장소에 중복 추가하지 마세요.
- 모델 가정, 수식, 명령, 물성값을 추측으로 채우지 말고 근거가 없으면 "미확인"으로 표기하세요.

## 7. Git 워크플로

- 기본 브랜치는 `main`이며, 작업은 항상 별도 브랜치에서 진행하고 PR로 병합합니다.
- 기존 `copilot/*` 브랜치를 **임의로 삭제하거나 병합하지 마세요.** 상태는
  `docs/handoff/CURRENT_STATUS.md`를 참고하고, 정리는 사용자 확인 후 진행합니다.
- `main`에 직접 푸시하거나 히스토리를 재작성(force push, rebase 후 강제 푸시)하지 마세요.
- 커밋 전 변경 파일 목록을 확인하고 `LNG_BOR_Model.xlsx`, `__pycache__` 등 의도치 않은
  바이너리 변경이 섞이지 않았는지 점검하세요.

## 8. 변경 전/후 체크리스트

1. `python generate_bor_model.py` 실행 성공 여부 확인
2. `python validate_model.py`가 `✅ 모든 검증 통과`로 끝나는지 확인
3. 수식·행 번호를 바꿨다면 `INPUT_ROWS`와 검증 상수 동기화 확인
4. `README.md` 표와 `PropertyDB` 값 일치 확인
5. `git status`로 불필요한 바이너리/캐시 변경이 없는지 확인
6. 비밀정보가 포함되지 않았는지 확인

## 9. 인수인계 문서

- @docs/handoff/PROJECT_CONTEXT.md — 프로젝트 배경과 요구사항
- @docs/handoff/CURRENT_STATUS.md — 현재 브랜치·PR·구현 상태
- @docs/handoff/DECISIONS.md — 지금까지의 기술적 결정과 근거
- @docs/handoff/TODO.md — 남은 작업과 우선순위
- @docs/handoff/FIRST_PROMPT.md — Claude에 붙여 넣을 첫 프롬프트
