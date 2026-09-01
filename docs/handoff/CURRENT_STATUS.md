# CURRENT_STATUS — 현재 상태

> 기준 시점: 2026-09-01 (이 문서 작성 시점의 GitHub API/`git log` 조회 결과)

## 구현 상태

- `main` 기준으로 워크북 생성·검증 파이프라인이 **동작합니다.** 아래 명령을 실제로 실행해 확인했습니다.

  ```bash
  pip install -r requirements.txt
  python generate_bor_model.py   # ✅ 저장 완료: LNG_BOR_Model.xlsx
  python validate_model.py       # ✅ 모든 검증 통과
  ```

- `validate_model.py`가 통과시킨 항목: 시트 구성, `BOR_Calc` 조성 수식 열 참조,
  D열 note 회귀 검사, `Input` 체적 자동계산 행, 시트 XML `<f>` 태그(63개),
  기본 입력값 독립 계산 BOR `0.03001 %/day`, `Input` 드롭다운 3개.
- 로컬 검증은 Python 3.12.3에서 수행했고, CI는 Python 3.11을 사용합니다.
- 한글 글꼴이 없는 환경에서는 도면 라벨에 `Glyph ... missing from font` 경고가 발생합니다
  (스크립트는 실패하지 않음). CI는 `fonts-nanum`을 설치합니다.

## 브랜치 현황 (증거 기반)

| 브랜치 | HEAD SHA | 연관 PR | 상태 |
|---|---|---|---|
| `main` | `1292b46` | — | 기본 브랜치. 최신 커밋은 PR #7 병합 커밋 |
| `copilot/bor-prediction-excel-model` | `2c3eb12` | #1 | PR **닫힘(미병합)** — BOR을 병상이용률로 오해한 초기 시도 |
| `copilot/lnt-boil-off-rate-prediction` | `28bf3b0` | #2 | PR 병합됨 |
| `copilot/update-bor-model` | `1c29070` | #3 | PR 병합됨(대상 브랜치는 `copilot/lnt-boil-off-rate-prediction`) |
| `copilot/fix-bor-model-sheet-names` | `271d68e` | #4 | PR 병합됨 |
| `copilot/fix-bor-calc-formulas` | `ec273f3` | #5 | PR 병합됨 |
| `copilot/fix-name-error-in-bor-calc-and-measure` | `2471be2` | #6 | PR 병합됨 |
| `copilot/update-tank-drawing-in-input-sheet` | `256091d` | #7 | PR 병합됨 |
| `copilot/prepare-handoff-to-claude-code` | 이 PR | #8 | 진행 중 — 인수인계 문서 작업 |

### 남은 불확실성

- 병합된 `copilot/*` 브랜치들은 삭제되지 않고 남아 있습니다. 병합 후 추가 커밋이 있는지,
  브랜치별 HEAD가 `main`에 완전히 포함되는지는 **이 환경의 얕은 클론(shallow clone)에서 확인하지 못했습니다.**
  정리 전에 `git fetch --all` 후 `git branch --merged main`으로 직접 검증하세요.
- `copilot/bor-prediction-excel-model`(PR #1)만 미병합입니다. 폐기 대상으로 보이지만
  **사용자 확인 없이 삭제하지 마세요.**

## 저장소 위생 상태

- `__pycache__/generate_bor_model.cpython-312.pyc`, `__pycache__/validate_model.cpython-312.pyc`가
  **git에 추적되고 있습니다.** 스크립트를 한 번만 실행해도 변경 상태로 나타납니다.
- `.gitignore` 파일이 **없습니다.**
- 생성 산출물 `LNG_BOR_Model.xlsx`가 저장소에 커밋되어 있습니다(의도된 것인지 **미확인**).
- Copilot 전용 지침 파일(`.github/copilot-instructions.md` 등)은 **존재하지 않습니다.**
  `.github/` 아래에는 `workflows/validate.yml`만 있습니다. 따라서 재사용할 Copilot 지침은 없고,
  프로젝트 규칙은 `README.md`와 코드에서 추론했습니다.
