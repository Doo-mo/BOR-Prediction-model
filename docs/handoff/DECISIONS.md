# DECISIONS — 지금까지의 결정과 근거

> 각 항목은 PR 제목/병합 상태, 코드, README에서 확인한 사실만 기록합니다.
> 결정의 배경 대화가 남아 있지 않은 항목은 "배경 미확인"으로 표시합니다.

## D1. BOR = Boil-Off Rate (병상이용률 아님)

- 근거: PR #1 "Add BOR prediction Excel model (병상이용률 예측)"은 **닫히고 병합되지 않았고**,
  이후 PR #2 "Add LNG Boil-Off Rate (BOR) prediction Excel model"이 병합되었습니다.
- 현재 `README.md`와 `generate_bor_model.py` docstring에 "병상이용률과 전혀 관계없다"고 명시되어 있습니다.
- 결론: 이 저장소의 BOR은 항상 LNG 증발률입니다.

## D2. 결과는 값이 아니라 엑셀 수식으로 저장

- 근거: `validate_model.py`가 시트 XML의 `<f>` 태그 개수를 검사하고,
  `BOR_Calc` 수식의 열 참조를 정규화 비교합니다.
- 결론: openpyxl로 계산 결과값만 써넣는 방식으로 바꾸지 마세요.

## D3. 시트 탭 이름은 ASCII(`Input` / `PropertyDB` / `BOR_Calc` / `Measure`)

- 근거: PR #4 "Fix Excel formula removal: rename sheet tabs to ASCII
  (Input/PropertyDB/BOR_Calc/Measure)" 병합.
- 이유: 한글 탭명 사용 시 수식이 제거되는 문제가 있었습니다(PR 제목 근거).
- 시트 내부 라벨·설명은 계속 한글을 사용합니다(`README.md` 참고 문구).

## D4. LNG 조성 수식과 CI 검증 워크플로 정비

- 근거: PR #3 "Clarify BOR result labeling and add automated workbook validation (script + CI)",
  PR #5 "Fix BOR_Calc LNG composition formulas and restore workbook validation workflow" 병합.
- 결론: `validate_model.py`와 `.github/workflows/validate.yml`은 회귀 방지 장치이므로 유지해야 합니다.

## D5. `Input` 행 번호 단일 관리 + 셀 연동 탱크 사양 패널

- 근거: PR #6 "Fix D-column `#NAME?` notes, add Input auto-volume rows, and rebase cross-sheet refs
  with linked tank spec panel" 병합. 코드의 `INPUT_ROWS` 딕셔너리 주석
  ("시트 간 참조 불일치를 방지한다").
- 결론: 행 번호를 코드 여러 곳에 하드코딩하지 말고 `INPUT_ROWS`를 통해 참조하세요.
- D열 메모 문자열이 `=`로 시작하면 엑셀에서 `#NAME?`가 되므로,
  `validate_model.py`가 이를 회귀 검사합니다.

## D6. 탱크 도면은 matplotlib으로 생성해 삽입, 수치는 셀에 표시

- 근거: PR #7 "Copilot/update tank drawing in input sheet" 병합, `render_tank_images()`,
  `TANK_ANNOTATIONS = ["내부체적", "내경", "길이", "단열두께", "진공도", "배관조건"]`,
  README의 "도면에는 … 명칭만 지시선으로 표기하며, 실제 수치는 옆의 셀 연동 사양 패널에 표시" 문구.
- 도면 PNG는 `tempfile.TemporaryDirectory()`에 만들고 저장소에 남기지 않습니다.
- 한글 라벨 출력을 위해 CI에서 `fonts-nanum`을 설치합니다.

## D7. 물성값은 문헌 대표값 사용

- 근거: `README.md`의 단열재 k값 / LNG 성분 물성 / 배관 보냉계수 표와 "대표 문헌값" 문구.
- 출처 문헌명은 저장소에 기록되어 있지 **않습니다(배경 미확인)**. 값 변경 시 근거를 함께 남기세요.

## D8. Copilot → Claude 인수인계 문서화 (이 PR)

- 근거: 사용자 요청 "지금 내 깃허브 레포지토리와 채팅 내용들을 클로드로 옮길거야".
- 범위: `CLAUDE.md` + `docs/handoff/*` 문서만 추가하며 모델 코드는 변경하지 않습니다.
