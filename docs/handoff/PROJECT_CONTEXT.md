# PROJECT_CONTEXT — 프로젝트 배경

> 근거: `README.md`, `generate_bor_model.py`, `validate_model.py`, `.github/workflows/validate.yml`,
> GitHub PR #1~#8 제목/상태. 근거가 없는 항목은 "미확인"으로 표기했습니다.

## 무엇을 만드는가

LNG 극저온 저장 탱크의 **BOR(Boil-Off Rate, %/day)** 를 설계 단계에서 예측하기 위한
**엑셀 워크북 자동 생성기**입니다. 최종 사용자는 엑셀에서 입력값만 바꿔가며 BOR을 확인합니다.

- 산출물: `LNG_BOR_Model.xlsx` (시트 4개: `Input`, `PropertyDB`, `BOR_Calc`, `Measure`)
- 생성 도구: `generate_bor_model.py` (openpyxl로 워크북, matplotlib로 탱크 도면 PNG 생성)
- 검증 도구: `validate_model.py` + GitHub Actions 워크플로 `Validate LNG BOR Model`

## 왜 엑셀인가 (관찰된 설계 의도)

- 워크북에는 값이 아니라 **엑셀 수식**(`VLOOKUP`, `PI()` 등)이 저장되어, Python 없이도
  현장에서 입력값을 바꿔 재계산할 수 있습니다.
- `validate_model.py`가 시트 XML의 `<f>` 태그 개수를 검사하는 것도 "수식이 살아 있어야 한다"는
  요구사항을 지키기 위한 장치입니다.

## 모델 개요 (README.md에 기재된 관계식 그대로)

```
Q [W] = ΔT / R_total
        ΔT = 외부온도 - LNG 평균비등점
        R_total = 단열재두께 / (keff × 탱크표면적 A)

BOG 질량유량 [kg/s] = Q / (LNG 평균잠열 ΔHvap [J/kg])
BOG 일일증발량 [kg/day] = BOG유량 × 86400 × 압력보정계수
LNG 총 질량 [kg] = LNG 체적 × LNG 평균밀도
BOR [%/day] = (BOG 일일증발량 / LNG 총 질량) × 100
```

보정계수(코드/README 기준):

- 압력보정계수 `= MAX(0.8, 1 − (P−1.013)×0.05)`
- 진공보정계수: `< 10 Pa → 0.7`, `< 100 Pa → 0.85`, 대기압 → `1.0`
- 배관 보냉계수: 보냉없음 `1.0`, 일반보냉 `0.4`, 진공배관 `0.1`
- 안정화 시간 보정: `< 6h → 1.15`, `< 12h → 1.05`, `≥ 12h → 1.0`

물성 DB(단열재 k값, LNG 성분 물성) 전체 표는 `README.md`에 있으므로 여기서 중복 기재하지 않습니다.

## 데이터 입력/출력

- 입력: 코드에 내장된 기본값 + 사용자가 엑셀 `Input` 시트에서 수정하는 값
  (탱크 내경·길이·충전율·압력·외기온도·단열재/두께·진공도·배관조건·LNG 조성 3성분·계측 시간)
- 외부 데이터 파일, 데이터베이스, API 연동은 **없습니다.**
- 출력: `LNG_BOR_Model.xlsx` (저장소에도 커밋되어 있고 CI 아티팩트로도 업로드)

## 이해관계자 / 사용 맥락

- 저장소 소유자 `@Doo-mo`가 단독 작성. 지금까지의 변경은 모두 GitHub Copilot 코딩 에이전트 PR로 진행되었습니다.
- 실제 현장 적용 대상 설비·프로젝트명은 저장소에 근거가 없습니다. **미확인**이며 추측하지 마세요.

## 알려진 한계 (README.md 기재)

- 정상상태 열전달 기반의 **설계 예측용** 모델. 일사량, 지면 열전도, 과도현상은 미반영.
- 예측·실측 BOR 오차가 10% 이상이면 입력 조건 재확인 권고.
- LNG 성분 물성값은 대표 문헌값.
