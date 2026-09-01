# TODO — 남은 작업과 우선순위

> 아래 항목은 저장소 관찰에서 도출한 **제안**입니다. 사용자가 확정한 로드맵이 아니므로
> 작업 전에 우선순위를 사용자와 확인하세요. 완료된 작업으로 오해하지 마세요.

## P1 — 인수인계 직후 확인

- [ ] Claude Code에서 `@docs/handoff/*` import가 정상 인식되는지 확인
- [ ] 로컬에서 `pip install -r requirements.txt` → `python generate_bor_model.py`
      → `python validate_model.py` 재현 (한글 글꼴 설치 상태 포함)
- [ ] `git fetch --all` 후 `git branch --merged main`으로 `copilot/*` 브랜치 병합 여부 직접 검증

## P2 — 저장소 위생 (사용자 승인 필요)

- [ ] `__pycache__/*.pyc` 추적 해제(`git rm -r --cached __pycache__`) 및 `.gitignore` 추가
- [ ] 생성 산출물 `LNG_BOR_Model.xlsx`를 계속 커밋할지 결정
      (CI 아티팩트로도 제공되므로 중복일 수 있음 — **의도 미확인**)
- [ ] 미병합 브랜치 `copilot/bor-prediction-excel-model`(PR #1) 처리 방침 결정.
      **임의 삭제 금지**

## P3 — 모델·문서 품질

- [ ] 물성값(단열재 k, LNG 성분 물성)의 출처 문헌을 `README.md` 또는 `PropertyDB`에 명시
- [ ] `Measure` 시트 실측 비교 결과에 대한 판정 기준(오차 10%)을 워크북 안에서도 안내할지 검토
- [ ] LNG 조성 성분 수를 3개 초과로 확장할지 검토(현재 `INPUT_ROWS`의 `comp1~comp3` 고정)

## P4 — 개발 편의 (도입 전 사용자 확인 필수)

- [ ] 린터/포매터(예: `ruff`) 또는 테스트 프레임워크 도입 여부 결정.
      현재 저장소에는 어떤 린터·테스트 프레임워크도 없습니다.
- [ ] `validate_model.py` 실행 시 워크북이 재생성되어 git 상태가 더러워지는 문제 완화 방안 검토
      (예: 임시 경로에 생성 후 검증)

## 미확인 사항 (사용자에게 질문할 것)

- 이 모델이 적용될 실제 설비/프로젝트와 요구 정확도
- 실측 BOG 데이터의 보관 위치와 공유 가능 여부 (저장소에는 실측 데이터 없음)
- 향후 과도상태(transient)·일사량 반영 등 모델 확장 계획 유무
