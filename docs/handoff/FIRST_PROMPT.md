# FIRST_PROMPT — Claude에 붙여 넣을 첫 프롬프트

아래 블록을 그대로 복사해 Claude Code(또는 Claude)에 첫 메시지로 붙여 넣으세요.

```text
이 저장소는 GitHub Copilot에서 Claude로 인수인계된 프로젝트입니다.
저장소: Doo-mo/BOR-Prediction-model (LNG 탱크 BOR = Boil-Off Rate 예측 엑셀 모델 생성기)

작업을 시작하기 전에 다음 순서로 진행해 주세요.

1. 저장소 루트의 CLAUDE.md를 먼저 읽고, 이어서 docs/handoff/ 아래 문서를 모두 읽으세요.
   - docs/handoff/PROJECT_CONTEXT.md
   - docs/handoff/CURRENT_STATUS.md
   - docs/handoff/DECISIONS.md
   - docs/handoff/TODO.md
2. 실제 코드를 직접 확인하세요: generate_bor_model.py, validate_model.py,
   requirements.txt, README.md, .github/workflows/validate.yml
3. 문서에 적힌 내용이 현재 코드와 일치하는지 검증하세요. 특히 다음을 실제로 실행해 확인해 주세요.
   pip install -r requirements.txt
   python generate_bor_model.py
   python validate_model.py
   (validate_model.py는 LNG_BOR_Model.xlsx를 재생성하므로, 코드 변경이 없다면
    실행 후 git checkout LNG_BOR_Model.xlsx 로 되돌려 주세요.)
4. 문서와 실제 상태가 다른 부분, 그리고 "미확인"으로 표시된 항목 중 나에게 물어봐야 할 질문을
   목록으로 정리해 주세요.
5. 그다음 앞으로의 작업 계획을 우선순위와 함께 제안해 주세요.

주의사항:
- 5번 계획에 내가 동의하기 전에는 코드나 파일을 수정하지 마세요.
- 모델 가정, 수식, 물성값, 명령어를 추측으로 만들어내지 마세요. 근거가 없으면 "미확인"이라고 말해 주세요.
- API 키, 비밀번호, 사내 주소 등 비밀 정보를 문서나 코드에 넣지 마세요.
- copilot/* 브랜치를 임의로 삭제하거나 병합하지 마세요.
- 답변은 한국어로 해 주세요.
```

## 참고

- 문서 링크: [PROJECT_CONTEXT](PROJECT_CONTEXT.md) · [CURRENT_STATUS](CURRENT_STATUS.md) ·
  [DECISIONS](DECISIONS.md) · [TODO](TODO.md) · [CLAUDE.md](../../CLAUDE.md)
