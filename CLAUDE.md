@AGENTS.md

# Claude Code 전용 설정

- 훅은 `.claude/settings.json`에 정의되어 있습니다. 주민등록번호·카드번호 형식이 포함된 파일 쓰기를 차단합니다.
- 커스텀 명령:
  - `/start-service` — 새 과제를 시작하고 `docs/00`부터 순서대로 안내받습니다.
  - `/gate-check` — 게이트 상태와 문서 누락을 점검하고 결과를 요약합니다.
- 작업 후에는 `node scripts/check-gates.mjs`와 `node scripts/check-traceability.mjs`를 실행해 결과를 확인합니다.
- 이 파일은 Claude Code 전용 설정만 담습니다. 공통 규칙은 모두 `AGENTS.md`에 있습니다. 공통 규칙을 고칠 때는 `AGENTS.md`만 수정합니다.
