# Git 작업 규칙

## 브랜치 전략

- 기본 전략은 `feature -> develop -> main`입니다.
- 허용 브랜치명은 다음과 같습니다.
  - `develop`
  - `main`
  - `feature/<name>`
  - `fix/<name>`
  - `hotfix/<name>`
- `main` 직접 커밋과 직접 push는 지양하고 PR 병합으로 반영합니다.

## 커밋 메시지

- Conventional Commit 형식을 사용합니다.
- 형식: `<type>(optional-scope): <subject>`
- 허용 type은 다음과 같습니다.
  - `feat`
  - `fix`
  - `refactor`
  - `test`
  - `docs`
  - `chore`
  - `perf`
- 예시:
  - `feat(api): add job progress endpoint`
  - `fix(excel): sort unified sheet by time and name`
  - `test(anomaly): add edge-case coverage`

## 코드 품질 검사

- Python 품질 검사는 `ruff + pytest`를 사용합니다.
- 표준 명령:
  - `.\.venv\Scripts\python -m ruff format 01_Program scripts`
  - `.\.venv\Scripts\python -m ruff check 01_Program scripts`
  - `.\.venv\Scripts\python -m pytest 01_Program --ignore=01_Program/test_db.py`
- `test_db.py`는 실제 DB 연결 테스트이므로 기본 자동 검사에서 제외합니다.
- 문서만 변경한 작업은 Python 테스트를 생략할 수 있습니다.

## pre-commit / pre-push

- `.pre-commit-config.yaml` 기준으로 다음을 검사합니다.
  - pre-commit: `ruff format`, `ruff check`
  - commit-msg: Conventional Commit 형식 검증
  - pre-push: 브랜치명 정책, `pytest`
- 설치:
  - `.\.venv\Scripts\python -m pre_commit install`
  - `.\.venv\Scripts\python -m pre_commit install --hook-type commit-msg`
  - `.\.venv\Scripts\python -m pre_commit install --hook-type pre-push`

## 산출물(`02_Result`) 정책

- 원칙적으로 `02_Result`는 생성 산출물 폴더로 간주하고 기본 미추적합니다.
- 예외가 필요하면 `.gitkeep` 또는 명시적 예외 패턴으로만 추적합니다.
- 이미 추적 중인 기존 산출물은 이 정책 도입 후 별도 정리 커밋에서 제거합니다.

## 작업 완료 후 원격 반영

- 사용자 요청 작업이 완료되면 변경 범위를 확인합니다.
- 이번 작업과 관련 없는 기존 변경은 커밋에 포함하지 않습니다.
- 관련 변경만 stage하고 `git diff --cached --stat`로 staged 범위를 확인합니다.
- 커밋 메시지는 Conventional Commit 규칙을 따릅니다.
- 커밋 후 현재 브랜치의 원격 추적 브랜치로 push합니다.
