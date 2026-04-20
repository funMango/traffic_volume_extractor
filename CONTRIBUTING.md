# CONTRIBUTING

## 1) 브랜치 전략

- 기본 전략: `feature -> develop -> main`
- 허용 브랜치명:
  - `develop`
  - `main`
  - `feature/<name>`
  - `fix/<name>`
  - `hotfix/<name>`
- `main` 직접 커밋/푸시는 지양하고, PR 병합으로만 반영합니다.

## 2) 커밋 메시지 규칙

- 형식: `<type>(optional-scope): <subject>`
- 허용 type:
  - `feat`, `fix`, `refactor`, `test`, `docs`, `chore`, `perf`
- 예시:
  - `feat(api): add job progress endpoint`
  - `fix(excel): sort unified sheet by time and name`
  - `test(anomaly): add edge-case coverage`

## 3) 코드 품질 규칙

- Python 품질검사는 `ruff + pytest`를 사용합니다.
- 표준 명령:
  - `.\.venv\Scripts\python -m ruff format 01_Program scripts`
  - `.\.venv\Scripts\python -m ruff check 01_Program scripts`
  - `.\.venv\Scripts\python -m pytest 01_Program --ignore=01_Program/test_db.py`
- `test_db.py`는 실제 DB 연결 테스트이므로 기본 자동검사에서 제외합니다.

## 4) pre-commit / pre-push

- `.pre-commit-config.yaml` 기준으로 다음을 검사합니다.
  - pre-commit: `ruff format`, `ruff check`
  - commit-msg: Conventional Commit 형식 검증
  - pre-push: 브랜치명 정책 + `pytest`
- 설치:
  - `.\.venv\Scripts\python -m pre_commit install`
  - `.\.venv\Scripts\python -m pre_commit install --hook-type commit-msg`
  - `.\.venv\Scripts\python -m pre_commit install --hook-type pre-push`

## 5) 산출물(`02_Result`) 정책

- 원칙: `02_Result`는 생성 산출물 폴더로 간주하고 기본 미추적합니다.
- 예외가 필요하면 `.gitkeep` 또는 명시적 예외 패턴으로만 추적합니다.
- 이미 추적 중인 기존 산출물은 이 정책 도입 후 별도 정리 커밋에서 제거합니다.
