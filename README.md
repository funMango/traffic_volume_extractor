# 교통량 추출 시스템

Python 기반 교통량 추출/이상탐지 프로젝트입니다.  
이 저장소는 `feature -> develop -> main` 브랜치 전략과 자동 품질검사(ruff + pytest)를 기본으로 사용합니다.

## 빠른 시작

```powershell
# 프로젝트 루트 기준
python -m venv .venv
.\.venv\Scripts\python -m pip install --upgrade pip
.\.venv\Scripts\python -m pip install pre-commit ruff pytest fastapi pydantic numpy openpyxl python-dotenv plotly tqdm oracledb
```

## 표준 검사 명령

```powershell
.\.venv\Scripts\python -m ruff format 01_Program scripts
.\.venv\Scripts\python -m ruff check 01_Program scripts
.\.venv\Scripts\python -m pytest 01_Program --ignore=01_Program/test_db.py
```

## Git 규칙 요약

- 브랜치:
  - 허용: `develop`, `main`, `feature/*`, `fix/*`, `hotfix/*`
- 머지 흐름:
  - 기능/수정 브랜치 -> `develop`
  - 릴리스 시 `develop` -> `main`
- 커밋 메시지:
  - `Conventional Commits` 형식 사용
  - 예: `feat(api): add daily summary endpoint`

## pre-commit 설치

```powershell
.\.venv\Scripts\python -m pre_commit install
.\.venv\Scripts\python -m pre_commit install --hook-type commit-msg
.\.venv\Scripts\python -m pre_commit install --hook-type pre-push
```

상세 운영 규칙은 `CONTRIBUTING.md`를 참고하세요.
