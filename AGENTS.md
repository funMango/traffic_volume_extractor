# AGENTS.md

## 문서 안내

이 파일은 프로젝트 작업 문서의 길잡이입니다. 실제 규칙과 매뉴얼은 `docs` 아래 문서를 기준으로 확인합니다.

- `docs/rules/database.md`: `.env`, 데이터베이스 접근 권한, 읽기 전용 원칙, 금지 작업
- `docs/rules/terminal.md`: 파일 읽기, 파일 목록 확인, 단일 파일 검색, Python 실행, 파일 수정 방식
- `docs/rules/architecture.md`: Clean Architecture 계층, 의존성 방향, UseCase 의존 원칙, Domain 격리 원칙
- `docs/rules/git.md`: 브랜치, 커밋 메시지, 품질 검사, hook, 산출물, 원격 반영 기준
- `docs/log`: 월 단위 명령 수행 기록
- `docs/manual`: `01_Program` 서비스 사용법

## 작업 기준

- 작업 전 관련 `docs/rules` 문서를 확인합니다.
- 서비스 실행과 사용법은 `docs/manual` 문서를 따릅니다.
- Git 작업 및 원격 반영은 `docs/rules/git.md`를 따릅니다.
- 명령이 완료되면 `docs/log`의 해당 월 로그에 기록합니다.
- 로그를 쓰는 행위 자체는 중복 방지를 위해 별도로 기록하지 않습니다.
- 작업 완료 후 관련 변경만 커밋하고 현재 브랜치의 원격 추적 브랜치로 push합니다.
