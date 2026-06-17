# 아키텍처 규칙

## 기준

- Clean Architecture 기준으로 작성한다.
- 계층은 Presentation, Application, Domain, Infrastructure로 나눈다.
- 의존성은 항상 Domain 방향으로 향한다.
- Domain은 외부 계층, 프레임워크, DB, API에 의존하지 않는다.

## 설계 원칙

- 단일 책임 원칙을 지킨다.
- 하나의 클래스, 함수, 파일은 하나의 명확한 책임만 가진다.
- UseCase는 구체 구현체가 아니라 Interface에 의존한다.
- 구현체는 Infrastructure 계층에 둔다.

## 금지

- Controller에 비즈니스 로직을 작성하지 않는다.
- Domain에서 DB, API, 파일 시스템에 직접 접근하지 않는다.
- UseCase에서 구체 구현체를 직접 생성하지 않는다.
- `Service`, `Manager`, `Util` 같은 넓은 이름은 남용하지 않는다.
