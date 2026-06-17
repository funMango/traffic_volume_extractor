# 교통량 보정 방식

이 문서는 `traffic_api`의 현재 코드 기준 교통량 이상탐지 및 보정 흐름을 설명한다.
API 스키마, 보정 로직, 데이터베이스 접근 코드는 이 문서에서 변경하지 않는다.

## 요청 흐름

보정 요청은 Clean Architecture 계층을 지나 기존 이상탐지 구현으로 위임된다.

1. `traffic_api.presentation.routes`
   - `/corrected-traffic`, `/corrected-traffic-drct`, `/corrected-traffic-acsr`, `/corrected-traffic-crsrd`, `/corrected-traffic-vknd` 요청을 받는다.
   - 요청 검증은 `traffic_api.presentation.schemas`의 `JobRequest`, `RawTrafficVkndRequest`가 담당한다.
2. `traffic_api.application.traffic_use_cases.TrafficQueryUseCase`
   - Presentation 계층의 요청 객체를 받아 `TrafficAnalysisGateway` 인터페이스로 전달한다.
   - UseCase는 구체 구현체를 직접 생성하지 않는다.
3. `traffic_api.infrastructure.legacy_analysis_gateway.LegacyTrafficAnalysisGateway`
   - 기존 `api_server.py` 함수 호출을 감싼 Infrastructure 어댑터다.
   - 서버 초기화 시 휴일 정보와 DRCT 방향 존재 캐시를 준비한다.
4. `api_server.py`
   - API 응답 형태를 만든다.
   - 교차로명, 접근로명, 방향명, 차종명 같은 메타데이터를 결합한다.
   - 정상 슬롯과 이상 슬롯을 함께 응답 슬롯으로 직렬화한다.
5. `이상탐지.py` / `anomaly_core.py`
   - `이상탐지.py`는 DB 조회와 기존 실행 흐름을 담당한다.
   - `anomaly_core.py`는 기준선 생성, 이상치 판정, 보정값 산정 같은 순수 계산 로직을 담당한다.

## 기준선 생성

기준선은 대상 기간의 각 슬롯을 비교할 기준 통계다.

- 기준연도는 대상연도를 제외하고 `AVAILABLE_YEARS`에서 가까운 과거 1개와 가까운 미래 1개를 우선 선택한다.
- 기준 데이터가 부족한 셀은 fallback 연도를 추가로 사용한다. 현재 규칙은 대상연도가 `DATA_START_YEAR`보다 크면 전년도, 아니면 다음 연도다.
- 휴일과 요일은 `get_day_type()`으로 구분한다.
  - `0`: 평일
  - `1`: 토요일
  - `2`: 일요일 또는 공휴일
- ACSR/DRCT 기준 셀은 접근로 또는 방향, 요일유형, 시간, 월 단위로 구성된다.
- VKND 기준 셀은 접근로, 차종, 요일유형, 시간, 15분 분, 월 단위로 구성된다.
- 각 셀은 `median`, `q1`, `q3`, `n_clean`, `zero_rate`를 가진다.
- `n_clean`은 낮은 이상값을 최대 3회 제거한 뒤 남은 정상 표본 수다.
- `zero_rate`는 기준연도에서 기대되는 슬롯 수 대비 실제 수집일이 없는 비율이다.
- `zero_rate >= 0.8`이면 기준 범위를 확장한다.
  - 1단계: 전월, 당월, 익월
  - 2단계: 연중 전체 월

## 이상치 판정

판정은 대상 기간의 슬롯별로 수행된다.

- 결측이고 `zero_rate < 0.2`이면 `A형`이다.
- 결측이고 `0.2 <= zero_rate < 0.8`이면 `B형`이다.
- 결측이고 `zero_rate >= 0.8`이면 평소에도 수집이 드문 셀로 보고 정상 취급한다.
- 값이 존재하고 기준 통계가 있으면 다음 조건 중 하나를 만족할 때 `B형`이다.
  - `value / median < 0.5`
  - `value < q1 - 1.5 * IQR`
- 그 외 슬롯은 정상이다.

## 보정값 산정

이상 슬롯에 대해서만 보정값을 산정한다.

- 같은 단위의 전후 정상값을 찾는다.
  - ACSR: 같은 접근로, 같은 시간
  - DRCT: 같은 접근로 방향, 같은 시간
  - VKND: 같은 접근로, 같은 차종, 같은 시간, 같은 15분 분
- 전후 정상값이 모두 있고 거리 합이 14일 이내이면 선형보간을 사용한다.
- 선형보간이 불가하고 기준 통계가 있으면 기준선 `median`을 사용한다.
- 보정값은 `round()`로 반올림한다.
- 정상 슬롯은 `corrected_value`가 `null`이며, 집계 시 원시값을 사용한다.
- 신뢰도는 다음 중 하나라도 해당하면 `LOW`다.
  - `n_clean < 8`
  - fallback 표본이 사용됨
  - 선형보간 거리 합이 7일 초과
- 위 조건에 해당하지 않는 보정값은 `OK`다.

## 집계 방식

### ACSR 기본 보정

`POST /corrected-traffic`은 `analyse_node()` 결과를 기반으로 접근로 단위 `slots`를 반환한다.

- 이상 슬롯은 `anomaly_type`, `corrected_value`, `correction_method`, `confidence`, `baseline`을 포함한다.
- 정상 슬롯은 `anomaly_type`과 `corrected_value`가 `null`이다.
- 정상 슬롯의 집계값은 원시 `traffic_volume`을 그대로 사용한다.

### DRCT 보정과 ACSR/교차로 집계

`POST /corrected-traffic-drct`는 방향 단위 보정 결과와 접근로 집계를 함께 반환한다.

- `drct_slots`가 먼저 생성된다.
- `drct_cd=00`은 응답 집계 대상에서 제외된다.
- `slots`는 DRCT 슬롯을 접근로 단위로 합산한 결과다.
- `POST /corrected-traffic-acsr`는 DRCT 기반 접근로 집계 `slots`만 반환한다.
- `POST /corrected-traffic-crsrd`는 접근로 집계를 다시 교차로 단위로 합산한 `slots`를 반환한다.
- 집계 중 `corrected_value`가 `null`인 정상 슬롯은 원시값을 보정 후 값처럼 사용해 합산한다.
- 하위 슬롯에 `A형`과 `B형`이 함께 있으면 상위 집계의 `anomaly_type`은 `A+B혼합`이다.

### VKND 차종 보정

`POST /corrected-traffic-vknd`는 차종 단위 보정 결과를 반환한다.

- 원천 보정 단위는 15분이다.
- `interval=15m`이면 15분 단위 `vknd_slots`를 그대로 반환한다.
- `interval=1h`이면 15분 보정 결과를 1시간 단위로 합산해 `vknd_slots`를 만든다.
- 응답은 다음 세 묶음을 반환한다.
  - `vknd_slots`: 접근로, 차종, 시간 단위 상세
  - `slots`: 접근로와 차종 단위 집계
  - `node_slots`: 교차로와 차종 단위 집계
- 1시간 집계도 정상 슬롯의 `corrected_value == null`은 원시값으로 합산한다.
- 15분 하위 슬롯 중 하나라도 `LOW`이면 1시간 집계 신뢰도도 `LOW`다.

## 코드 기준 위치

- API 라우터: `01_Program/09_else/traffic_api/presentation/routes.py`
- 요청 스키마: `01_Program/09_else/traffic_api/presentation/schemas.py`
- UseCase: `01_Program/09_else/traffic_api/application/traffic_use_cases.py`
- Legacy Gateway: `01_Program/09_else/traffic_api/infrastructure/legacy_analysis_gateway.py`
- API 응답 조립: `01_Program/09_else/api_server.py`
- 기존 분석 오케스트레이션: `01_Program/09_else/이상탐지.py`
- 순수 보정 계산: `01_Program/09_else/anomaly_core.py`
- 집계 도메인 함수: `01_Program/09_else/traffic_api/domain/aggregation.py`
