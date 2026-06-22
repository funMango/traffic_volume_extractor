# traffic_api 사용법

## 서버 실행

`01_Program/09_else/start_api_server.cmd`를 실행한다.

```bat
cmd /c "C:\01_Project\09_교통량 추출 시스템\01_Program\09_else\start_api_server.cmd"
```

이 스크립트는 `traffic_api.main` 기반 FastAPI 서버를 실행한다. 기본 포트는 `8000`이며, 이미 `8000` 포트가 LISTENING 상태이면 중복 실행하지 않고 종료한다.

## 실행 구조

- 진입점: `01_Program/09_else/traffic_api/main.py`
- FastAPI 앱 팩토리: `create_app()`
- 라우터: `traffic_api.presentation.routes.router`
- 기본 실행 방식: `python -m traffic_api.main`
- 기본 바인딩: `0.0.0.0:8000`

## 주요 API 경로

현재 `traffic_api.presentation.routes`에서 확인된 주요 경로는 다음과 같다.

- `POST /jobs`: 교통량 처리 작업을 생성한다.
- `GET /jobs/{job_id}`: 생성된 작업 상태 또는 결과를 조회한다.
- `POST /corrected-traffic`: 보정 교통량 결과를 조회한다.
- `POST /corrected-traffic-acsr`: DRCT 5분 원천 기반 접근로 기준 보정 교통량 결과를 조회한다.
- `POST /corrected-traffic-crsrd`: DRCT 5분 원천 기반 교차로 기준 보정 교통량 결과를 조회한다.
- `POST /corrected-traffic-drct`: DRCT 5분 원천 기반 방향 기준 보정 교통량 결과를 조회한다.
- `POST /corrected-traffic-vknd`: 차종 기준 보정 교통량 결과를 조회한다.
- `GET /vehicle-kinds`: 차종 기준 목록을 조회한다.
- `POST /anomaly-daily-summary`: 이상탐지 일별 요약을 조회한다.
- `GET /intersections`: 교차로 목록을 조회한다. `refresh=true` 쿼리로 갱신 조회할 수 있다.

## 기본 요청 형식

대부분의 `POST` 경로는 다음 필드를 받는다.

```json
{
  "node_ids": ["12345"],
  "date_start": "2026-06-01",
  "date_end": "2026-06-17",
  "hours_preset": "all"
}
```

- `node_ids`: 1개 이상의 노드 ID
- `date_start`, `date_end`: `YYYY-MM-DD` 형식
- `hours`: 선택 필드, 0부터 23까지의 정수 배열
- `hours_preset`: 선택 필드, `all` 또는 `peak`
- `hours`와 `hours_preset`은 동시에 지정하지 않는다.

`POST /corrected-traffic-drct`, `POST /corrected-traffic-acsr`, `POST /corrected-traffic-crsrd`는 DRCT 계열 보정 API다.

- 원천 데이터는 `S_CRSRD_DRCT_TRF_5MI`의 DRCT 5분 슬롯이다.
- 보정은 5분 단위로 수행하며, 값이 없는 슬롯만 보정한다.
- 수집값이 있는 슬롯은 원시값 사용 후 상위 집계에 포함한다.
- `interval`: `5m`, `1h`, `1d` 중 하나를 지정할 수 있으며, 기본값은 `1h`다.
- `interval=5m`이면 5분 보정 슬롯을 반환하고, `interval=1h` 또는 `interval=1d`이면 5분 보정 결과를 각각 1시간 또는 1일 단위로 합산한다.
- `POST /corrected-traffic-acsr`와 `POST /corrected-traffic-crsrd`도 같은 DRCT 5분 보정 결과를 접근로 또는 교차로 단위로 집계한다.

`POST /corrected-traffic-vknd`는 추가로 다음 필드를 사용할 수 있다.

- `interval`: `15m` 또는 `1h`, 기본값은 `1h`
- `vknd_codes`: 선택 필드, 차종 코드 배열
