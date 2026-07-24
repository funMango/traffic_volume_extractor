"""산업길사거리 차종별 1시간 교통량 원천을 읽기 전용으로 추출한다."""

from __future__ import annotations

import argparse
import json
import os
import re
from datetime import datetime
from pathlib import Path

import oracledb
from dotenv import load_dotenv


PROJECT_DIR = Path(__file__).resolve().parents[2]
ENV_PATH = PROJECT_DIR / "00_Data" / ".env"
TRAFFIC_TABLE = "S_CRSRD_VKND_TRF_1HH"
NODE_ID = "BC001N0114"
START_AT = datetime(2026, 7, 13, 8, 0, 0)
END_AT = datetime(2026, 7, 13, 10, 0, 0)
APPROACHES = (
    ("동(서향)", "ACSR000001"),
    ("서(동향)", "ACSR000002"),
    ("북(남향)", "ACSR000004"),
    ("남(북향)", "ACSR000003"),
)
VEHICLES = (
    ("승용차", "2"),
    ("소형트럭", "10000"),
    ("중형버스", "80000"),
    ("대형버스", "20"),
    ("대형트럭", "20000"),
)
BLOCKED_SQL_PATTERN = re.compile(
    r"\b(INSERT|UPDATE|DELETE|MERGE|DROP|TRUNCATE|ALTER|CREATE|COMMIT|ROLLBACK|GRANT|REVOKE)\b",
    re.IGNORECASE,
)


def connect_read_only():
    load_dotenv(ENV_PATH)
    required = ("DB_USER", "DB_PASSWORD", "DB_HOST", "DB_SERVICE")
    missing = [name for name in required if not os.getenv(name)]
    if missing:
        raise RuntimeError(f"DB 접속 환경변수가 없습니다: {', '.join(missing)}")

    connection = oracledb.connect(
        user=os.environ["DB_USER"].strip(),
        password=os.environ["DB_PASSWORD"].strip(),
        host=os.environ["DB_HOST"].strip(),
        port=int(os.getenv("DB_PORT", "1521").strip()),
        service_name=os.environ["DB_SERVICE"].strip(),
    )
    with connection.cursor() as cursor:
        cursor.execute("SET TRANSACTION READ ONLY")
    return connection


def ensure_select_sql(sql: str) -> None:
    normalized = re.sub(r"--[^\r\n]*", " ", sql).strip()
    if not normalized.upper().startswith("SELECT"):
        raise RuntimeError("읽기 전용 보호: SELECT 문만 실행할 수 있습니다.")
    if BLOCKED_SQL_PATTERN.search(normalized):
        raise RuntimeError("읽기 전용 보호: 쓰기 또는 DDL SQL은 실행할 수 없습니다.")


def extract_rows(connection) -> list[dict[str, object]]:
    sql = f"""
        SELECT
            v.ACSR_ID,
            TRIM(TO_CHAR(v.VKND_CD)) AS VKND_CD,
            TO_NUMBER(TO_CHAR(v.TOT_DT, 'HH24')) AS HOUR,
            NVL(SUM(v.TRF_QNTY), 0) AS TRAFFIC_VOLUME
        FROM {TRAFFIC_TABLE} v
        WHERE v.NODE_ID = :node_id
          AND v.TOT_DT >= :start_at
          AND v.TOT_DT < :end_at
          AND v.ACSR_ID IN (:approach_0, :approach_1, :approach_2, :approach_3)
          AND TRIM(TO_CHAR(v.VKND_CD)) IN (:vehicle_0, :vehicle_1, :vehicle_2, :vehicle_3, :vehicle_4)
        GROUP BY
            v.ACSR_ID,
            TRIM(TO_CHAR(v.VKND_CD)),
            TO_NUMBER(TO_CHAR(v.TOT_DT, 'HH24'))
        ORDER BY v.ACSR_ID, VKND_CD, HOUR
    """
    ensure_select_sql(sql)
    bind_values = {
        "node_id": NODE_ID,
        "start_at": START_AT,
        "end_at": END_AT,
        **{f"approach_{index}": code for index, (_, code) in enumerate(APPROACHES)},
        **{f"vehicle_{index}": code for index, (_, code) in enumerate(VEHICLES)},
    }
    with connection.cursor() as cursor:
        cursor.execute(sql, bind_values)
        return [
            {
                "approach_id": str(approach_id).strip(),
                "vehicle_code": str(vehicle_code).strip(),
                "hour": int(hour),
                "traffic_volume": int(traffic_volume or 0),
            }
            for approach_id, vehicle_code, hour, traffic_volume in cursor.fetchall()
        ]


def build_payload(rows: list[dict[str, object]]) -> dict[str, object]:
    return {
        "intersection": "산업길사거리",
        "source_table": TRAFFIC_TABLE,
        "query_condition": (
            "NODE_ID=BC001N0114, TOT_DT>=2026-07-13 08:00:00, "
            "TOT_DT<2026-07-13 10:00:00, 접근로 4개, 차종 5개"
        ),
        "periods": [
            {"hour": 8, "label": "7월 13일_8~9시"},
            {"hour": 9, "label": "7월 13일_9~10시"},
        ],
        "approaches": [{"name": name, "id": approach_id} for name, approach_id in APPROACHES],
        "vehicles": [{"name": name, "code": code} for name, code in VEHICLES],
        "rows": rows,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="JSON 저장 경로")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    connection = connect_read_only()
    try:
        rows = extract_rows(connection)
    finally:
        connection.close()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(build_payload(rows), ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"원천 행 수: {len(rows)}")
    print(f"저장 완료: {args.output}")


if __name__ == "__main__":
    main()
