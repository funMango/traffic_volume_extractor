"""Export 15-minute vehicle-class traffic for Bakchon Bridge three-way intersection.

The extractor is intentionally limited to read-only Oracle access.  It resolves
the intersection and approach from master data, aggregates every vehicle class
in the requested time slot, and writes an auditable Excel workbook.
"""

from __future__ import annotations

import argparse
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import oracledb
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


PROJECT_DIR = Path(__file__).resolve().parents[2]
ENV_PATH = PROJECT_DIR / "00_Data" / ".env"
DEFAULT_OUTPUT_PATH = (
    PROJECT_DIR / "02_Result" / "박촌교삼거리_서_동향_20260713_1000_1015_차종별_교통량.xlsx"
)

INTERSECTION_NAME = "박촌교삼거리"
APPROACH_NAME = "서 (동향)"
START_AT = datetime(2026, 7, 13, 10, 0, 0)
END_AT = datetime(2026, 7, 13, 10, 15, 0)

INTERSECTION_TABLE = "M_CRSRD_INF"
APPROACH_TABLE = "M_CRSRD_ACSR_INF"
VEHICLE_CODE_TABLE = "M_CD_INF"
TRAFFIC_TABLE = "S_CRSRD_VKND_TRF_15MI"

BLOCKED_SQL_PATTERN = re.compile(
    r"\b(INSERT|UPDATE|DELETE|MERGE|DROP|TRUNCATE|ALTER|CREATE|COMMIT|ROLLBACK|GRANT|REVOKE)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ResolvedTarget:
    """Master-data identifiers required for the traffic query."""

    node_id: str
    intersection_name: str
    approach_id: str
    approach_name: str


@dataclass(frozen=True)
class VehicleTraffic:
    """One vehicle class total for the requested 15-minute slot."""

    vehicle_code: str
    vehicle_name: str
    traffic_volume: int


def _normalize_name(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def _strip_sql_comments(sql: str) -> str:
    without_blocks = re.sub(r"/\*.*?\*/", " ", sql, flags=re.DOTALL)
    return re.sub(r"--[^\r\n]*", " ", without_blocks)


def ensure_select_sql(sql: str) -> None:
    """Reject any SQL other than a single, non-transactional SELECT statement."""
    normalized = _strip_sql_comments(sql).strip()
    if not normalized.upper().startswith("SELECT"):
        raise RuntimeError("읽기 전용 보호: SELECT 문만 실행할 수 있습니다.")
    if ";" in normalized:
        raise RuntimeError("읽기 전용 보호: 다중 SQL 문은 실행할 수 없습니다.")
    if BLOCKED_SQL_PATTERN.search(normalized):
        raise RuntimeError("읽기 전용 보호: 쓰기·DDL·트랜잭션 SQL은 실행할 수 없습니다.")


def execute_select(cursor: Any, sql: str, params: dict[str, Any]) -> list[tuple[Any, ...]]:
    """Execute a validated SELECT query and return all rows."""
    ensure_select_sql(sql)
    cursor.execute(sql, params)
    return list(cursor.fetchall())


def connect_read_only(env_path: Path = ENV_PATH) -> Any:
    """Open an Oracle connection and immediately protect it as read-only."""
    load_dotenv(env_path)
    required = ("DB_USER", "DB_PASSWORD", "DB_HOST", "DB_SERVICE")
    missing = [name for name in required if not os.getenv(name)]
    if missing:
        raise RuntimeError("DB 접속 환경변수가 없습니다: " + ", ".join(missing))

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


def resolve_target(cursor: Any) -> ResolvedTarget:
    """Resolve the requested intersection and westbound approach from master data."""
    intersection_rows = execute_select(
        cursor,
        f"""
        SELECT NODE_ID, CRSRD_NM
        FROM {INTERSECTION_TABLE}
        WHERE REPLACE(CRSRD_NM, ' ', '') = REPLACE(:intersection_name, ' ', '')
        ORDER BY NODE_ID
        """,
        {"intersection_name": INTERSECTION_NAME},
    )
    if len(intersection_rows) != 1:
        raise RuntimeError(
            f"교차로 기준정보가 1건이어야 합니다: {INTERSECTION_NAME} ({len(intersection_rows)}건)"
        )
    node_id, intersection_name = intersection_rows[0]

    approach_rows = execute_select(
        cursor,
        f"""
        SELECT ACSR_ID, ACSR_NM
        FROM {APPROACH_TABLE}
        WHERE NODE_ID = :node_id
        ORDER BY ACSR_ID
        """,
        {"node_id": node_id},
    )
    expected_names = {
        _normalize_name(APPROACH_NAME),
        _normalize_name(f"{intersection_name}-{APPROACH_NAME}"),
    }
    matching_approaches = [
        row for row in approach_rows if _normalize_name(row[1]) in expected_names
    ]
    if len(matching_approaches) != 1:
        raise RuntimeError(
            f"접근로 기준정보가 1건이어야 합니다: {APPROACH_NAME} ({len(matching_approaches)}건)"
        )
    approach_id, approach_name = matching_approaches[0]
    return ResolvedTarget(
        node_id=str(node_id).strip(),
        intersection_name=str(intersection_name).strip(),
        approach_id=str(approach_id).strip(),
        approach_name=str(approach_name).strip(),
    )


def load_vehicle_code_names(cursor: Any) -> dict[str, str]:
    """Return vehicle-class display names keyed by the DB vehicle code."""
    rows = execute_select(
        cursor,
        f"""
        SELECT TRIM(TO_CHAR(CD)) AS VKND_CD, CD_NM
        FROM {VEHICLE_CODE_TABLE}
        WHERE GRP_CD = 'VHCL_ATTR_CD'
        ORDER BY TRIM(TO_CHAR(CD))
        """,
        {},
    )
    return {str(code).strip(): str(name).strip() for code, name in rows}


def load_vehicle_traffic(
    cursor: Any, target: ResolvedTarget, vehicle_code_names: dict[str, str]
) -> list[VehicleTraffic]:
    """Aggregate every vehicle class in the exact requested 15-minute interval."""
    rows = execute_select(
        cursor,
        f"""
        SELECT
            TRIM(TO_CHAR(VKND_CD)) AS VKND_CD,
            NVL(SUM(TRF_QNTY), 0) AS TRAFFIC_VOLUME
        FROM {TRAFFIC_TABLE}
        WHERE NODE_ID = :node_id
          AND ACSR_ID = :approach_id
          AND TOT_DT >= :start_at
          AND TOT_DT < :end_at
        GROUP BY TRIM(TO_CHAR(VKND_CD))
        ORDER BY TRIM(TO_CHAR(VKND_CD))
        """,
        {
            "node_id": target.node_id,
            "approach_id": target.approach_id,
            "start_at": START_AT,
            "end_at": END_AT,
        },
    )
    return [
        VehicleTraffic(
            vehicle_code=str(code).strip(),
            vehicle_name=vehicle_code_names.get(str(code).strip(), "미등록 차종"),
            traffic_volume=int(volume or 0),
        )
        for code, volume in rows
    ]


def save_workbook(target: ResolvedTarget, rows: list[VehicleTraffic], output_path: Path) -> int:
    """Create a concise, auditable workbook and return its overall total."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    total = sum(row.traffic_volume for row in rows)
    workbook = Workbook()
    traffic_sheet = workbook.active
    traffic_sheet.title = "차종별 교통량"
    traffic_sheet.freeze_panes = "A5"

    traffic_sheet.merge_cells("A1:C1")
    traffic_sheet["A1"] = "박촌교삼거리-서 (동향) 15분 차종별 교통량"
    traffic_sheet["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    traffic_sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    traffic_sheet["A1"].alignment = Alignment(horizontal="center")
    traffic_sheet.merge_cells("A2:C2")
    traffic_sheet["A2"] = "조회 구간: 2026-07-13 10:00:00 이상 ~ 10:15:00 미만"
    traffic_sheet["A2"].alignment = Alignment(horizontal="center")
    headers = ["차종 코드", "차종명", "15분 교통량"]
    for column, header in enumerate(headers, start=1):
        cell = traffic_sheet.cell(row=4, column=column, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="5B9BD5")
        cell.alignment = Alignment(horizontal="center")
    for row_index, row in enumerate(rows, start=5):
        traffic_sheet.cell(row=row_index, column=1, value=row.vehicle_code)
        traffic_sheet.cell(row=row_index, column=2, value=row.vehicle_name)
        traffic_sheet.cell(row=row_index, column=3, value=row.traffic_volume)
    total_row = 5 + len(rows)
    traffic_sheet.cell(row=total_row, column=1, value="전체 합계")
    traffic_sheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    traffic_sheet.cell(row=total_row, column=3, value=total)
    for cell in traffic_sheet[total_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for row in traffic_sheet.iter_rows(min_row=4, max_row=total_row, min_col=1, max_col=3):
        for cell in row:
            cell.alignment = Alignment(horizontal="center" if cell.column != 2 else "left")
    for row_number in range(5, total_row + 1):
        traffic_sheet.cell(row=row_number, column=3).number_format = "#,##0"
    traffic_sheet.column_dimensions["A"].width = 16
    traffic_sheet.column_dimensions["B"].width = 24
    traffic_sheet.column_dimensions["C"].width = 16
    traffic_sheet.auto_filter.ref = f"A4:C{max(4, total_row - 1)}"

    info_sheet = workbook.create_sheet("조회정보")
    info_rows = [
        ("교차로", target.intersection_name),
        ("NODE_ID", target.node_id),
        ("접근로", target.approach_name),
        ("ACSR_ID", target.approach_id),
        ("원천 테이블", TRAFFIC_TABLE),
        ("조회 시작", START_AT),
        ("조회 종료(미포함)", END_AT),
        ("조회 차종 수", len(rows)),
        ("전체 합계", total),
    ]
    info_sheet.append(["항목", "값"])
    for cell in info_sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    for info_row in info_rows:
        info_sheet.append(info_row)
    for row in info_sheet.iter_rows(min_row=2, max_row=len(info_rows) + 1, min_col=1, max_col=2):
        row[0].font = Font(bold=True)
    info_sheet["B7"].number_format = "yyyy-mm-dd hh:mm:ss"
    info_sheet["B8"].number_format = "yyyy-mm-dd hh:mm:ss"
    info_sheet["B10"].number_format = "#,##0"
    info_sheet.column_dimensions["A"].width = 20
    info_sheet.column_dimensions["B"].width = 34
    workbook.save(output_path)
    return total


def verify_workbook(
    output_path: Path, expected_rows: list[VehicleTraffic], expected_total: int
) -> None:
    """Reopen the workbook and validate structure and numeric reconciliation."""
    workbook = load_workbook(output_path, data_only=False)
    if workbook.sheetnames != ["차종별 교통량", "조회정보"]:
        raise RuntimeError("Excel 검증 실패: 시트 구성이 일치하지 않습니다.")
    sheet = workbook["차종별 교통량"]
    if [sheet.cell(4, column).value for column in range(1, 4)] != [
        "차종 코드",
        "차종명",
        "15분 교통량",
    ]:
        raise RuntimeError("Excel 검증 실패: 헤더가 일치하지 않습니다.")
    actual_rows = [
        (
            str(sheet.cell(row_number, 1).value),
            str(sheet.cell(row_number, 2).value),
            sheet.cell(row_number, 3).value,
        )
        for row_number in range(5, 5 + len(expected_rows))
    ]
    expected = [(row.vehicle_code, row.vehicle_name, row.traffic_volume) for row in expected_rows]
    if actual_rows != expected:
        raise RuntimeError("Excel 검증 실패: 차종별 집계가 일치하지 않습니다.")
    total_cell = sheet.cell(5 + len(expected_rows), 3)
    if total_cell.value != expected_total or total_cell.number_format != "#,##0":
        raise RuntimeError("Excel 검증 실패: 전체 합계 또는 숫자 형식이 일치하지 않습니다.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_PATH)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    connection = connect_read_only()
    try:
        with connection.cursor() as cursor:
            target = resolve_target(cursor)
            vehicle_code_names = load_vehicle_code_names(cursor)
            rows = load_vehicle_traffic(cursor, target, vehicle_code_names)
    finally:
        connection.close()

    total = save_workbook(target, rows, args.output)
    verify_workbook(args.output, rows, total)
    print(f"추출 차종 수: {len(rows)}")
    print(f"전체 합계: {total:,}")
    print(f"결과 파일: {args.output}")


if __name__ == "__main__":
    main()
