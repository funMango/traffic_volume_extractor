#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""월별 시간당 교차로 교통량 공공데이터 CSV 추출."""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterator

try:
    import oracledb
except ImportError:  # pragma: no cover - runtime dependency check
    oracledb = None

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - runtime dependency check
    load_dotenv = None


BASE_DIR = Path(__file__).resolve().parents[2]
ENV_PATH = BASE_DIR / "00_Data" / ".env"
OUTPUT_DIR = BASE_DIR / "02_Result" / "공공데이터"

TRAFFIC_TABLE = "S_CRSRD_TRF_1HH"
INTERSECTION_TABLE = "M_CRSRD_INF"
CSV_HEADERS = ["시간", "교차로명", "교통량"]
DEFAULT_FETCH_SIZE = 1000

MONTH_PATTERN = re.compile(r"\d{6}")
TIME_TEXT_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}")
CSV_TIME_TEXT_FORMULA_PATTERN = re.compile(r'="(\d{4}-\d{2}-\d{2} \d{2}:\d{2})"')
BLOCKED_SQL_PATTERN = re.compile(
    r"\b(INSERT|UPDATE|DELETE|MERGE|DROP|TRUNCATE|ALTER|CREATE|COMMIT|ROLLBACK|GRANT|REVOKE)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Month:
    year: int
    month: int

    @property
    def text(self) -> str:
        return f"{self.year:04d}{self.month:02d}"


class ConsoleProgress:
    """CSV 저장 진행률을 한 줄로 갱신한다."""

    def __init__(self, total: int, width: int = 30) -> None:
        self.total = max(total, 0)
        self.width = width
        self.done = 0
        self._frames = ["|", "/", "-", "\\"]
        self._frame_idx = 0
        self._last_render = 0.0
        self._last_len = 0

    def advance(self, step: int = 1) -> None:
        self.done = min(self.total, self.done + step)
        now = time.perf_counter()
        if self.done == self.total or now - self._last_render >= 0.1:
            self.render()
            self._last_render = now

    def finish(self) -> None:
        self.done = self.total
        self.render()
        print()

    def render(self) -> None:
        if self.total <= 0:
            pct = 100.0
            filled = self.width
        else:
            pct = self.done / self.total * 100
            filled = int(self.width * self.done / self.total)

        frame = self._frames[self._frame_idx]
        self._frame_idx = (self._frame_idx + 1) % len(self._frames)
        bar = "#" * filled + "-" * (self.width - filled)
        line = f"\r{frame} [{bar}] {pct:6.2f}% ({self.done:,}/{self.total:,}) CSV 저장"
        pad = max(0, self._last_len - len(line))
        print(line + (" " * pad), end="", flush=True)
        self._last_len = len(line)


def _strip_sql_comments(sql: str) -> str:
    sql = re.sub(r"/\*.*?\*/", " ", sql, flags=re.DOTALL)
    return re.sub(r"--[^\n\r]*", " ", sql)


def _ensure_select_sql(sql: str) -> None:
    normalized = _strip_sql_comments(sql).strip()
    if not normalized:
        raise RuntimeError("빈 SQL은 실행할 수 없습니다.")
    if not normalized.upper().startswith("SELECT"):
        raise RuntimeError("읽기 전용 보호: SELECT 문만 실행할 수 있습니다.")
    if ";" in normalized.rstrip(";"):
        raise RuntimeError("읽기 전용 보호: 복수 SQL 문은 실행할 수 없습니다.")
    if BLOCKED_SQL_PATTERN.search(normalized):
        raise RuntimeError("읽기 전용 보호: 쓰기/DDL/트랜잭션 SQL은 실행할 수 없습니다.")


def _execute_select(
    cursor: Any,
    sql: str,
    params: dict[str, Any] | None = None,
) -> list[tuple[Any, ...]]:
    _ensure_select_sql(sql)
    cursor.execute(sql, params or {})
    return list(cursor.fetchall())


def _iter_select(
    cursor: Any,
    sql: str,
    params: dict[str, Any],
    fetch_size: int,
) -> Iterator[tuple[Any, ...]]:
    _ensure_select_sql(sql)
    cursor.arraysize = fetch_size
    cursor.execute(sql, params)
    while True:
        rows = cursor.fetchmany(fetch_size)
        if not rows:
            break
        yield from rows


def _connect_db(env_path: Path) -> Any:
    if oracledb is None:
        raise RuntimeError("oracledb 패키지가 설치되어 있지 않습니다.")
    if load_dotenv is None:
        raise RuntimeError("python-dotenv 패키지가 설치되어 있지 않습니다.")

    load_dotenv(env_path)
    required = ["DB_USER", "DB_PASSWORD", "DB_HOST", "DB_SERVICE"]
    missing = [name for name in required if not os.getenv(name)]
    if missing:
        raise RuntimeError(
            "DB 접속 환경변수가 없습니다: "
            + ", ".join(missing)
            + f" (.env 경로 확인: {env_path})"
        )

    conn = oracledb.connect(
        user=os.getenv("DB_USER", "").strip(),
        password=os.getenv("DB_PASSWORD", "").strip(),
        host=os.getenv("DB_HOST", "").strip(),
        port=int(os.getenv("DB_PORT", "1521").strip()),
        service_name=os.getenv("DB_SERVICE", "").strip(),
    )
    try:
        with conn.cursor() as cur:
            cur.execute("SET TRANSACTION READ ONLY")
    except Exception:
        conn.close()
        raise
    return conn


def _get_table_columns(cursor: Any, table_name: str) -> set[str]:
    rows = _execute_select(
        cursor,
        """
        SELECT COLUMN_NAME
        FROM ALL_TAB_COLUMNS
        WHERE TABLE_NAME = :table_name
        """,
        {"table_name": table_name.upper()},
    )
    return {str(row[0]).upper() for row in rows}


def _validate_tables(cursor: Any) -> None:
    required = {
        TRAFFIC_TABLE: {"NODE_ID", "TOT_DT", "TRF_QNTY"},
        INTERSECTION_TABLE: {"NODE_ID", "CRSRD_NM"},
    }
    for table_name, required_columns in required.items():
        columns = _get_table_columns(cursor, table_name)
        missing = sorted(required_columns - columns)
        if missing:
            raise RuntimeError(f"{table_name} 필수 컬럼이 없습니다: {', '.join(missing)}")


def parse_month(value: str) -> Month:
    month_text = value.strip()
    if not MONTH_PATTERN.fullmatch(month_text):
        raise ValueError("월은 YYYYMM 형식으로 입력해 주세요. 예: 202604")

    year = int(month_text[:4])
    month = int(month_text[4:])
    if not 1 <= month <= 12:
        raise ValueError("월 값은 01부터 12까지 입력해 주세요.")
    return Month(year=year, month=month)


def next_month(month: Month) -> Month:
    if month.month == 12:
        return Month(year=month.year + 1, month=1)
    return Month(year=month.year, month=month.month + 1)


def month_bounds(month: Month) -> tuple[date, date]:
    end = next_month(month)
    return date(month.year, month.month, 1), date(end.year, end.month, 1)


def output_path_for_month(month: Month, output_dir: Path = OUTPUT_DIR) -> Path:
    return output_dir / f"{month.text}_공공데이터.csv"


def _read_month_from_stdin() -> Month:
    while True:
        raw = input("조회 월(YYYYMM): ").strip()
        try:
            return parse_month(raw)
        except ValueError as exc:
            print(f"  {exc}")


def _traffic_query_sql() -> str:
    return f"""
        SELECT TO_CHAR(grouped.TRAFFIC_DT, 'YYYY-MM-DD HH24:MI') AS TIME_TEXT,
               grouped.INTERSECTION_NAME,
               grouped.TRAFFIC_VOLUME
        FROM (
            SELECT t.TOT_DT AS TRAFFIC_DT,
                   t.NODE_ID,
                   i.CRSRD_NM AS INTERSECTION_NAME,
                   SUM(t.TRF_QNTY) AS TRAFFIC_VOLUME
            FROM {TRAFFIC_TABLE} t
            JOIN {INTERSECTION_TABLE} i
              ON i.NODE_ID = t.NODE_ID
            WHERE t.TOT_DT >= :date_start
              AND t.TOT_DT < :date_end_next
            GROUP BY t.TOT_DT,
                     t.NODE_ID,
                     i.CRSRD_NM
        ) grouped
        ORDER BY grouped.TRAFFIC_DT,
                 grouped.INTERSECTION_NAME,
                 grouped.NODE_ID
        """


def _traffic_count_sql() -> str:
    return f"""
        SELECT COUNT(*)
        FROM (
            SELECT t.TOT_DT,
                   t.NODE_ID,
                   i.CRSRD_NM
            FROM {TRAFFIC_TABLE} t
            JOIN {INTERSECTION_TABLE} i
              ON i.NODE_ID = t.NODE_ID
            WHERE t.TOT_DT >= :date_start
              AND t.TOT_DT < :date_end_next
            GROUP BY t.TOT_DT,
                     t.NODE_ID,
                     i.CRSRD_NM
        )
        """


def _format_traffic_volume(value: Any) -> str | int | float:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        if value == value.to_integral_value():
            return int(value)
        return format(value.normalize(), "f")
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return value
    return value


def _normalize_time_text(value: Any) -> str:
    time_value = str(value)
    if not TIME_TEXT_PATTERN.fullmatch(time_value):
        raise RuntimeError(f"시간 형식이 다릅니다: {time_value}")
    return time_value


def _format_csv_time_cell(value: Any) -> str:
    time_value = _normalize_time_text(value)
    return f'="{time_value}"'


def _extract_csv_time_text(value: str) -> str:
    match = CSV_TIME_TEXT_FORMULA_PATTERN.fullmatch(value)
    if not match:
        raise RuntimeError(f"CSV 시간 컬럼이 Excel 텍스트식이 아닙니다: {value}")

    time_value = match.group(1)
    if not TIME_TEXT_PATTERN.fullmatch(time_value):
        raise RuntimeError(f"CSV 시간 형식이 다릅니다: {time_value}")
    return time_value


def _fetch_output_count(cursor: Any, params: dict[str, Any]) -> int:
    rows = _execute_select(cursor, _traffic_count_sql(), params)
    if len(rows) != 1:
        raise RuntimeError(f"COUNT 결과가 1행이 아닙니다: {len(rows)}행")
    return int(rows[0][0])


def _write_csv(
    cursor: Any,
    output_path: Path,
    params: dict[str, Any],
    expected_count: int,
    fetch_size: int,
) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    progress = ConsoleProgress(expected_count)
    rows_written = 0

    with output_path.open("w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.writer(fp)
        writer.writerow(CSV_HEADERS)
        if expected_count == 0:
            progress.finish()
            return rows_written

        for time_text, intersection_name, traffic_volume in _iter_select(
            cursor,
            _traffic_query_sql(),
            params,
            fetch_size=fetch_size,
        ):
            writer.writerow(
                [
                    _format_csv_time_cell(time_text),
                    str(intersection_name),
                    _format_traffic_volume(traffic_volume),
                ]
            )
            rows_written += 1
            progress.advance(1)

    progress.finish()
    return rows_written


def _verify_csv(output_path: Path, expected_count: int) -> None:
    if not output_path.exists():
        raise RuntimeError(f"CSV 파일이 생성되지 않았습니다: {output_path}")

    with output_path.open("r", newline="", encoding="utf-8-sig") as fp:
        reader = csv.reader(fp)
        try:
            header = next(reader)
        except StopIteration as exc:
            raise RuntimeError("CSV 파일이 비어 있습니다.") from exc

        if header != CSV_HEADERS:
            raise RuntimeError(f"CSV 헤더가 다릅니다: {header}")

        row_count = 0
        for row in reader:
            row_count += 1
            if len(row) != len(CSV_HEADERS):
                raise RuntimeError(f"CSV 컬럼 수가 다릅니다: {row}")
            _extract_csv_time_text(row[0])

    if row_count != expected_count:
        raise RuntimeError(f"CSV 행 수가 COUNT 결과와 다릅니다: CSV={row_count}, COUNT={expected_count}")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="월별 시간당 교차로 교통량 공공데이터 CSV 생성")
    parser.add_argument("--month", help="조회 월을 YYYYMM 형식으로 지정합니다. 예: 202604")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="CSV 저장 폴더")
    parser.add_argument(
        "--fetch-size",
        type=int,
        default=DEFAULT_FETCH_SIZE,
        help=f"DB fetchmany 크기. 기본값: {DEFAULT_FETCH_SIZE}",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    conn = None
    try:
        month = parse_month(args.month) if args.month else _read_month_from_stdin()
        if args.fetch_size <= 0:
            raise RuntimeError("--fetch-size는 1 이상이어야 합니다.")

        date_start, date_end_next = month_bounds(month)
        params = {"date_start": date_start, "date_end_next": date_end_next}
        output_path = output_path_for_month(month, Path(args.output_dir))

        print(f"조회 월: {month.text}", flush=True)
        print("DB 연결 및 읽기전용 보호 시작", flush=True)
        conn = _connect_db(ENV_PATH)
        print("SET TRANSACTION READ ONLY 완료", flush=True)
        with conn.cursor() as cur:
            print("필수 테이블/컬럼 검증 중...", flush=True)
            _validate_tables(cur)
            print("전체 출력 행 수 계산 중...", flush=True)
            expected_count = _fetch_output_count(cur, params)
            print(f"전체 출력 행 수: {expected_count:,}행", flush=True)
            rows_written = _write_csv(
                cur,
                output_path,
                params,
                expected_count,
                fetch_size=args.fetch_size,
            )

        if rows_written != expected_count:
            raise RuntimeError(
                f"저장 행 수가 COUNT 결과와 다릅니다: 저장={rows_written}, COUNT={expected_count}"
            )
        _verify_csv(output_path, expected_count)
    except ValueError as exc:
        print(f"실행을 중단합니다: {exc}", file=sys.stderr)
        return 1
    except RuntimeError as exc:
        print(f"실행을 중단합니다: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - DB driver/runtime errors
        print(f"실행을 중단합니다: {exc}", file=sys.stderr)
        return 1
    finally:
        if conn is not None:
            conn.close()

    print("전체 진행률 100.00%", flush=True)
    print(f"저장 경로: {output_path.resolve()}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
