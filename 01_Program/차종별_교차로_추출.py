#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""차종별 교차로 교통량 CSV 추출 프로그램.

S_CRSRD_VKND_TRF_1HH를 읽기 전용 SELECT로 조회해
교차로, 시간, 교차로 방향, 방향, 차종, 교통량 컬럼의 CSV를 생성한다.
"""

from __future__ import annotations

import csv
import difflib
import os
import re
import sqlite3
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from threading import Event, Lock, Thread
from typing import Iterable

try:
    import oracledb
except ImportError:  # pragma: no cover - 실행 환경 안내용
    oracledb = None

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - 실행 환경 안내용
    load_dotenv = None


BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / "00_Data" / ".env"
RESULT_DIR = BASE_DIR / "02_Result" / "차종별_교통량_추출"
CSV_HEADER = ["교차로", "시간", "교차로 방향", "방향", "차종", "교통량"]
ALL_INTERSECTION_TOKENS = {"*", "전체", "all"}
PEAK_HOURS = [7, 8, 16, 17, 18]
OUTPUT_FORMAT_EXTENSIONS = {
    "csv": ".csv",
    "xlsx": ".xlsx",
    "db": ".db",
}
OUTPUT_FORMAT_LABELS = {
    "csv": "CSV 저장",
    "xlsx": "XLSX 저장",
    "db": "SQLite DB 저장",
}


@dataclass(frozen=True)
class Period:
    start: date
    end: date
    token: str


@dataclass(frozen=True)
class Intersection:
    node_id: int | str
    name: str


class StepProgress:
    """스피너와 퍼센트를 함께 출력하는 간단한 진행 표시기."""

    def __init__(self, label: str, total: int = 1) -> None:
        self.label = label
        self.total = max(total, 1)
        self.done = 0
        self._frames = ("|", "/", "-", "\\")
        self._frame_index = 0
        self._last_len = 0
        self._lock = Lock()
        self._stop = Event()
        self._thread: Thread | None = None

    def start(self) -> None:
        if self._thread is not None:
            return
        self._thread = Thread(target=self._run, daemon=True)
        self._thread.start()

    def advance(self, step: int = 1) -> None:
        with self._lock:
            self.done = min(self.total, self.done + step)
        self._render()

    def finish(self) -> None:
        with self._lock:
            self.done = self.total
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=1.0)
        self._render()
        print()

    def _run(self) -> None:
        while not self._stop.wait(0.12):
            self._frame_index = (self._frame_index + 1) % len(self._frames)
            self._render()

    def _render(self) -> None:
        with self._lock:
            done = self.done
            total = self.total
        pct = int((done / total) * 100)
        frame = self._frames[self._frame_index]
        line = f"\r{self.label} {frame} {pct:3d}% ({done}/{total})"
        pad = " " * max(0, self._last_len - len(line))
        print(line + pad, end="", flush=True)
        self._last_len = len(line)


def run_instant_step(label: str, work) -> object:
    progress = StepProgress(label)
    progress.start()
    try:
        result = work()
    finally:
        progress.finish()
    return result


def print_step_done(label: str) -> None:
    print(f"{label} | 100% (1/1)")


def connect_db():
    """00_Data/.env를 로드한 뒤 DB에 연결한다."""
    if oracledb is None:
        raise RuntimeError("oracledb 패키지가 설치되어 있지 않습니다.")
    if load_dotenv is None:
        raise RuntimeError("python-dotenv 패키지가 설치되어 있지 않습니다.")

    load_dotenv(ENV_PATH)

    required = ["DB_USER", "DB_PASSWORD", "DB_HOST", "DB_SERVICE"]
    missing = [name for name in required if not os.getenv(name)]
    if missing:
        raise RuntimeError(
            "DB 접속 환경변수가 없습니다: "
            + ", ".join(missing)
            + f"\n.env 경로를 확인하세요: {ENV_PATH}"
        )

    conn = oracledb.connect(
        user=os.getenv("DB_USER", "").strip(),
        password=os.getenv("DB_PASSWORD", "").strip(),
        host=os.getenv("DB_HOST", "").strip(),
        port=int(os.getenv("DB_PORT", "1521").strip()),
        service_name=os.getenv("DB_SERVICE", "").strip(),
    )
    with conn.cursor() as cur:
        cur.execute("SET TRANSACTION READ ONLY")
    return conn


def normalize_code(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_yymmdd(value: str) -> date:
    if not re.fullmatch(r"\d{6}", value):
        raise ValueError(f"날짜 형식이 올바르지 않습니다: {value}")
    return date(2000 + int(value[:2]), int(value[2:4]), int(value[4:6]))


def month_end(year: int, month: int) -> date:
    if month == 12:
        return date(year, 12, 31)
    return date(year, month + 1, 1) - timedelta(days=1)


def parse_period_segment(segment: str) -> Period:
    token = re.sub(r"\s+", "", segment)
    if re.fullmatch(r"\d{6}~\d{6}", token):
        start_raw, end_raw = token.split("~", 1)
        start = parse_yymmdd(start_raw)
        end = parse_yymmdd(end_raw)
    elif re.fullmatch(r"\d{4}", token):
        year = 2000 + int(token[:2])
        month = int(token[2:4])
        start = date(year, month, 1)
        end = month_end(year, month)
    elif re.fullmatch(r"\d{6}", token):
        start = parse_yymmdd(token)
        end = start
    else:
        raise ValueError(f"기간 형식이 올바르지 않습니다: {segment}")
    if start > end:
        raise ValueError(f"시작일이 종료일보다 늦습니다: {segment}")
    return Period(start=start, end=end, token=token)


def parse_periods(raw: str) -> list[Period]:
    parts = [part.strip() for part in raw.split(",") if part.strip()]
    if not parts:
        raise ValueError("기간을 입력해 주세요.")
    return [parse_period_segment(part) for part in parts]


def parse_time_range(raw: str) -> list[int]:
    value = raw.strip().lower()
    if value in {"24시간", "1"}:
        return list(range(24))
    if value in {"첨두시", "2"}:
        return PEAK_HOURS.copy()

    if value in {"직접입력", "직접", "3"}:
        value = input("시간대 직접 입력 (예: 07~09, 18~20): ").strip()

    hours: set[int] = set()
    for segment in [part.strip() for part in value.split(",") if part.strip()]:
        match = re.fullmatch(r"(\d{1,2})(?::00)?\s*~\s*(\d{1,2})(?::00)?", segment)
        if not match:
            raise ValueError(f"시간대 형식이 올바르지 않습니다: {segment}")
        start_hour = int(match.group(1))
        end_hour = int(match.group(2))
        if not (0 <= start_hour <= 23 and 1 <= end_hour <= 24 and start_hour < end_hour):
            raise ValueError(f"시간대 범위가 올바르지 않습니다: {segment}")
        hours.update(range(start_hour, end_hour))
    if not hours:
        raise ValueError("시간대를 입력해 주세요.")
    return sorted(hours)


def input_periods() -> list[Period]:
    print("\n[기간 입력]")
    print("  예: 260410~260412 / 2604 / 260410 / 2604, 2605")
    while True:
        try:
            periods = parse_periods(input("기간: "))
        except ValueError as exc:
            print(f"  오류: {exc}")
            continue
        for period in periods:
            print(f"  - {period.token}: {period.start} ~ {period.end}")
        return periods


def input_hours() -> list[int]:
    print("\n[시간대 입력]")
    print("  1 또는 24시간: 0~23시")
    print("  2 또는 첨두시: 07~09, 16~19 (7, 8, 16, 17, 18시)")
    print("  3 또는 직접입력: 예) 07~09, 18~20")
    while True:
        try:
            hours = parse_time_range(input("시간대: "))
        except ValueError as exc:
            print(f"  오류: {exc}")
            continue
        print(f"  선택 시간: {', '.join(str(h) for h in hours)}시")
        return hours


def input_output_format() -> str:
    print("\n[출력 방식 입력]")
    print("  1 또는 csv  : CSV 파일")
    print("  2 또는 xlsx : Excel 파일")
    print("  3 또는 db   : SQLite DB 파일")
    print("  기본값: csv")
    choices = {
        "": "csv",
        "1": "csv",
        "csv": "csv",
        "2": "xlsx",
        "xlsx": "xlsx",
        "3": "db",
        "db": "db",
    }
    while True:
        selected = choices.get(input("출력 방식: ").strip().lower())
        if selected is None:
            print("  오류: csv, xlsx, db 또는 1, 2, 3 중 하나를 입력해 주세요.")
            continue
        print(f"  선택 출력 방식: {selected}")
        return selected


def load_reference_data(conn) -> tuple[list[Intersection], dict[str, str], dict[str, str]]:
    with conn.cursor() as cur:
        cur.execute("SELECT NODE_ID, CRSRD_NM FROM M_CRSRD_INF ORDER BY CRSRD_NM")
        intersections = [Intersection(node_id=row[0], name=row[1]) for row in cur.fetchall()]

        cur.execute(
            """
            SELECT TRIM(TO_CHAR(CD)) AS CD, CD_NM
            FROM M_CD_INF
            WHERE GRP_CD = 'DRCT_CD'
            """
        )
        drct_names = {normalize_code(code): name for code, name in cur.fetchall()}

        cur.execute(
            """
            SELECT TRIM(TO_CHAR(CD)) AS CD, CD_NM
            FROM M_CD_INF
            WHERE GRP_CD = 'VHCL_ATTR_CD'
            """
        )
        vknd_names = {normalize_code(code): name for code, name in cur.fetchall()}

    return intersections, drct_names, vknd_names


def resolve_intersections(raw: str, all_intersections: list[Intersection]) -> tuple[list[Intersection], bool]:
    raw = raw.strip()
    if raw.lower() in ALL_INTERSECTION_TOKENS:
        return all_intersections.copy(), True

    by_name = {item.name: item for item in all_intersections}
    names = [item.name for item in all_intersections]
    selected: list[Intersection] = []
    seen: set[int | str] = set()

    for part in [part.strip() for part in raw.split(",") if part.strip()]:
        item = by_name.get(part)
        if item is None:
            candidates = difflib.get_close_matches(part, names, n=5, cutoff=0.3)
            candidate_text = ", ".join(candidates) if candidates else "후보 없음"
            raise LookupError(f"'{part}'을(를) 찾을 수 없습니다. 유사 교차로: {candidate_text}")
        if item.node_id not in seen:
            selected.append(item)
            seen.add(item.node_id)

    if not selected:
        raise ValueError("교차로를 입력해 주세요.")
    return selected, False


def input_intersections(all_intersections: list[Intersection]) -> tuple[list[Intersection], bool]:
    print("\n[교차로 입력]")
    print("  전체: *, 전체, all")
    print("  복수: 교차로이름, 교차로이름")
    print(f"  전체 교차로 수: {len(all_intersections)}개")
    while True:
        try:
            selected, is_all = resolve_intersections(input("교차로: "), all_intersections)
        except (LookupError, ValueError) as exc:
            print(f"  오류: {exc}")
            continue
        if is_all:
            print(f"  전체 교차로 {len(selected)}개 선택")
        else:
            print(f"  선택 교차로: {', '.join(item.name for item in selected)}")
        return selected, is_all


def make_safe_filename_part(value: str) -> str:
    cleaned = re.sub(r'[\\/*?:"<>|]', "_", value).strip()
    return cleaned or "_"


def make_filename(intersections: list[Intersection], is_all: bool, period_tokens: Iterable[str]) -> str:
    period_part = "_".join(period_tokens)
    if is_all:
        name_part = "전체교차로"
    elif len(intersections) == 1:
        name_part = intersections[0].name
    else:
        name_part = f"{intersections[0].name}외{len(intersections) - 1}개"
    return f"{make_safe_filename_part(name_part)}_{period_part}"


def build_in_clause(prefix: str, values: Iterable[object], bind_params: dict) -> str:
    names = []
    for index, value in enumerate(values):
        bind_name = f"{prefix}{index}"
        bind_params[bind_name] = value
        names.append(f":{bind_name}")
    return ", ".join(names)


def fetch_rows_for_step(
    conn,
    intersection: Intersection,
    period: Period,
    hours: list[int],
) -> list[tuple]:
    start_dt = datetime.combine(period.start, datetime.min.time())
    end_next = datetime.combine(period.end + timedelta(days=1), datetime.min.time())
    bind_params: dict[str, object] = {
        "node_id": intersection.node_id,
        "start_dt": start_dt,
        "end_next": end_next,
    }
    hour_clause = ""
    if set(hours) != set(range(24)):
        hour_clause = (
            "AND TO_NUMBER(TO_CHAR(v.TOT_DT, 'HH24')) IN "
            f"({build_in_clause('hour', sorted(set(hours)), bind_params)})"
        )

    sql = f"""
        SELECT
            c.CRSRD_NM AS CRSRD_NM,
            v.TOT_DT AS TOT_DT,
            a.ACSR_NM AS ACSR_NM,
            TRIM(TO_CHAR(v.DRCT_CD)) AS DRCT_CD,
            d.CD_NM AS DRCT_NM,
            TRIM(TO_CHAR(v.VKND_CD)) AS VKND_CD,
            k.CD_NM AS VKND_NM,
            SUM(NVL(v.TRF_QNTY, 0)) AS TRF_QNTY
        FROM S_CRSRD_VKND_TRF_1HH v
        JOIN M_CRSRD_INF c
          ON c.NODE_ID = v.NODE_ID
        LEFT JOIN M_CRSRD_ACSR_INF a
          ON a.NODE_ID = v.NODE_ID
         AND a.ACSR_ID = v.ACSR_ID
        LEFT JOIN M_CD_INF d
          ON d.GRP_CD = 'DRCT_CD'
         AND TRIM(TO_CHAR(d.CD)) = TRIM(TO_CHAR(v.DRCT_CD))
        LEFT JOIN M_CD_INF k
          ON k.GRP_CD = 'VHCL_ATTR_CD'
         AND TRIM(TO_CHAR(k.CD)) = TRIM(TO_CHAR(v.VKND_CD))
        WHERE v.NODE_ID = :node_id
          AND v.TOT_DT >= :start_dt
          AND v.TOT_DT < :end_next
          {hour_clause}
          AND TRIM(TO_CHAR(v.VKND_CD)) NOT IN ('0', '1')
          AND NVL(k.CD_NM, TRIM(TO_CHAR(v.VKND_CD))) <> '미확인'
        GROUP BY
            c.CRSRD_NM,
            v.TOT_DT,
            a.ACSR_NM,
            TRIM(TO_CHAR(v.DRCT_CD)),
            d.CD_NM,
            TRIM(TO_CHAR(v.VKND_CD)),
            k.CD_NM
        ORDER BY
            c.CRSRD_NM,
            v.TOT_DT,
            a.ACSR_NM,
            d.CD_NM,
            k.CD_NM
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, bind_params)
        return cur.fetchall()


def format_time(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value)


def normalize_result_row(row: tuple) -> list[object]:
    crsrd_nm, tot_dt, acsr_nm, drct_cd, drct_nm, vknd_cd, vknd_nm, trf_qnty = row
    return [
        crsrd_nm,
        format_time(tot_dt),
        acsr_nm or "",
        drct_nm or normalize_code(drct_cd),
        vknd_nm or normalize_code(vknd_cd),
        int(trf_qnty or 0),
    ]


def fetch_all_rows(conn, intersections: list[Intersection], periods: list[Period], hours: list[int]) -> list[list]:
    progress = StepProgress("[3/4] 교통량 조회", total=len(intersections) * len(periods))
    progress.start()
    rows: list[list] = []
    try:
        for intersection in intersections:
            for period in periods:
                fetched = fetch_rows_for_step(conn, intersection, period, hours)
                rows.extend(normalize_result_row(row) for row in fetched)
                progress.advance()
    finally:
        progress.finish()
    rows.sort(key=lambda row: (row[0], row[1], row[2], row[3], row[4]))
    return rows


def save_csv(rows: list[list], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADER)
        writer.writerows(rows)


def save_xlsx(rows: list[list], output_path: Path) -> None:
    try:
        from openpyxl import Workbook
    except ImportError as exc:  # pragma: no cover - 실행 환경 안내용
        raise RuntimeError(
            "XLSX 저장에는 openpyxl 패키지가 필요합니다. "
            "설치 후 다시 실행해 주세요: python -m pip install openpyxl"
        ) from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "교통량"
    ws.append(CSV_HEADER)
    for row in rows:
        ws.append(row)
    wb.save(output_path)


def recreate_sqlite_table(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP TABLE IF EXISTS traffic_by_vehicle;

        CREATE TABLE traffic_by_vehicle (
            id INTEGER PRIMARY KEY,
            intersection TEXT NOT NULL,
            observed_at TEXT NOT NULL,
            approach_direction TEXT NOT NULL,
            direction TEXT NOT NULL,
            vehicle_type TEXT NOT NULL,
            traffic_volume INTEGER NOT NULL
        );
        """
    )


def create_sqlite_indexes(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_traffic_intersection_time
            ON traffic_by_vehicle (intersection, observed_at);
        CREATE INDEX IF NOT EXISTS idx_traffic_time
            ON traffic_by_vehicle (observed_at);
        CREATE INDEX IF NOT EXISTS idx_traffic_vehicle_type
            ON traffic_by_vehicle (vehicle_type);
        CREATE INDEX IF NOT EXISTS idx_traffic_direction
            ON traffic_by_vehicle (direction);
        """
    )


def to_sqlite_record(row: list[object]) -> tuple[str, str, str, str, str, int]:
    return (
        str(row[0]),
        str(row[1]),
        str(row[2]),
        str(row[3]),
        str(row[4]),
        int(row[5] or 0),
    )


def save_sqlite(rows: list[list], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(output_path)
    try:
        recreate_sqlite_table(conn)
        conn.executemany(
            """
            INSERT INTO traffic_by_vehicle (
                intersection,
                observed_at,
                approach_direction,
                direction,
                vehicle_type,
                traffic_volume
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            [to_sqlite_record(row) for row in rows],
        )
        create_sqlite_indexes(conn)
        conn.commit()
    finally:
        conn.close()


def save_rows(rows: list[list], output_path: Path, output_format: str) -> None:
    if output_format == "csv":
        save_csv(rows, output_path)
    elif output_format == "xlsx":
        save_xlsx(rows, output_path)
    elif output_format == "db":
        save_sqlite(rows, output_path)
    else:
        raise ValueError(f"지원하지 않는 출력 방식입니다: {output_format}")


def main() -> None:
    started = time.perf_counter()
    print("=" * 60)
    print("  차종별 교차로 교통량 추출 프로그램")
    print("=" * 60)
    output_format = input_output_format()

    try:
        conn = connect_db()
    except Exception as exc:
        print(f"\nDB 연결 실패: {exc}")
        sys.exit(1)

    try:
        all_intersections, _drct_names, _vknd_names = run_instant_step(
            "[1/4] 코드/교차로 정보 로드",
            lambda: load_reference_data(conn),
        )

        selected_intersections, is_all = input_intersections(all_intersections)
        periods = input_periods()
        hours = input_hours()
        print_step_done("[2/4] 입력값 해석 및 조회 조건 확정")

        rows = fetch_all_rows(conn, selected_intersections, periods, hours)

        def _save_work():
            base_filename = make_filename(
                selected_intersections,
                is_all,
                (period.token for period in periods),
            )
            filename = base_filename + OUTPUT_FORMAT_EXTENSIONS[output_format]
            output_path = RESULT_DIR / filename
            save_rows(rows, output_path, output_format)
            return output_path

        save_label = f"[4/4] {OUTPUT_FORMAT_LABELS[output_format]}"
        output_path = run_instant_step(save_label, _save_work)
    finally:
        conn.close()

    elapsed = time.perf_counter() - started
    print(f"\n저장 완료: {output_path}")
    print(f"레코드 수: {len(rows):,}건")
    print(f"소요 시간: {elapsed:.1f}초")


if __name__ == "__main__":
    main()
