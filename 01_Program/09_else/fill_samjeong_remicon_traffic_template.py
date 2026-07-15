#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fill the Samjeong-dong remicon traffic template from a read-only SQLite DB."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re
import shutil
import sqlite3
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_DB_PATH = BASE_DIR / "00_Data" / "삼정동_레미콘" / "삼정동_레미콘_데이터.db"
DEFAULT_TEMPLATE_PATH = (
    BASE_DIR / "00_Data" / "삼정동_레미콘" / "삼정동_레미콘_교통량_작성서식.xlsx"
)
DEFAULT_OUTPUT_PATH = BASE_DIR / "02_Result" / "09_기타" / "삼정동_레미콘_교통량_작성서식_결과.xlsx"

TABLE_NAME = "vehicle_detection"
SHEET_NAME = "작성서식"
REQUIRED_COLUMNS = {
    "id",
    "collected_at",
    "location_name",
    "vehicle_number",
}
REMICON_PATTERN = re.compile(r"^014[가-힣]")
HOLIDAYS = {
    date(2026, 5, 1),
    date(2026, 5, 5),
    date(2026, 5, 25),
    date(2026, 6, 3),
}

MONTH_ROWS = {
    "5월": range(4, 12),
    "6월": range(12, 20),
}
MONTH_BOUNDS = {
    "5월": ("2026-05-01 00:00:00", "2026-06-01 00:00:00"),
    "6월": ("2026-06-01 00:00:00", "2026-07-01 00:00:00"),
}
EXPECTED_HEADER_VALUES = {
    "C2": "월별 통행량",
    "F2": "평일 평균 통행량",
    "I2": "평일 첨두시(07:00~09:00, 17:00~19:00) 평균 통행량",
    "C3": "총 통행량(대)",
    "D3": "레미콘 통행량(대)",
    "E3": "비율(%)",
    "F3": "총 통행량(대)",
    "G3": "레미콘 통행량(대)",
    "H3": "비율(%)",
    "I3": "총 통행량(대)",
    "J3": "레미콘 통행량(대)",
    "K3": "비율(%)",
}
LOCATION_MAPPING = {
    "박촌교 삼거리": [
        "박촌교 삼거리[남향]",
        "박촌교 삼거리[북향]",
        "박촌교 삼거리[서향(순방향)]",
        "박촌교 삼거리[서향(역방향)]",
    ],
    "봉오고가교사거리": [
        "봉오고가교 사거리[남향]",
        "봉오고가교 사거리[서향]",
    ],
    "삼정고가삼거리": [
        "삼정고가 삼거리[동향]",
        "삼정고가 삼거리[서향]",
    ],
    "삼정교사거리": [
        "삼정교 사거리[북동향]",
    ],
    "산업길사거리": [
        "산업길 사거리[남향]",
        "산업길 사거리[북향]",
    ],
    "봉오대로사거리": [
        "봉오대로 사거리[동향]",
        "봉오대로 사거리[서향]",
    ],
    "자동차검사소": [
        "자동차검사소",
    ],
    "삼정동 320-1": [
        "삼정동320-1 부천IC",
    ],
}


@dataclass(frozen=True)
class TrafficCounts:
    total: int
    remicon: int

    @property
    def ratio(self) -> float:
        if self.total == 0:
            return 0.0
        ratio = Decimal(self.remicon) / Decimal(self.total) * Decimal("100")
        return float(ratio.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


@dataclass(frozen=True)
class TrafficMetrics:
    monthly: TrafficCounts
    weekday: TrafficCounts
    peak: TrafficCounts


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill the Samjeong-dong remicon traffic template.")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_PATH)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_PATH)
    return parser.parse_args()


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    resolved = db_path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"DB file does not exist: {resolved}")
    conn = sqlite3.connect(f"file:{resolved.as_posix()}?mode=ro", uri=True)
    conn.create_function("IS_REMICON", 1, is_remicon)
    return conn


def is_remicon(vehicle_number: object) -> int:
    return int(bool(REMICON_PATTERN.match(str(vehicle_number or ""))))


def validate_database(conn: sqlite3.Connection) -> None:
    tables = {
        str(row[0])
        for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall()
    }
    if TABLE_NAME not in tables:
        raise RuntimeError(f"Missing required table: {TABLE_NAME}")

    columns = {str(row[1]) for row in conn.execute(f"PRAGMA table_info({TABLE_NAME})").fetchall()}
    missing_columns = sorted(REQUIRED_COLUMNS - columns)
    if missing_columns:
        raise RuntimeError("Missing required columns: " + ", ".join(missing_columns))


def validate_locations(conn: sqlite3.Connection) -> None:
    expected_locations = {name for names in LOCATION_MAPPING.values() for name in names}
    actual_locations = {
        str(row[0])
        for row in conn.execute(f"SELECT DISTINCT location_name FROM {TABLE_NAME}").fetchall()
    }
    missing_locations = sorted(expected_locations - actual_locations)
    if missing_locations:
        raise RuntimeError("Missing mapped DB locations: " + ", ".join(missing_locations))


def validate_template(ws: Worksheet) -> None:
    for cell_ref, expected_value in EXPECTED_HEADER_VALUES.items():
        actual_value = ws[cell_ref].value
        if actual_value != expected_value:
            raise RuntimeError(
                f"Unexpected template header {cell_ref}: {actual_value!r} != {expected_value!r}"
            )

    for month_label, rows in MONTH_ROWS.items():
        month_cell = ws[f"A{rows.start}"].value
        if month_cell != month_label:
            raise RuntimeError(
                f"Unexpected month label A{rows.start}: {month_cell!r} != {month_label!r}"
            )
        actual_labels = [str(ws[f"B{row}"].value or "") for row in rows]
        expected_labels = list(LOCATION_MAPPING)
        if actual_labels != expected_labels:
            raise RuntimeError(f"Unexpected location labels for {month_label}: {actual_labels!r}")


def make_placeholders(values: Iterable[str]) -> str:
    return ", ".join("?" for _ in values)


def fetch_counts(
    conn: sqlite3.Connection,
    locations: list[str],
    start_at: str,
    end_at: str,
    extra_where: str = "",
    extra_params: tuple[object, ...] = (),
) -> TrafficCounts:
    placeholders = make_placeholders(locations)
    sql = f"""
        SELECT COUNT(*) AS total_count,
               SUM(CASE WHEN IS_REMICON(vehicle_number) = 1 THEN 1 ELSE 0 END) AS remicon_count
        FROM {TABLE_NAME}
        WHERE collected_at >= ?
          AND collected_at < ?
          AND location_name IN ({placeholders})
          {extra_where}
    """
    params: tuple[object, ...] = (start_at, end_at, *locations, *extra_params)
    total, remicon = conn.execute(sql, params).fetchone()
    return TrafficCounts(total=int(total or 0), remicon=int(remicon or 0))


def fetch_metrics(
    conn: sqlite3.Connection, month_label: str, locations: list[str]
) -> TrafficMetrics:
    start_at, end_at = MONTH_BOUNDS[month_label]
    holiday_values = tuple(day.isoformat() for day in HOLIDAYS)
    holiday_placeholders = make_placeholders(holiday_values)
    weekday_where = f"""
          AND CAST(strftime('%w', collected_at) AS INTEGER) NOT IN (0, 6)
          AND date(collected_at) NOT IN ({holiday_placeholders})
    """
    peak_where = """
          AND (
                (time(collected_at) >= '07:00:00' AND time(collected_at) < '09:00:00')
             OR (time(collected_at) >= '17:00:00' AND time(collected_at) < '19:00:00')
          )
    """
    return TrafficMetrics(
        monthly=fetch_counts(conn, locations, start_at, end_at),
        weekday=fetch_counts(
            conn,
            locations,
            start_at,
            end_at,
            extra_where=weekday_where,
            extra_params=holiday_values,
        ),
        peak=fetch_counts(
            conn,
            locations,
            start_at,
            end_at,
            extra_where=weekday_where + peak_where,
            extra_params=holiday_values,
        ),
    )


def write_counts(ws: Worksheet, row: int, metrics: TrafficMetrics) -> None:
    values = [
        metrics.monthly.total,
        metrics.monthly.remicon,
        metrics.monthly.ratio,
        metrics.weekday.total,
        metrics.weekday.remicon,
        metrics.weekday.ratio,
        metrics.peak.total,
        metrics.peak.remicon,
        metrics.peak.ratio,
    ]
    for col_offset, value in enumerate(values, start=3):
        cell = ws.cell(row=row, column=col_offset)
        cell.value = value
        cell.number_format = "0.00" if col_offset in {5, 8, 11} else "#,##0"


def fill_template(db_path: Path, template_path: Path, output_path: Path) -> None:
    if not template_path.exists():
        raise FileNotFoundError(f"Template file does not exist: {template_path.resolve()}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    conn = connect_readonly(db_path)
    try:
        validate_database(conn)
        validate_locations(conn)
        wb = load_workbook(output_path)
        if SHEET_NAME not in wb.sheetnames:
            raise RuntimeError(f"Missing required sheet: {SHEET_NAME}")
        ws = wb[SHEET_NAME]
        validate_template(ws)

        for month_label, rows in MONTH_ROWS.items():
            for row in rows:
                template_location = str(ws[f"B{row}"].value or "")
                metrics = fetch_metrics(
                    conn,
                    month_label,
                    LOCATION_MAPPING[template_location],
                )
                write_counts(ws, row, metrics)

        wb.save(output_path)
        wb.close()
    finally:
        conn.close()

    validate_output(output_path)


def validate_output(output_path: Path) -> None:
    wb = load_workbook(output_path, data_only=True)
    try:
        ws = wb[SHEET_NAME]
        empty_cells = [
            cell.coordinate
            for row in ws.iter_rows(min_row=4, max_row=19, min_col=3, max_col=11)
            for cell in row
            if cell.value is None
        ]
        if empty_cells:
            raise RuntimeError("Output has empty data cells: " + ", ".join(empty_cells))

        for row in ws.iter_rows(min_row=4, max_row=19, min_col=3, max_col=11):
            for cell in row:
                expected_format = "0.00" if cell.column in {5, 8, 11} else "#,##0"
                if cell.number_format != expected_format:
                    raise RuntimeError(
                        f"Unexpected number format {cell.coordinate}: {cell.number_format!r}"
                    )
    finally:
        wb.close()


def main() -> int:
    args = parse_args()
    fill_template(args.db, args.template, args.output)
    print(f"output={args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
