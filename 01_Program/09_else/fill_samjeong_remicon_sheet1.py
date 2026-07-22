#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fill the Samjeong remicon Sheet1 traffic summary from SQLite data."""

from __future__ import annotations

import argparse
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re
import sqlite3

from openpyxl import load_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DB_PATH = PROJECT_ROOT / "00_Data" / "삼정동_레미콘" / "삼정동_레미콘_데이터.db"
DEFAULT_WORKBOOK_PATH = PROJECT_ROOT / "02_Result" / "09_기타" / "삼정동_레미콘_교통량_수정_v1.xlsx"
SHEET_NAME = "Sheet1"
TABLE_NAME = "vehicle_detection"
HOLIDAYS = {date(2026, 5, 1), date(2026, 5, 5), date(2026, 5, 25), date(2026, 6, 3)}
PEAK_HOURS = (7, 8, 17, 18)
REMICON_PATTERNS = (
    re.compile(r"^014[가-힣][0-9]\*{3}$"),
    re.compile(r"^[가-힣]{2,}14[가-힣][0-9]\*{3}$"),
)
LOCATION_BY_DIRECTION = {
    ("박촌교 삼거리", "남향"): "박촌교 삼거리[남향]",
    ("박촌교 삼거리", "북향"): "박촌교 삼거리[북향]",
    ("박촌교 삼거리", "서향(순방향)"): "박촌교 삼거리[서향(순방향)]",
    ("박촌교 삼거리", "서향(역방향)"): "박촌교 삼거리[서향(역방향)]",
    ("봉오고가교 사거리", "남향"): "봉오고가교 사거리[남향]",
    ("봉오고가교 사거리", "서향"): "봉오고가교 사거리[서향]",
    ("삼정고가 삼거리", "동향"): "삼정고가 삼거리[동향]",
    ("삼정고가 삼거리", "서향"): "삼정고가 삼거리[서향]",
    ("삼정교 사거리", "북동향"): "삼정교 사거리[북동향]",
    ("산업길 사거리", "남향"): "산업길 사거리[남향]",
    ("산업길 사거리", "북향"): "산업길 사거리[북향]",
    ("봉오대로 사거리", "동향"): "봉오대로 사거리[동향]",
    ("봉오대로 사거리", "서향"): "봉오대로 사거리[서향]",
    ("자동차검사소", "-"): "자동차검사소",
    ("삼정동 320-1", "-"): "삼정동320-1 부천IC",
}


@dataclass(frozen=True)
class Period:
    start: date
    end: date
    daily_divisor: int
    weekday_divisor: int


@dataclass(frozen=True)
class Counts:
    total: int
    remicon: int

    def rounded_average(self, divisor: int) -> tuple[int, int, float]:
        total_average = Decimal(self.total) / divisor
        remicon_average = Decimal(self.remicon) / divisor
        ratio = Decimal("0") if self.total == 0 else Decimal(self.remicon) * 100 / self.total
        return (
            int(total_average.quantize(Decimal("1"), rounding=ROUND_HALF_UP)),
            int(remicon_average.quantize(Decimal("1"), rounding=ROUND_HALF_UP)),
            float(ratio.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)),
        )


PERIODS = (
    (range(4, 19), Period(date(2026, 5, 1), date(2026, 6, 1), 31, 18)),
    (range(19, 34), Period(date(2026, 6, 1), date(2026, 7, 1), 30, 21)),
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--verify-only", action="store_true")
    return parser.parse_args()


def is_remicon(vehicle_number: object) -> int:
    return int(any(pattern.fullmatch(str(vehicle_number or "")) for pattern in REMICON_PATTERNS))


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(f"file:{db_path.resolve().as_posix()}?mode=ro", uri=True)
    connection.create_function("IS_REMICON", 1, is_remicon)
    return connection


def weekday_dates(period: Period) -> list[str]:
    dates: list[str] = []
    current = period.start
    while current < period.end:
        if current.weekday() < 5 and current not in HOLIDAYS:
            dates.append(current.isoformat())
        current += timedelta(days=1)
    if len(dates) != period.weekday_divisor:
        raise RuntimeError(f"Unexpected weekday count: {len(dates)}")
    return dates


def query_counts(
    connection: sqlite3.Connection, locations: Iterable[str], period: Period
) -> tuple[dict[str, Counts], dict[str, Counts], dict[str, Counts]]:
    location_list = tuple(locations)
    placeholders = ", ".join("?" for _ in location_list)
    weekday_list = weekday_dates(period)
    weekday_placeholders = ", ".join("?" for _ in weekday_list)
    sql = f"""
        SELECT location_name,
               COUNT(*) AS monthly_total,
               SUM(IS_REMICON(vehicle_number)) AS monthly_remicon,
               SUM(CASE WHEN date(collected_at) IN ({weekday_placeholders}) THEN 1 ELSE 0 END),
               SUM(CASE WHEN date(collected_at) IN ({weekday_placeholders})
                         THEN IS_REMICON(vehicle_number) ELSE 0 END),
               SUM(CASE WHEN date(collected_at) IN ({weekday_placeholders})
                            AND CAST(strftime('%H', collected_at) AS INTEGER) IN (7, 8, 17, 18)
                         THEN 1 ELSE 0 END),
               SUM(CASE WHEN date(collected_at) IN ({weekday_placeholders})
                            AND CAST(strftime('%H', collected_at) AS INTEGER) IN (7, 8, 17, 18)
                         THEN IS_REMICON(vehicle_number) ELSE 0 END)
        FROM {TABLE_NAME}
        WHERE collected_at >= ? AND collected_at < ?
          AND location_name IN ({placeholders})
        GROUP BY location_name
    """
    values = (
        *weekday_list,
        *weekday_list,
        *weekday_list,
        *weekday_list,
        period.start.isoformat(),
        period.end.isoformat(),
        *location_list,
    )
    monthly: dict[str, Counts] = {}
    weekday: dict[str, Counts] = {}
    peak: dict[str, Counts] = {}
    for row in connection.execute(sql, values):
        location, *counts = row
        monthly[str(location)] = Counts(int(counts[0] or 0), int(counts[1] or 0))
        weekday[str(location)] = Counts(int(counts[2] or 0), int(counts[3] or 0))
        peak[str(location)] = Counts(int(counts[4] or 0), int(counts[5] or 0))
    missing = set(location_list) - monthly.keys()
    if missing:
        raise RuntimeError(f"No DB data for locations: {sorted(missing)}")
    return monthly, weekday, peak


def row_locations(worksheet: object, rows: range) -> dict[int, str]:
    current_intersection = ""
    result: dict[int, str] = {}
    for row in rows:
        intersection = worksheet.cell(row=row, column=2).value
        if intersection:
            current_intersection = str(intersection)
        direction = str(worksheet.cell(row=row, column=3).value or "")
        try:
            result[row] = LOCATION_BY_DIRECTION[(current_intersection, direction)]
        except KeyError as error:
            raise RuntimeError(
                f"Unmapped worksheet row {row}: {current_intersection} / {direction}"
            ) from error
    return result


def expected_values(workbook_path: Path, db_path: Path) -> dict[int, tuple[object, ...]]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    try:
        worksheet = workbook[SHEET_NAME]
        row_map = {
            row: location
            for rows, _ in PERIODS
            for row, location in row_locations(worksheet, rows).items()
        }
    finally:
        workbook.close()
    results: dict[int, tuple[object, ...]] = {}
    connection = connect_readonly(db_path)
    try:
        for rows, period in PERIODS:
            monthly, weekday, peak = query_counts(connection, row_map.values(), period)
            for row in rows:
                location = row_map[row]
                results[row] = (
                    *monthly[location].rounded_average(period.daily_divisor),
                    *weekday[location].rounded_average(period.weekday_divisor),
                    *peak[location].rounded_average(period.weekday_divisor),
                )
    finally:
        connection.close()
    return results


def apply_or_verify(
    workbook_path: Path, results: dict[int, tuple[object, ...]], verify_only: bool
) -> None:
    workbook = load_workbook(workbook_path, data_only=False)
    try:
        worksheet = workbook[SHEET_NAME]
        for row, values in results.items():
            actual = tuple(worksheet.cell(row=row, column=column).value for column in range(6, 15))
            if verify_only:
                if actual != values:
                    raise RuntimeError(f"Mismatch in Sheet1 row {row}: {actual!r} != {values!r}")
                continue
            for column in (4, 5):
                worksheet.cell(row=row, column=column).value = None
            for column, value in enumerate(values, start=6):
                cell = worksheet.cell(row=row, column=column)
                cell.value = value
                cell.number_format = "0.00" if column in (8, 11, 14) else "#,##0"
        if not verify_only:
            workbook.save(workbook_path)
    finally:
        workbook.close()


def main() -> int:
    args = parse_args()
    results = expected_values(args.workbook, args.db)
    apply_or_verify(args.workbook, results, args.verify_only)
    for row, values in results.items():
        print(row, *values, sep="\t")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
