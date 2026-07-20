#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Update the Samjeong-dong remicon workbook from the read-only SQLite source."""

from __future__ import annotations

import argparse
import re
import sqlite3
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_DB_PATH = BASE_DIR / "00_Data" / "삼정동_레미콘" / "삼정동_레미콘_데이터.db"
DEFAULT_WORKBOOK_PATH = (
    BASE_DIR / "02_Result" / "09_기타" / "삼정동_레미콘_교통량_지역명14포함.xlsx"
)

TABLE_NAME = "vehicle_detection"
TEMPLATE_SHEET = "작성서식"
AVERAGE_SHEET = "월별 일평균"
DETAIL_COLUMNS = [
    "id",
    "collected_at",
    "registered_at",
    "device_id",
    "location_name",
    "lane",
    "lane_direction_name",
    "vehicle_number",
    "owner_registered_area",
    "source_file",
    "source_sheet",
]
REMICON_PATTERNS = (
    re.compile(r"^014[가-힣]"),
    re.compile(r"^[가-힣]{2}14[가-힣][0-9]\*{3}$"),
)
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
HOLIDAYS = {date(2026, 5, 1), date(2026, 5, 5), date(2026, 5, 25), date(2026, 6, 3)}
MONTHS = (("5월", date(2026, 5, 1), date(2026, 6, 1)), ("6월", date(2026, 6, 1), date(2026, 7, 1)))
LOCATION_GROUPS = {
    "박촌교 삼거리": [
        "박촌교 삼거리[남향]",
        "박촌교 삼거리[북향]",
        "박촌교 삼거리[서향(순방향)]",
        "박촌교 삼거리[서향(역방향)]",
    ],
    "봉오고가교사거리": ["봉오고가교 사거리[남향]", "봉오고가교 사거리[서향]"],
    "삼정고가삼거리": ["삼정고가 삼거리[동향]", "삼정고가 삼거리[서향]"],
    "삼정교사거리": ["삼정교 사거리[북동향]"],
    "산업길사거리": ["산업길 사거리[남향]", "산업길 사거리[북향]"],
    "봉오대로사거리": ["봉오대로 사거리[동향]", "봉오대로 사거리[서향]"],
    "자동차검사소": ["자동차검사소"],
    "삼정동 320-1": ["삼정동320-1 부천IC"],
}


@dataclass(frozen=True)
class Counts:
    total: int
    remicon: int


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--verify-only", action="store_true")
    return parser.parse_args()


def is_remicon(value: object) -> int:
    return int(any(pattern.fullmatch(str(value or "")) for pattern in REMICON_PATTERNS))


def detail_sheet_name(location_name: str) -> str:
    return INVALID_SHEET_CHARS.sub("_", location_name).strip()[:31]


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    resolved = db_path.resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"Database file does not exist: {resolved}")
    conn = sqlite3.connect(f"file:{resolved.as_posix()}?mode=ro", uri=True)
    conn.create_function("IS_REMICON", 1, is_remicon)
    return conn


def validate_database(conn: sqlite3.Connection) -> None:
    columns = {str(row[1]) for row in conn.execute(f"PRAGMA table_info({TABLE_NAME})")}
    required = set(DETAIL_COLUMNS)
    if missing := sorted(required - columns):
        raise RuntimeError("Missing required columns: " + ", ".join(missing))


def day_count(start: date, end: date, criterion: str, group_name: str) -> int:
    dates: list[date] = []
    current = start
    while current < end:
        eligible = criterion == "월별" or (current.weekday() < 5 and current not in HOLIDAYS)
        if group_name == "자동차검사소" and date(2026, 5, 20) <= current <= date(2026, 5, 25):
            eligible = False
        if eligible:
            dates.append(current)
        current += timedelta(days=1)
    return len(dates)


def fetch_counts(
    conn: sqlite3.Connection, locations: list[str], start: date, end: date, criterion: str
) -> Counts:
    placeholders = ", ".join("?" for _ in locations)
    conditions = ""
    params: list[object] = [
        f"{start.isoformat()} 00:00:00",
        f"{end.isoformat()} 00:00:00",
        *locations,
    ]
    if criterion != "월별":
        holiday_placeholders = ", ".join("?" for _ in HOLIDAYS)
        conditions += f" AND CAST(strftime('%w', collected_at) AS INTEGER) NOT IN (0, 6) AND date(collected_at) NOT IN ({holiday_placeholders})"
        params.extend(day.isoformat() for day in sorted(HOLIDAYS))
    if criterion == "평일 첨두시":
        conditions += " AND ((time(collected_at) >= '07:00:00' AND time(collected_at) < '09:00:00') OR (time(collected_at) >= '17:00:00' AND time(collected_at) < '19:00:00'))"
    total, remicon = conn.execute(
        f"SELECT COUNT(*), SUM(IS_REMICON(vehicle_number)) FROM {TABLE_NAME} WHERE collected_at >= ? AND collected_at < ? AND location_name IN ({placeholders}){conditions}",
        params,
    ).fetchone()
    return Counts(int(total or 0), int(remicon or 0))


def round_two(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def write_average_sheet(workbook: Any, conn: sqlite3.Connection) -> None:
    if AVERAGE_SHEET in workbook.sheetnames:
        del workbook[AVERAGE_SHEET]
    ws = workbook.create_sheet(AVERAGE_SHEET, 1)
    headers = [
        "월",
        "교차로 묶음",
        "기준",
        "분모(일)",
        "총 통행량 일평균",
        "레미콘 통행량 일평균",
        "레미콘 비율(%)",
    ]
    ws.append(headers)
    for month_label, start, end in MONTHS:
        for group_name, locations in LOCATION_GROUPS.items():
            for criterion in ("월별", "평일", "평일 첨두시"):
                divisor = day_count(start, end, criterion, group_name)
                counts = fetch_counts(conn, locations, start, end, criterion)
                total_average = round_two(Decimal(counts.total) / Decimal(divisor))
                remicon_average = round_two(Decimal(counts.remicon) / Decimal(divisor))
                ratio = (
                    round_two(Decimal(counts.remicon) * Decimal("100") / Decimal(counts.total))
                    if counts.total
                    else 0.0
                )
                ws.append(
                    [
                        month_label,
                        group_name,
                        criterion,
                        divisor,
                        total_average,
                        remicon_average,
                        ratio,
                    ]
                )
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=7):
        row[0].number_format = "0"
        for cell in row[1:]:
            cell.number_format = "0.00"
    widths = (10, 22, 16, 12, 22, 24, 18)
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def load_detail_rows(conn: sqlite3.Connection) -> dict[str, list[dict[str, Any]]]:
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        f"SELECT {', '.join(DETAIL_COLUMNS)} FROM {TABLE_NAME} WHERE collected_at >= ? AND collected_at < ? AND vehicle_number LIKE '%14%' ORDER BY location_name, collected_at, id",
        ("2026-05-01 00:00:00", "2026-07-01 00:00:00"),
    )
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        if is_remicon(row["vehicle_number"]):
            grouped[str(row["location_name"] or "location_name_없음")].append(
                {column: row[column] for column in DETAIL_COLUMNS}
            )
    return dict(sorted(grouped.items()))


def update_log_sheets(workbook: Any, grouped_rows: dict[str, list[dict[str, Any]]]) -> None:
    expected_names = {detail_sheet_name(location_name) for location_name in grouped_rows}
    actual_names = {
        sheet_name
        for sheet_name in workbook.sheetnames
        if [cell.value for cell in workbook[sheet_name][1]] == DETAIL_COLUMNS
    }
    if actual_names != expected_names:
        raise RuntimeError(
            f"Detail sheet mismatch: expected {sorted(expected_names)}, got {sorted(actual_names)}; "
            f"workbook sheets={workbook.sheetnames}"
        )
    for location_name, rows in grouped_rows.items():
        ws = workbook[detail_sheet_name(location_name)]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for row in rows:
            ws.append([row[column] for column in DETAIL_COLUMNS])
        ws.auto_filter.ref = ws.dimensions


def validate_workbook(workbook_path: Path, grouped_rows: dict[str, list[dict[str, Any]]]) -> None:
    wb = load_workbook(workbook_path, read_only=False, data_only=True)
    try:
        if AVERAGE_SHEET not in wb.sheetnames:
            raise RuntimeError("Monthly average sheet is missing")
        ws = wb[AVERAGE_SHEET]
        if ws.max_row != 49 or ws.auto_filter.ref != ws.dimensions:
            raise RuntimeError("Monthly average sheet shape or filter is invalid")
        divisors = {
            (row[0], row[1], row[2]): row[3] for row in ws.iter_rows(min_row=2, values_only=True)
        }
        if divisors[("5월", "자동차검사소", "월별")] != 25:
            raise RuntimeError("Automobile inspection May monthly divisor must be 25")
        if any(
            divisors[("5월", group, "월별")] != 31
            for group in LOCATION_GROUPS
            if group != "자동차검사소"
        ):
            raise RuntimeError("May monthly divisors must be 31")
        if any(divisors[("6월", group, "월별")] != 30 for group in LOCATION_GROUPS):
            raise RuntimeError("June monthly divisors must be 30")
        if (
            divisors[("5월", "자동차검사소", "평일")] != 15
            or divisors[("5월", "박촌교 삼거리", "평일")] != 18
            or divisors[("6월", "박촌교 삼거리", "평일")] != 21
        ):
            raise RuntimeError("Weekday divisor values are invalid")
        for location_name, expected_rows in grouped_rows.items():
            sheet_name = detail_sheet_name(location_name)
            ws = wb[sheet_name]
            if ws.auto_filter.ref != ws.dimensions or ws.max_row - 1 != len(expected_rows):
                raise RuntimeError(f"Detail sheet validation failed: {sheet_name}")
            previous: tuple[str, int] | None = None
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not is_remicon(row[7]):
                    raise RuntimeError(f"Unexpected vehicle number in {sheet_name}")
                key = (str(row[1]), int(row[0]))
                if previous is not None and key < previous:
                    raise RuntimeError(f"Unsorted detail sheet: {sheet_name}")
                previous = key
    finally:
        wb.close()


def validate_average_values(workbook_path: Path, conn: sqlite3.Connection) -> None:
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        ws = wb[AVERAGE_SHEET]
        for row in ws.iter_rows(min_row=2, values_only=True):
            month_label, group_name, criterion, divisor, total_avg, remicon_avg, ratio = row
            start, end = next((start, end) for label, start, end in MONTHS if label == month_label)
            expected_divisor = day_count(start, end, str(criterion), str(group_name))
            counts = fetch_counts(
                conn, LOCATION_GROUPS[str(group_name)], start, end, str(criterion)
            )
            expected = (
                expected_divisor,
                round_two(Decimal(counts.total) / Decimal(expected_divisor)),
                round_two(Decimal(counts.remicon) / Decimal(expected_divisor)),
                round_two(Decimal(counts.remicon) * Decimal("100") / Decimal(counts.total))
                if counts.total
                else 0.0,
            )
            if (divisor, total_avg, remicon_avg, ratio) != expected:
                raise RuntimeError(
                    f"Monthly average DB reaggregation mismatch: {row} != {expected}"
                )
    finally:
        wb.close()


def main() -> int:
    args = parse_args()
    if not args.workbook.is_file():
        raise FileNotFoundError(f"Workbook file does not exist: {args.workbook.resolve()}")
    conn = connect_readonly(args.db)
    try:
        validate_database(conn)
        grouped_rows = load_detail_rows(conn)
        if len(grouped_rows) != 15:
            raise RuntimeError(f"Expected 15 detail locations, got {len(grouped_rows)}")
        if not args.verify_only:
            workbook = load_workbook(args.workbook)
            try:
                if TEMPLATE_SHEET not in workbook.sheetnames:
                    raise RuntimeError(f"Missing required sheet: {TEMPLATE_SHEET}")
                write_average_sheet(workbook, conn)
                update_log_sheets(workbook, grouped_rows)
                workbook.save(args.workbook)
            finally:
                workbook.close()
        validate_workbook(args.workbook, grouped_rows)
        validate_average_values(args.workbook, conn)
    finally:
        conn.close()
    print(f"workbook={args.workbook.resolve()}")
    print(f"detail_rows={sum(len(rows) for rows in grouped_rows.values())}")
    print(f"detail_sheets={len(grouped_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
