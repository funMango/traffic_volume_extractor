#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export Samjeong-dong remicon 014 Korean-prefixed vehicle detections by location."""

from __future__ import annotations

import argparse
from calendar import monthrange
from collections import Counter
import re
import sqlite3
from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parents[2]
DB_PATH = BASE_DIR / "00_Data" / "삼정동_레미콘_데이터.db"
OUTPUT_DIR = BASE_DIR / "02_Result" / "09_기타"
OUTPUT_PATH = OUTPUT_DIR / "삼정동_레미콘_014차량_location별.xlsx"

TABLE_NAME = "vehicle_detection"
EXPECTED_TOTAL_ROWS = 8793
EXPECTED_LOCATION_COUNT = 5
VEHICLE_NUMBER_PATTERN = re.compile(r"^014[가-힣]")
INVALID_SHEET_CHARS = re.compile(r"[\[\]\:\*\?\/\\]")
MONTH_PATTERN = re.compile(r"^\d{4}-\d{2}$")

ANALYSIS_SHEET_NAME = "월평균분석"
ROAD_LOCATION_GROUPS = [
    ("신흥로", ["산업길 사거리[남향]", "삼정동320-1 부천IC"]),
    (
        "오정로",
        ["삼정고가 삼거리[동향]", "삼정교 사거리[북동향]", "자동차검사소"],
    ),
]

SOURCE_COLUMNS = [
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


def make_analysis_header(months: list[str]) -> list[str]:
    header = ["도로", "location_name"]
    for month in months:
        header.extend([f"{month} 총대수", f"{month} 월일수", f"{month} 일평균"])
    header.extend(["전체 총대수", "전체 일수", "전체 일평균"])
    return header


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export rows whose vehicle_number starts with 014 and a Korean 4th character."
    )
    parser.add_argument("--db", type=Path, default=DB_PATH, help="Input SQLite DB path")
    parser.add_argument("--output", type=Path, default=OUTPUT_PATH, help="Output xlsx path")
    parser.add_argument(
        "--skip-expected-check",
        action="store_true",
        help="Do not fail when the current row/location count differs from the known plan.",
    )
    return parser.parse_args()


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    resolved = db_path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"DB file does not exist: {resolved}")
    uri = f"file:{resolved.as_posix()}?mode=ro"
    return sqlite3.connect(uri, uri=True)


def validate_table_columns(conn: sqlite3.Connection) -> None:
    rows = conn.execute(f"PRAGMA table_info({TABLE_NAME})").fetchall()
    actual_columns = {str(row[1]) for row in rows}
    missing_columns = [column for column in SOURCE_COLUMNS if column not in actual_columns]
    if missing_columns:
        raise RuntimeError("Missing required columns: " + ", ".join(missing_columns))


def load_filtered_rows(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    sql = f"""
        SELECT {", ".join(SOURCE_COLUMNS)}
        FROM {TABLE_NAME}
        WHERE vehicle_number LIKE '014%'
        ORDER BY collected_at ASC, id ASC
    """
    conn.row_factory = sqlite3.Row
    rows = []
    for row in conn.execute(sql):
        vehicle_number = str(row["vehicle_number"] or "")
        if VEHICLE_NUMBER_PATTERN.match(vehicle_number):
            rows.append({column: row[column] for column in SOURCE_COLUMNS})
    return rows


def group_by_location(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        location_name = str(row["location_name"] or "location_name_없음")
        grouped[location_name].append(row)
    return dict(sorted(grouped.items(), key=lambda item: item[0]))


def extract_collected_month(row: dict[str, Any]) -> str:
    collected_at = str(row["collected_at"] or "")
    month = collected_at[:7]
    if not MONTH_PATTERN.match(month):
        raise RuntimeError(f"Invalid collected_at month value: {collected_at}")
    return month


def count_month_days(month: str) -> int:
    year, month_number = (int(part) for part in month.split("-"))
    return monthrange(year, month_number)[1]


def get_analysis_months(grouped_rows: dict[str, list[dict[str, Any]]]) -> list[str]:
    months = {
        extract_collected_month(row)
        for location_rows in grouped_rows.values()
        for row in location_rows
    }
    if not months:
        raise RuntimeError("No collected_at months found")
    return sorted(months)


def get_road_location_map() -> dict[str, str]:
    road_by_location: dict[str, str] = {}
    for road_name, location_names in ROAD_LOCATION_GROUPS:
        for location_name in location_names:
            if location_name in road_by_location:
                raise RuntimeError(f"Duplicate road location mapping: {location_name}")
            road_by_location[location_name] = road_name
    return road_by_location


def validate_location_group_mapping(grouped_rows: dict[str, list[dict[str, Any]]]) -> None:
    expected_locations = set(get_road_location_map())
    actual_locations = set(grouped_rows)
    unmapped_locations = sorted(actual_locations - expected_locations)
    missing_locations = sorted(expected_locations - actual_locations)
    if unmapped_locations:
        raise RuntimeError("Unmapped locations: " + ", ".join(unmapped_locations))
    if missing_locations:
        raise RuntimeError("Missing expected locations: " + ", ".join(missing_locations))


def build_analysis_row(
    road_name: str,
    location_name: str,
    month_counts: Counter[str],
    months: list[str],
    month_days: dict[str, int],
) -> list[Any]:
    row: list[Any] = [road_name, location_name]
    total_count = 0
    total_days = 0
    for month in months:
        month_count = month_counts[month]
        days = month_days[month]
        row.extend([month_count, days, month_count / days])
        total_count += month_count
        total_days += days
    row.extend([total_count, total_days, total_count / total_days])
    return row


def build_analysis_rows(
    grouped_rows: dict[str, list[dict[str, Any]]],
) -> tuple[list[str], list[list[Any]]]:
    validate_location_group_mapping(grouped_rows)
    months = get_analysis_months(grouped_rows)
    month_days = {month: count_month_days(month) for month in months}
    rows: list[list[Any]] = []

    location_month_counts = {
        location_name: Counter(extract_collected_month(row) for row in location_rows)
        for location_name, location_rows in grouped_rows.items()
    }

    for road_name, location_names in ROAD_LOCATION_GROUPS:
        subtotal_counts: Counter[str] = Counter()
        for location_name in location_names:
            month_counts = location_month_counts[location_name]
            subtotal_counts.update(month_counts)
            rows.append(
                build_analysis_row(
                    road_name,
                    location_name,
                    month_counts,
                    months,
                    month_days,
                )
            )
        rows.append(
            build_analysis_row(
                road_name,
                "소계",
                subtotal_counts,
                months,
                month_days,
            )
        )

    return make_analysis_header(months), rows


def make_unique_sheet_name(raw_name: str, used_names: set[str]) -> str:
    clean_name = INVALID_SHEET_CHARS.sub("_", raw_name).strip() or "Sheet"
    base_name = clean_name[:31]
    name = base_name
    suffix = 1
    while name in used_names:
        tail = f"_{suffix}"
        name = base_name[: 31 - len(tail)] + tail
        suffix += 1
    used_names.add(name)
    return name


def style_header(ws: Worksheet, fill: PatternFill, font: Font) -> None:
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font


def fit_columns(ws: Worksheet, column_names: list[str]) -> None:
    for column_index, column_name in enumerate(column_names, start=1):
        letter = get_column_letter(column_index)
        max_length = len(column_name)
        for cell in ws[letter][1:]:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[letter].width = min(max(max_length + 2, 10), 42)


def write_analysis_sheet(
    wb: Workbook,
    grouped_rows: dict[str, list[dict[str, Any]]],
    header_fill: PatternFill,
    header_font: Font,
) -> None:
    header, rows = build_analysis_rows(grouped_rows)
    ws = wb.create_sheet(ANALYSIS_SHEET_NAME, 0)
    ws.append(header)
    for row in rows:
        ws.append(row)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    style_header(ws, header_fill, header_font)

    for row in ws.iter_rows(min_row=2):
        if row[1].value == "소계":
            for cell in row:
                cell.font = Font(bold=True)

    for column_index, column_name in enumerate(header, start=1):
        letter = get_column_letter(column_index)
        if column_name.endswith("일평균"):
            for cell in ws[letter][1:]:
                cell.number_format = "0.00"
        elif (
            column_name.endswith("총대수")
            or column_name.endswith("월일수")
            or column_name == "전체 일수"
        ):
            for cell in ws[letter][1:]:
                cell.number_format = "0"

    fit_columns(ws, header)


def write_workbook(grouped_rows: dict[str, list[dict[str, Any]]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)
    used_sheet_names: set[str] = set()

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    write_analysis_sheet(wb, grouped_rows, header_fill, header_font)

    for location_name, rows in grouped_rows.items():
        sheet_name = make_unique_sheet_name(location_name, used_sheet_names)
        ws = wb.create_sheet(sheet_name)
        ws.append(SOURCE_COLUMNS)
        for row in rows:
            ws.append([row[column] for column in SOURCE_COLUMNS])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        style_header(ws, header_fill, header_font)
        fit_columns(ws, SOURCE_COLUMNS)

    wb.save(output_path)


def validate_rows(
    rows: list[dict[str, Any]], grouped_rows: dict[str, list[dict[str, Any]]]
) -> None:
    bad_values = [
        str(row["vehicle_number"])
        for row in rows
        if not VEHICLE_NUMBER_PATTERN.match(str(row["vehicle_number"] or ""))
    ]
    if bad_values:
        raise RuntimeError("Unexpected vehicle_number values: " + ", ".join(bad_values[:10]))

    for location_name, location_rows in grouped_rows.items():
        collected_at_values = [str(row["collected_at"] or "") for row in location_rows]
        if collected_at_values != sorted(collected_at_values):
            raise RuntimeError(f"Rows are not sorted by collected_at: {location_name}")


def validate_saved_workbook(
    output_path: Path, grouped_rows: dict[str, list[dict[str, Any]]]
) -> None:
    wb = load_workbook(output_path, read_only=True, data_only=True)
    expected_sheet_count = len(grouped_rows) + 1
    if len(wb.sheetnames) != expected_sheet_count:
        raise RuntimeError(
            f"Saved workbook sheet count mismatch: expected {expected_sheet_count}, got {len(wb.sheetnames)}"
        )
    if wb.sheetnames[0] != ANALYSIS_SHEET_NAME:
        raise RuntimeError(
            f"First sheet mismatch: expected {ANALYSIS_SHEET_NAME}, got {wb.sheetnames[0]}"
        )

    expected_counts = sorted(len(rows) for rows in grouped_rows.values())
    actual_counts = sorted(wb[sheet_name].max_row - 1 for sheet_name in wb.sheetnames[1:])
    if actual_counts != expected_counts:
        raise RuntimeError(
            f"Saved workbook row counts mismatch: {actual_counts} != {expected_counts}"
        )
    wb.close()


def main() -> None:
    args = parse_args()
    conn = connect_readonly(args.db)
    try:
        validate_table_columns(conn)
        rows = load_filtered_rows(conn)
    finally:
        conn.close()

    grouped_rows = group_by_location(rows)
    validate_rows(rows, grouped_rows)

    if not args.skip_expected_check:
        if len(rows) != EXPECTED_TOTAL_ROWS:
            raise RuntimeError(f"Expected {EXPECTED_TOTAL_ROWS:,} rows, got {len(rows):,}")
        if len(grouped_rows) != EXPECTED_LOCATION_COUNT:
            raise RuntimeError(
                f"Expected {EXPECTED_LOCATION_COUNT} locations, got {len(grouped_rows)}"
            )

    write_workbook(grouped_rows, args.output)
    validate_saved_workbook(args.output, grouped_rows)

    print(f"output={args.output.resolve()}")
    print(f"total_rows={len(rows)}")
    print(f"sheet_count={len(grouped_rows)}")
    for location_name, location_rows in grouped_rows.items():
        print(f"{location_name}\t{len(location_rows)}")


if __name__ == "__main__":
    main()
