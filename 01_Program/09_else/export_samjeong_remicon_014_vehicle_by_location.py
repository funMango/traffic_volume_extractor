#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export Samjeong-dong remicon 014 Korean-prefixed vehicle detections by location."""

from __future__ import annotations

import argparse
import re
import sqlite3
from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
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


def write_workbook(grouped_rows: dict[str, list[dict[str, Any]]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)
    used_sheet_names: set[str] = set()

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for location_name, rows in grouped_rows.items():
        sheet_name = make_unique_sheet_name(location_name, used_sheet_names)
        ws = wb.create_sheet(sheet_name)
        ws.append(SOURCE_COLUMNS)
        for row in rows:
            ws.append([row[column] for column in SOURCE_COLUMNS])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        for column_index, column_name in enumerate(SOURCE_COLUMNS, start=1):
            letter = get_column_letter(column_index)
            max_length = len(column_name)
            for cell in ws[letter][1:]:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))
            ws.column_dimensions[letter].width = min(max(max_length + 2, 10), 42)

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
    expected_sheet_count = len(grouped_rows)
    if len(wb.sheetnames) != expected_sheet_count:
        raise RuntimeError(
            f"Saved workbook sheet count mismatch: expected {expected_sheet_count}, got {len(wb.sheetnames)}"
        )

    expected_counts = sorted(len(rows) for rows in grouped_rows.values())
    actual_counts = sorted(wb[sheet_name].max_row - 1 for sheet_name in wb.sheetnames)
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
