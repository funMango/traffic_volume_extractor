#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Add Samjeong-dong remicon camera log sheets to the existing workbook."""

from __future__ import annotations

import argparse
from collections import defaultdict
from pathlib import Path
import re
import sqlite3
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_DB_PATH = BASE_DIR / "00_Data" / "삼정동_레미콘" / "삼정동_레미콘_데이터.db"
DEFAULT_WORKBOOK_PATH = BASE_DIR / "02_Result" / "09_기타" / "삼정동_레미콘_교통량.xlsx"

TABLE_NAME = "vehicle_detection"
TEMPLATE_SHEET_NAME = "작성서식"
WEEKEND_SUMMARY_SHEET_NAME = "주말집계"
START_AT = "2026-05-01 00:00:00"
END_AT = "2026-07-01 00:00:00"
REMICON_PATTERN = re.compile(r"^014[가-힣]")
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
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
SUMMARY_COLUMNS = [
    "location_name",
    "주말 총 통행량",
    "주말 레미콘 통행량",
    "주말 레미콘 비율(%)",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Add location_name-based remicon log sheets to Samjeong workbook."
    )
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK_PATH)
    return parser.parse_args()


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    resolved = db_path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"DB file does not exist: {resolved}")
    conn = sqlite3.connect(f"file:{resolved.as_posix()}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def validate_schema(conn: sqlite3.Connection) -> None:
    tables = {
        str(row[0])
        for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall()
    }
    if TABLE_NAME not in tables:
        raise RuntimeError(f"Missing required table: {TABLE_NAME}")

    columns = {str(row[1]) for row in conn.execute(f"PRAGMA table_info({TABLE_NAME})").fetchall()}
    missing_columns = [column for column in SOURCE_COLUMNS if column not in columns]
    if missing_columns:
        raise RuntimeError("Missing required columns: " + ", ".join(missing_columns))


def fetch_rows(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    sql = f"""
        SELECT {", ".join(SOURCE_COLUMNS)}
        FROM {TABLE_NAME}
        WHERE collected_at >= ?
          AND collected_at < ?
          AND vehicle_number LIKE '014%'
        ORDER BY location_name ASC, collected_at ASC, id ASC
    """
    rows: list[dict[str, Any]] = []
    for row in conn.execute(sql, (START_AT, END_AT)):
        vehicle_number = str(row["vehicle_number"] or "")
        if REMICON_PATTERN.match(vehicle_number):
            rows.append({column: row[column] for column in SOURCE_COLUMNS})
    return rows


def fetch_weekend_summary(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    total_sql = f"""
        SELECT location_name, COUNT(*) AS total_count
        FROM {TABLE_NAME}
        WHERE collected_at >= ?
          AND collected_at < ?
          AND strftime('%w', collected_at) IN ('0', '6')
        GROUP BY location_name
        ORDER BY location_name ASC
    """
    remicon_sql = f"""
        SELECT location_name, vehicle_number
        FROM {TABLE_NAME}
        WHERE collected_at >= ?
          AND collected_at < ?
          AND strftime('%w', collected_at) IN ('0', '6')
          AND vehicle_number LIKE '014%'
        ORDER BY location_name ASC
    """

    totals = {
        str(row["location_name"] or "location_name_없음"): int(row["total_count"])
        for row in conn.execute(total_sql, (START_AT, END_AT))
    }
    remicon_counts = dict.fromkeys(totals, 0)
    for row in conn.execute(remicon_sql, (START_AT, END_AT)):
        vehicle_number = str(row["vehicle_number"] or "")
        if REMICON_PATTERN.match(vehicle_number):
            location_name = str(row["location_name"] or "location_name_없음")
            remicon_counts[location_name] = remicon_counts.get(location_name, 0) + 1

    summary_rows: list[dict[str, Any]] = []
    for location_name in sorted(totals):
        total_count = totals[location_name]
        remicon_count = remicon_counts.get(location_name, 0)
        ratio = 0.0 if total_count == 0 else round(remicon_count / total_count * 100, 2)
        summary_rows.append(
            {
                "location_name": location_name,
                "주말 총 통행량": total_count,
                "주말 레미콘 통행량": remicon_count,
                "주말 레미콘 비율(%)": ratio,
            }
        )
    return summary_rows


def group_by_location(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        location_name = str(row["location_name"] or "location_name_없음")
        grouped[location_name].append(row)
    return dict(sorted(grouped.items(), key=lambda item: item[0]))


def safe_sheet_base(raw_name: str) -> str:
    return INVALID_SHEET_CHARS.sub("_", raw_name).strip() or "Sheet"


def build_sheet_name_map(locations: list[str], existing_non_log_names: set[str]) -> dict[str, str]:
    used_names = set(existing_non_log_names)
    sheet_names: dict[str, str] = {}
    for location in locations:
        base_name = safe_sheet_base(location)[:31]
        sheet_name = base_name
        suffix = 1
        while sheet_name in used_names:
            tail = f"_{suffix}"
            sheet_name = base_name[: 31 - len(tail)] + tail
            suffix += 1
        used_names.add(sheet_name)
        sheet_names[location] = sheet_name
    return sheet_names


def display_width(value: object) -> int:
    text = "" if value is None else str(value)
    return sum(2 if ord(ch) > 127 else 1 for ch in text)


def style_header(ws: Worksheet) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font


def fit_columns(ws: Worksheet, rows: list[dict[str, Any]]) -> None:
    fit_columns_for_dict_rows(ws, SOURCE_COLUMNS, rows)


def fit_columns_for_dict_rows(
    ws: Worksheet, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    for column_index, column_name in enumerate(columns, start=1):
        max_width = display_width(column_name)
        for row in rows:
            max_width = max(max_width, display_width(row[column_name]))
        ws.column_dimensions[get_column_letter(column_index)].width = min(
            max(max_width + 2, 10), 48
        )


def write_log_sheet(wb: Any, sheet_name: str, rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(SOURCE_COLUMNS)
    for row in rows:
        ws.append([row[column] for column in SOURCE_COLUMNS])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    style_header(ws)
    fit_columns(ws, rows)


def write_weekend_summary_sheet(wb: Any, rows: list[dict[str, Any]]) -> None:
    if WEEKEND_SUMMARY_SHEET_NAME in wb.sheetnames:
        del wb[WEEKEND_SUMMARY_SHEET_NAME]

    insert_index = wb.sheetnames.index(TEMPLATE_SHEET_NAME) + 1
    ws = wb.create_sheet(WEEKEND_SUMMARY_SHEET_NAME, insert_index)
    ws.append(SUMMARY_COLUMNS)
    for row in rows:
        ws.append([row[column] for column in SUMMARY_COLUMNS])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    style_header(ws)
    for row_index in range(2, ws.max_row + 1):
        ws.cell(row=row_index, column=2).number_format = "#,##0"
        ws.cell(row=row_index, column=3).number_format = "#,##0"
        ws.cell(row=row_index, column=4).number_format = "0.00"
    fit_columns_for_dict_rows(ws, SUMMARY_COLUMNS, rows)


def update_workbook(
    workbook_path: Path,
    grouped_rows: dict[str, list[dict[str, Any]]],
    weekend_summary_rows: list[dict[str, Any]],
) -> None:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook file does not exist: {workbook_path.resolve()}")

    wb = load_workbook(workbook_path)
    if TEMPLATE_SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise RuntimeError(f"Missing required sheet: {TEMPLATE_SHEET_NAME}")

    log_sheet_names = set(build_sheet_name_map(list(grouped_rows), {TEMPLATE_SHEET_NAME}).values())
    for sheet_name in list(wb.sheetnames):
        if sheet_name in log_sheet_names:
            del wb[sheet_name]

    write_weekend_summary_sheet(wb, weekend_summary_rows)

    sheet_name_map = build_sheet_name_map(list(grouped_rows), set(wb.sheetnames))
    for location_name, rows in grouped_rows.items():
        write_log_sheet(wb, sheet_name_map[location_name], rows)

    wb.save(workbook_path)
    wb.close()


def validate_rows(
    rows: list[dict[str, Any]], grouped_rows: dict[str, list[dict[str, Any]]]
) -> None:
    for row in rows:
        vehicle_number = str(row["vehicle_number"] or "")
        if not REMICON_PATTERN.match(vehicle_number):
            raise RuntimeError(f"Unexpected vehicle_number: {vehicle_number}")
        collected_at = str(row["collected_at"] or "")
        if not (START_AT <= collected_at < END_AT):
            raise RuntimeError(f"Unexpected collected_at: {collected_at}")

    for location_name, location_rows in grouped_rows.items():
        keys = [(str(row["collected_at"] or ""), int(row["id"])) for row in location_rows]
        if keys != sorted(keys):
            raise RuntimeError(f"Rows are not sorted by collected_at/id: {location_name}")


def validate_weekend_summary_rows(
    conn: sqlite3.Connection, weekend_summary_rows: list[dict[str, Any]]
) -> None:
    distinct_location_count = conn.execute(
        f"""
        SELECT COUNT(*)
        FROM (
            SELECT DISTINCT location_name
            FROM {TABLE_NAME}
            WHERE collected_at >= ?
              AND collected_at < ?
              AND strftime('%w', collected_at) IN ('0', '6')
        )
        """,
        (START_AT, END_AT),
    ).fetchone()[0]
    if len(weekend_summary_rows) != int(distinct_location_count):
        raise RuntimeError(
            "Weekend summary location count mismatch: "
            f"{len(weekend_summary_rows)} != {distinct_location_count}"
        )

    for row in weekend_summary_rows:
        total_count = int(row["주말 총 통행량"])
        remicon_count = int(row["주말 레미콘 통행량"])
        ratio = float(row["주말 레미콘 비율(%)"])
        if remicon_count > total_count:
            raise RuntimeError(f"Weekend remicon count exceeds total count: {row['location_name']}")
        expected_ratio = 0.0 if total_count == 0 else round(remicon_count / total_count * 100, 2)
        if ratio != expected_ratio:
            raise RuntimeError(
                f"Weekend ratio mismatch for {row['location_name']}: {ratio} != {expected_ratio}"
            )


def validate_saved_workbook(
    workbook_path: Path,
    grouped_rows: dict[str, list[dict[str, Any]]],
    weekend_summary_rows: list[dict[str, Any]],
) -> None:
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if TEMPLATE_SHEET_NAME not in wb.sheetnames:
            raise RuntimeError(f"Missing required sheet after save: {TEMPLATE_SHEET_NAME}")
        if WEEKEND_SUMMARY_SHEET_NAME not in wb.sheetnames:
            raise RuntimeError(f"Missing required sheet after save: {WEEKEND_SUMMARY_SHEET_NAME}")

        sheet_name_map = build_sheet_name_map(
            list(grouped_rows),
            set(wb.sheetnames)
            - set(build_sheet_name_map(list(grouped_rows), {TEMPLATE_SHEET_NAME}).values())
            - {WEEKEND_SUMMARY_SHEET_NAME},
        )
        expected_log_names = set(sheet_name_map.values())
        actual_log_names = set(wb.sheetnames) - {TEMPLATE_SHEET_NAME, WEEKEND_SUMMARY_SHEET_NAME}
        if expected_log_names != actual_log_names:
            raise RuntimeError(
                f"Log sheet mismatch: expected {sorted(expected_log_names)}, got {sorted(actual_log_names)}"
            )

        summary_ws = wb[WEEKEND_SUMMARY_SHEET_NAME]
        summary_header = [cell.value for cell in next(summary_ws.iter_rows(min_row=1, max_row=1))]
        if summary_header != SUMMARY_COLUMNS:
            raise RuntimeError(
                f"Unexpected header in {WEEKEND_SUMMARY_SHEET_NAME}: {summary_header}"
            )
        if summary_ws.max_row - 1 != len(weekend_summary_rows):
            raise RuntimeError(
                f"Weekend summary row count mismatch: {summary_ws.max_row - 1} "
                f"!= {len(weekend_summary_rows)}"
            )
        for row_index, expected_row in enumerate(weekend_summary_rows, start=2):
            actual_row = {
                column: summary_ws.cell(row=row_index, column=column_index).value
                for column_index, column in enumerate(SUMMARY_COLUMNS, start=1)
            }
            if actual_row != expected_row:
                raise RuntimeError(
                    f"Weekend summary row mismatch at row {row_index}: "
                    f"{actual_row} != {expected_row}"
                )

        for location_name, sheet_name in sheet_name_map.items():
            ws = wb[sheet_name]
            header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            if header != SOURCE_COLUMNS:
                raise RuntimeError(f"Unexpected header in {sheet_name}: {header}")
            expected_count = len(grouped_rows[location_name])
            actual_count = ws.max_row - 1
            if actual_count != expected_count:
                raise RuntimeError(
                    f"Row count mismatch for {sheet_name}: {actual_count} != {expected_count}"
                )
    finally:
        wb.close()


def main() -> int:
    args = parse_args()
    conn = connect_readonly(args.db)
    try:
        validate_schema(conn)
        rows = fetch_rows(conn)
        weekend_summary_rows = fetch_weekend_summary(conn)
        validate_weekend_summary_rows(conn, weekend_summary_rows)
    finally:
        conn.close()

    grouped_rows = group_by_location(rows)
    validate_rows(rows, grouped_rows)
    update_workbook(args.workbook, grouped_rows, weekend_summary_rows)
    validate_saved_workbook(args.workbook, grouped_rows, weekend_summary_rows)

    print(f"workbook={args.workbook.resolve()}")
    print(f"total_rows={len(rows)}")
    print(f"sheet_count={len(grouped_rows)}")
    print(f"weekend_summary_rows={len(weekend_summary_rows)}")
    for location_name, location_rows in grouped_rows.items():
        print(f"{location_name}\t{len(location_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
