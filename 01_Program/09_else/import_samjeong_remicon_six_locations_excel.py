#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Import June 2026 raw Excel data for six Samjeong-dong locations into SQLite."""

from __future__ import annotations

import argparse
import sqlite3
import warnings
from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parents[2]
DB_PATH = BASE_DIR / "00_Data" / "삼정동_레미콘_데이터.db"
SOURCE_ROOT = BASE_DIR / "00_Data" / "2605_06_삼정동_레미콘_데이터"
SHEET_NAME = "data_0"
TABLE_NAME = "vehicle_detection"
SUMMARY_TABLE_NAME = "import_summary"
TOTAL_TARGET_ROWS = 2_434_263
EXPECTED_TOTAL_DB_ROWS = 5_779_563
EIGHT_COLUMN_HEADER = (
    "수집일시",
    "등록일시",
    "장비ID",
    "위치명",
    "차선",
    "차선방향명",
    "차량번호",
    "차량소유주등록지",
)
SIX_COLUMN_HEADER = (
    "수집일시",
    "등록일시",
    "장비ID",
    "위치명",
    "차량번호",
    "차량소유주등록지",
)
INSERT_COLUMNS = (
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
)


@dataclass(frozen=True)
class SourceFile:
    directory: str
    file_name: str
    expected_rows: int
    column_count: int


SOURCE_FILES = (
    SourceFile("봉오대로사거리", "260501_260506_봉오대로사거리.xlsx", 92_408, 8),
    SourceFile("봉오대로사거리", "260507_260511_봉오대로사거리.xlsx", 90_229, 8),
    SourceFile("봉오대로사거리", "260512_260515_봉오대로사거리.xlsx", 80_841, 8),
    SourceFile("봉오대로사거리", "260516_260520_봉오대로사거리.xlsx", 87_864, 8),
    SourceFile("봉오대로사거리", "260521_260525_봉오대로사거리.xlsx", 83_728, 8),
    SourceFile("봉오대로사거리", "260526_260530_봉오대로사거리.xlsx", 93_309, 8),
    SourceFile("봉오대로사거리", "260531_봉오대로사거리.xlsx", 12_901, 8),
    SourceFile("봉오대로사거리", "260601_260605_봉오대로사거리.xlsx", 94_449, 8),
    SourceFile("봉오대로사거리", "260606_260611_봉오대로사거리.xlsx", 98_615, 8),
    SourceFile("봉오대로사거리", "260612_260616_봉오대로사거리.xlsx", 88_552, 8),
    SourceFile("봉오대로사거리", "260617_260622_봉오대로사거리.xlsx", 99_924, 8),
    SourceFile("봉오대로사거리", "260623_260627_봉오대로사거리.xlsx", 91_220, 8),
    SourceFile("봉오대로사거리", "260628_260630_봉오대로사거리.xlsx", 50_293, 8),
    SourceFile("산업길사거리", "260501_260531_산업길사거리.xlsx", 60_930, 8),
    SourceFile("산업길사거리", "260601_260630_산업길사거리.xlsx", 67_116, 8),
    SourceFile("삼정고가삼거리", "260501_260519_삼정고가삼거리.xlsx", 98_320, 8),
    SourceFile("삼정고가삼거리", "260520_260531_삼정고가삼거리.xlsx", 73_363, 8),
    SourceFile("삼정고가삼거리", "260601_260614_삼정고가삼거리.xlsx", 97_312, 8),
    SourceFile("삼정고가삼거리", "260615_260626_삼정고가삼거리.xlsx", 94_269, 8),
    SourceFile("삼정고가삼거리", "260627_260630_삼정고가삼거리.xlsx", 27_421, 8),
    SourceFile("삼정교사거리", "260501_260630_삼정교사거리.xlsx", 58_243, 8),
    SourceFile("삼정동320-1부천IC", "260501_260510_삼정동320-1.xlsx", 74_199, 6),
    SourceFile("삼정동320-1부천IC", "260511_260520_삼정동320-1.xlsx", 91_872, 6),
    SourceFile("삼정동320-1부천IC", "260521_260531_삼정동320-1.xlsx", 91_811, 6),
    SourceFile("삼정동320-1부천IC", "260601_260610_삼정동320-1.xlsx", 92_280, 6),
    SourceFile("삼정동320-1부천IC", "260611_260620_삼정동320-1.xlsx", 91_286, 6),
    SourceFile("삼정동320-1부천IC", "260621_260630_삼정동320-1.xlsx", 89_964, 6),
    SourceFile("자동차검사소", "260501_260520_자동차검사소.xlsx", 80_166, 6),
    SourceFile("자동차검사소", "260521_260531_자동차검사소.xlsx", 31_626, 6),
    SourceFile("자동차검사소", "260601_260617_자동차검사소.xlsx", 91_626, 6),
    SourceFile("자동차검사소", "260618_260630_자동차검사소.xlsx", 58_126, 6),
)
EXPECTED_DIRECTORY_ROWS = {
    "봉오대로사거리": 1_064_333,
    "산업길사거리": 128_046,
    "삼정고가삼거리": 390_685,
    "삼정교사거리": 58_243,
    "삼정동320-1부천IC": 531_412,
    "자동차검사소": 261_544,
}
EXPECTED_LOCATION_NAMES = {
    "봉오대로사거리": {"봉오대로 사거리[동향]", "봉오대로 사거리[서향]"},
    "산업길사거리": {"산업길 사거리[남향]", "산업길 사거리[북향]"},
    "삼정고가삼거리": {"삼정고가 삼거리[동향]", "삼정고가 삼거리[서향]"},
    "삼정교사거리": {"삼정교 사거리[북동향]"},
    "삼정동320-1부천IC": {"삼정동320-1 부천IC"},
    "자동차검사소": {"자동차검사소"},
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db", type=Path, default=DB_PATH, help="Target SQLite DB path")
    parser.add_argument(
        "--source-root", type=Path, default=SOURCE_ROOT, help="Directory containing source folders"
    )
    parser.add_argument(
        "--verify-only",
        action="store_true",
        help="Validate the imported data without writing to SQLite",
    )
    return parser.parse_args()


def cell_to_text(value: Any) -> str:
    return "" if value is None else str(value)


def source_path(source_root: Path, source_file: SourceFile) -> Path:
    return source_root / source_file.directory / source_file.file_name


def validate_source_catalog(source_root: Path) -> None:
    if sum(source.expected_rows for source in SOURCE_FILES) != TOTAL_TARGET_ROWS:
        raise RuntimeError("Configured source row counts do not match the target total")
    if sum(EXPECTED_DIRECTORY_ROWS.values()) != TOTAL_TARGET_ROWS:
        raise RuntimeError("Configured directory row counts do not match the target total")

    expected_paths = {source_path(source_root, source) for source in SOURCE_FILES}
    missing_paths = sorted(path for path in expected_paths if not path.is_file())
    if missing_paths:
        raise FileNotFoundError("Missing source files: " + ", ".join(map(str, missing_paths)))


def validate_db_schema(conn: sqlite3.Connection) -> None:
    vehicle_columns = {row[1] for row in conn.execute(f"PRAGMA table_info({TABLE_NAME})")}
    missing_vehicle_columns = [column for column in INSERT_COLUMNS if column not in vehicle_columns]
    if missing_vehicle_columns:
        raise RuntimeError(
            "Missing vehicle_detection columns: " + ", ".join(missing_vehicle_columns)
        )

    summary_columns = {row[1] for row in conn.execute(f"PRAGMA table_info({SUMMARY_TABLE_NAME})")}
    required_summary_columns = {"source_file", "source_sheet", "column_count", "imported_rows"}
    missing_summary_columns = sorted(required_summary_columns - summary_columns)
    if missing_summary_columns:
        raise RuntimeError("Missing import_summary columns: " + ", ".join(missing_summary_columns))


def iter_source_rows(path: Path, source_file: SourceFile) -> Iterable[tuple[str, ...]]:
    expected_header = EIGHT_COLUMN_HEADER if source_file.column_count == 8 else SIX_COLUMN_HEADER
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(f"Missing sheet {SHEET_NAME!r}: {path.name}")
        worksheet = workbook[SHEET_NAME]
        worksheet.reset_dimensions()
        rows = worksheet.iter_rows(values_only=True)
        header = tuple(cell_to_text(value) for value in next(rows, ()))
        if header != expected_header:
            raise RuntimeError(f"Header mismatch in {path.name}: {header!r}")

        for row in rows:
            if len(row) != source_file.column_count:
                raise RuntimeError(f"Column count mismatch in {path.name}: {len(row)}")
            if source_file.column_count == 8:
                (
                    collected_at,
                    registered_at,
                    device_id,
                    lane,
                    direction,
                    location,
                    vehicle,
                    owner,
                ) = row
            else:
                collected_at, registered_at, device_id, location, vehicle, owner = row
                lane = direction = ""
            location_text = cell_to_text(location)
            if location_text not in EXPECTED_LOCATION_NAMES[source_file.directory]:
                raise RuntimeError(f"Unexpected location in {path.name}: {location_text!r}")
            yield (
                cell_to_text(collected_at),
                cell_to_text(registered_at),
                cell_to_text(device_id),
                location_text,
                cell_to_text(lane),
                cell_to_text(direction),
                cell_to_text(vehicle),
                cell_to_text(owner),
                source_file.file_name,
                SHEET_NAME,
            )
    finally:
        workbook.close()


def target_file_names() -> tuple[str, ...]:
    return tuple(source.file_name for source in SOURCE_FILES)


def delete_target_rows(conn: sqlite3.Connection) -> tuple[int, int]:
    placeholders = ", ".join("?" for _ in SOURCE_FILES)
    parameters = target_file_names()
    deleted_vehicle = conn.execute(
        f"DELETE FROM {TABLE_NAME} WHERE source_file IN ({placeholders})", parameters
    ).rowcount
    deleted_summary = conn.execute(
        f"DELETE FROM {SUMMARY_TABLE_NAME} WHERE source_file IN ({placeholders})", parameters
    ).rowcount
    return deleted_vehicle, deleted_summary


def import_source_file(conn: sqlite3.Connection, path: Path, source_file: SourceFile) -> int:
    placeholders = ", ".join("?" for _ in INSERT_COLUMNS)
    sql = f"INSERT INTO {TABLE_NAME} ({', '.join(INSERT_COLUMNS)}) VALUES ({placeholders})"
    batch: list[tuple[str, ...]] = []
    row_count = 0
    for row in iter_source_rows(path, source_file):
        batch.append(row)
        if len(batch) == 5_000:
            conn.executemany(sql, batch)
            row_count += len(batch)
            batch.clear()
    if batch:
        conn.executemany(sql, batch)
        row_count += len(batch)
    if row_count != source_file.expected_rows:
        raise RuntimeError(
            f"Row count mismatch in {source_file.file_name}: {row_count:,} != {source_file.expected_rows:,}"
        )
    conn.execute(
        f"INSERT INTO {SUMMARY_TABLE_NAME} (source_file, source_sheet, column_count, imported_rows) "
        "VALUES (?, ?, ?, ?)",
        (source_file.file_name, SHEET_NAME, source_file.column_count, row_count),
    )
    return row_count


def target_file_counts(conn: sqlite3.Connection) -> dict[str, int]:
    placeholders = ", ".join("?" for _ in SOURCE_FILES)
    return dict(
        conn.execute(
            f"SELECT source_file, COUNT(*) FROM {TABLE_NAME} "
            f"WHERE source_file IN ({placeholders}) GROUP BY source_file",
            target_file_names(),
        ).fetchall()
    )


def validate_import(conn: sqlite3.Connection) -> None:
    expected_file_counts = {source.file_name: source.expected_rows for source in SOURCE_FILES}
    file_counts = target_file_counts(conn)
    if file_counts != expected_file_counts:
        raise RuntimeError(f"Imported file counts mismatch: {file_counts!r}")

    directory_counts: dict[str, int] = defaultdict(int)
    for source in SOURCE_FILES:
        directory_counts[source.directory] += file_counts[source.file_name]
    if dict(directory_counts) != EXPECTED_DIRECTORY_ROWS:
        raise RuntimeError(f"Imported directory counts mismatch: {dict(directory_counts)!r}")

    placeholders = ", ".join("?" for _ in SOURCE_FILES)
    location_counts = dict(
        conn.execute(
            f"SELECT location_name, COUNT(*) FROM {TABLE_NAME} "
            f"WHERE source_file IN ({placeholders}) GROUP BY location_name",
            target_file_names(),
        ).fetchall()
    )
    allowed_locations = set().union(*EXPECTED_LOCATION_NAMES.values())
    if (
        set(location_counts) != allowed_locations
        or sum(location_counts.values()) != TOTAL_TARGET_ROWS
    ):
        raise RuntimeError(f"Imported location counts mismatch: {location_counts!r}")

    summary_counts = dict(
        conn.execute(
            f"SELECT source_file, imported_rows FROM {SUMMARY_TABLE_NAME} "
            f"WHERE source_file IN ({placeholders})",
            target_file_names(),
        ).fetchall()
    )
    if summary_counts != expected_file_counts:
        raise RuntimeError(f"Import summary counts mismatch: {summary_counts!r}")

    total_rows = conn.execute(f"SELECT COUNT(*) FROM {TABLE_NAME}").fetchone()[0]
    if total_rows != EXPECTED_TOTAL_DB_ROWS:
        raise RuntimeError(f"Total DB rows mismatch: {total_rows:,} != {EXPECTED_TOTAL_DB_ROWS:,}")

    print(f"target_rows={TOTAL_TARGET_ROWS}")
    print(f"total_vehicle_detection_rows={total_rows}")
    print("directory_counts=" + repr(dict(directory_counts)))
    print("location_counts=" + repr(location_counts))


def open_read_only_connection(db_path: Path) -> sqlite3.Connection:
    return sqlite3.connect(f"{db_path.resolve().as_uri()}?mode=ro", uri=True)


def main() -> None:
    args = parse_args()
    validate_source_catalog(args.source_root)
    if args.verify_only:
        conn = open_read_only_connection(args.db)
        try:
            validate_db_schema(conn)
            validate_import(conn)
        finally:
            conn.close()
        print("verification=passed")
        return

    conn = sqlite3.connect(args.db)
    try:
        validate_db_schema(conn)
        with conn:
            deleted_vehicle, deleted_summary = delete_target_rows(conn)
            inserted_rows = sum(
                import_source_file(conn, source_path(args.source_root, source), source)
                for source in SOURCE_FILES
            )
            if inserted_rows != TOTAL_TARGET_ROWS:
                raise RuntimeError(
                    f"Inserted total mismatch: {inserted_rows:,} != {TOTAL_TARGET_ROWS:,}"
                )
            validate_import(conn)
    finally:
        conn.close()
    print(f"deleted_vehicle_rows={deleted_vehicle}")
    print(f"deleted_import_summary_rows={deleted_summary}")
    print(f"inserted_rows={inserted_rows}")


if __name__ == "__main__":
    main()
