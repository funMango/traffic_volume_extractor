from __future__ import annotations

import csv
import sqlite3
from collections import defaultdict
from pathlib import Path
from typing import Iterable


BASE_DIR = Path("02_Result") / "차종별_교통량_추출" / "엠큐닉_2601_2604"
SOURCE_DB = BASE_DIR / "기후에너지제공데이터(260101~260430).db"
OUTPUT_DIR = BASE_DIR / "기후에너지제공데이터_10분할"
TABLE_NAME = "traffic_by_vehicle"
SPLIT_SIZES = [28, 28, 28, 28, 28, 28, 28, 28, 28, 34]
CSV_COLUMNS = [
    ("intersection", "교차로"),
    ("observed_at", "시간"),
    ("approach_direction", "교차로 방향"),
    ("direction", "방향"),
    ("vehicle_type", "차종"),
    ("traffic_volume", "교통량"),
]


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    uri = db_path.resolve().as_uri() + "?mode=ro"
    connection = sqlite3.connect(uri, uri=True)
    connection.execute("PRAGMA query_only = ON")
    return connection


def placeholders(values: Iterable[object]) -> str:
    return ", ".join("?" for _ in values)


def output_name(part_number: int, start_number: int, end_number: int) -> str:
    return f"기후에너지제공데이터_{part_number:02d}_교차로{start_number:03d}-{end_number:03d}.csv"


def collect_intersections(connection: sqlite3.Connection) -> list[str]:
    query = f"""
        SELECT intersection
        FROM {TABLE_NAME}
        GROUP BY intersection
        ORDER BY MIN(id)
    """
    return [row[0] for row in connection.execute(query)]


def split_intersections(intersections: list[str]) -> list[list[str]]:
    if sum(SPLIT_SIZES) != len(intersections):
        raise ValueError(
            f"Split sizes total {sum(SPLIT_SIZES)}, but DB has {len(intersections)} intersections."
        )

    chunks: list[list[str]] = []
    start = 0
    for size in SPLIT_SIZES:
        end = start + size
        chunks.append(intersections[start:end])
        start = end
    return chunks


def prepare_output_dir(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for csv_path in output_dir.glob("*.csv"):
        csv_path.unlink()


def write_csv(
    connection: sqlite3.Connection,
    path: Path,
    intersections: list[str],
) -> int:
    select_columns = ", ".join(db_column for db_column, _ in CSV_COLUMNS)
    query = f"""
        SELECT {select_columns}
        FROM {TABLE_NAME}
        WHERE intersection IN ({placeholders(intersections)})
        ORDER BY id
    """

    row_count = 0
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow([csv_column for _, csv_column in CSV_COLUMNS])
        for row in connection.execute(query, intersections):
            writer.writerow(row)
            row_count += 1
    return row_count


def write_splits(
    connection: sqlite3.Connection,
    output_dir: Path,
    chunks: list[list[str]],
) -> dict[Path, int]:
    prepare_output_dir(output_dir)

    row_counts: dict[Path, int] = {}
    start_number = 1
    for part_number, chunk in enumerate(chunks, start=1):
        end_number = start_number + len(chunk) - 1
        path = output_dir / output_name(part_number, start_number, end_number)
        row_counts[path] = write_csv(connection, path, chunk)
        start_number = end_number + 1
    return row_counts


def source_row_count(connection: sqlite3.Connection) -> int:
    return connection.execute(f"SELECT COUNT(*) FROM {TABLE_NAME}").fetchone()[0]


def verify_outputs(
    output_paths: list[Path],
    expected_header: list[str],
    expected_chunks: list[list[str]],
    expected_total_rows: int,
) -> list[str]:
    messages: list[str] = []
    total_output_rows = 0
    intersection_files: dict[str, set[int]] = defaultdict(set)

    if len(output_paths) == len(SPLIT_SIZES):
        messages.append(f"OK output file count: {len(output_paths)}")
    else:
        messages.append(f"FAIL output file count: expected {len(SPLIT_SIZES)}, got {len(output_paths)}")

    for part_index, path in enumerate(output_paths, start=1):
        with path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.reader(file)
            header = next(reader, None)
            if header != expected_header:
                messages.append(f"FAIL header mismatch: {path.name}")

            row_count = 0
            seen_in_part: set[str] = set()
            part_intersections: list[str] = []
            for row in reader:
                row_count += 1
                intersection = row[0]
                intersection_files[intersection].add(part_index)
                if intersection not in seen_in_part:
                    seen_in_part.add(intersection)
                    part_intersections.append(intersection)

        expected_intersections = expected_chunks[part_index - 1]
        if part_intersections == expected_intersections:
            messages.append(
                f"OK {path.name}: rows={row_count}, intersections={len(part_intersections)}"
            )
        else:
            messages.append(
                f"FAIL intersection range: {path.name}, "
                f"expected={len(expected_intersections)}, got={len(part_intersections)}"
            )
        total_output_rows += row_count

    if total_output_rows == expected_total_rows:
        messages.append(f"OK row count: source={expected_total_rows}, outputs={total_output_rows}")
    else:
        messages.append(f"FAIL row count: source={expected_total_rows}, outputs={total_output_rows}")

    expected_intersections = {intersection for chunk in expected_chunks for intersection in chunk}
    split_intersections = {
        intersection: parts for intersection, parts in intersection_files.items() if len(parts) != 1
    }
    missing_intersections = expected_intersections - set(intersection_files)
    extra_intersections = set(intersection_files) - expected_intersections
    if split_intersections or missing_intersections or extra_intersections:
        messages.append(
            "FAIL intersection membership: "
            f"split={len(split_intersections)}, missing={len(missing_intersections)}, "
            f"extra={len(extra_intersections)}"
        )
    else:
        messages.append(f"OK intersection membership: {len(expected_intersections)} unique intersections")

    return messages


def main() -> None:
    if not SOURCE_DB.exists():
        raise FileNotFoundError(SOURCE_DB)

    with connect_readonly(SOURCE_DB) as connection:
        intersections = collect_intersections(connection)
        chunks = split_intersections(intersections)
        row_counts = write_splits(connection, OUTPUT_DIR, chunks)
        messages = verify_outputs(
            list(row_counts),
            [csv_column for _, csv_column in CSV_COLUMNS],
            chunks,
            source_row_count(connection),
        )

    print(f"source: {SOURCE_DB}")
    print(f"output: {OUTPUT_DIR}")
    print(f"unique intersections: {len(intersections)}")
    for path, row_count in row_counts.items():
        print(f"wrote: {path.name} ({row_count} rows)")
    for message in messages:
        print(message)


if __name__ == "__main__":
    main()
