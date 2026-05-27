from __future__ import annotations

import csv
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
SOURCE_PATH = (
    BASE_DIR
    / "02_Result"
    / "차종별_교통량_추출"
    / "엠큐닉_2601_2604"
    / "기후에너지제공데이터(260101~260430).csv"
)
OUTPUT_DIR = SOURCE_PATH.parent / "기후에너지제공데이터_3분할"
EXPECTED_HEADER = ["교차로", "시간", "교차로 방향", "방향", "차종", "교통량"]
OUTPUT_SPECS = [
    ("기후에너지제공데이터_1_교차로001-095.csv", 95),
    ("기후에너지제공데이터_2_교차로096-190.csv", 95),
    ("기후에너지제공데이터_3_교차로191-286.csv", None),
]


def open_csv(path: Path, mode: str):
    return path.open(mode, encoding="utf-8-sig", newline="")


def read_header(reader: csv.reader) -> list[str]:
    try:
        return next(reader)
    except StopIteration as exc:
        raise RuntimeError("CSV file is empty.") from exc


def collect_intersections() -> tuple[list[str], list[str], int]:
    intersections: set[str] = set()
    row_count = 0

    with open_csv(SOURCE_PATH, "r") as source_file:
        reader = csv.reader(source_file)
        header = read_header(reader)
        if header != EXPECTED_HEADER:
            raise RuntimeError(f"Unexpected header: {header}")

        for row in reader:
            row_count += 1
            if len(row) != len(header):
                raise RuntimeError(f"Invalid column count at data row {row_count}: {row}")
            intersections.add(row[0])

    return header, sorted(intersections), row_count


def build_split_map(intersections: list[str]) -> dict[str, int]:
    split_map: dict[str, int] = {}
    start = 0

    for file_index, (_, size) in enumerate(OUTPUT_SPECS):
        end = len(intersections) if size is None else start + size
        for intersection in intersections[start:end]:
            split_map[intersection] = file_index
        start = end

    if len(split_map) != len(intersections):
        raise RuntimeError("Split map does not cover all intersections.")

    return split_map


def write_split_files(header: list[str], split_map: dict[str, int]) -> list[int]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    output_files = []
    writers = []
    row_counts = [0 for _ in OUTPUT_SPECS]

    try:
        for file_name, _ in OUTPUT_SPECS:
            output_file = open_csv(OUTPUT_DIR / file_name, "w")
            writer = csv.writer(output_file, lineterminator="\n")
            writer.writerow(header)
            output_files.append(output_file)
            writers.append(writer)

        with open_csv(SOURCE_PATH, "r") as source_file:
            reader = csv.reader(source_file)
            read_header(reader)

            for row in reader:
                file_index = split_map[row[0]]
                writers[file_index].writerow(row)
                row_counts[file_index] += 1
    finally:
        for output_file in output_files:
            output_file.close()

    return row_counts


def verify_outputs(
    header: list[str],
    source_intersections: list[str],
    source_row_count: int,
    expected_row_counts: list[int],
) -> None:
    output_intersections: list[set[str]] = []
    actual_row_counts: list[int] = []

    for file_index, (file_name, _) in enumerate(OUTPUT_SPECS):
        path = OUTPUT_DIR / file_name
        intersections: set[str] = set()
        row_count = 0

        with open_csv(path, "r") as output_file:
            reader = csv.reader(output_file)
            output_header = read_header(reader)
            if output_header != header:
                raise RuntimeError(f"Header mismatch in {file_name}: {output_header}")

            for row in reader:
                row_count += 1
                intersections.add(row[0])

        if row_count != expected_row_counts[file_index]:
            raise RuntimeError(
                f"Row count mismatch in {file_name}: "
                f"wrote {expected_row_counts[file_index]}, read {row_count}"
            )

        output_intersections.append(intersections)
        actual_row_counts.append(row_count)

    for left_index in range(len(output_intersections)):
        for right_index in range(left_index + 1, len(output_intersections)):
            overlap = output_intersections[left_index] & output_intersections[right_index]
            if overlap:
                raise RuntimeError(
                    f"Intersection overlap between files {left_index + 1} and "
                    f"{right_index + 1}: {sorted(overlap)[:5]}"
                )

    merged_intersections = set().union(*output_intersections)
    if merged_intersections != set(source_intersections):
        missing = set(source_intersections) - merged_intersections
        extra = merged_intersections - set(source_intersections)
        raise RuntimeError(
            f"Intersection mismatch. missing={sorted(missing)[:5]}, extra={sorted(extra)[:5]}"
        )

    if sum(actual_row_counts) != source_row_count:
        raise RuntimeError(
            f"Total row count mismatch: source={source_row_count}, "
            f"outputs={sum(actual_row_counts)}"
        )


def main() -> None:
    header, intersections, source_row_count = collect_intersections()
    split_map = build_split_map(intersections)
    row_counts = write_split_files(header, split_map)
    verify_outputs(header, intersections, source_row_count, row_counts)

    print(f"source_rows={source_row_count}")
    print(f"source_intersections={len(intersections)}")
    for index, ((file_name, _), row_count) in enumerate(zip(OUTPUT_SPECS, row_counts), start=1):
        first = (index - 1) * 95 + 1
        last = first + len([name for name, split_index in split_map.items() if split_index == index - 1]) - 1
        print(f"{file_name}: rows={row_count}, intersections={last - first + 1}")
    print(f"output_dir={OUTPUT_DIR}")


if __name__ == "__main__":
    main()
