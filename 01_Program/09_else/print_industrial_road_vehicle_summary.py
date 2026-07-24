#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""산업길사거리 원본 엑셀의 접근로별·시간대별 차종 합계를 콘솔에 출력한다."""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
import unicodedata

from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parents[2]
DEFAULT_INPUT_PATH = PROJECT_DIR / "02_Result" / "차종별_교통량_추출" / "산업길사거리_260714.xlsx"
REQUIRED_COLUMNS = ("시간", "교차로 방향", "방향", "차종", "교통량")
TARGET_VEHICLES = ("세단", "소형트럭", "중형버스", "대형버스", "대형트럭")
EXPECTED_APPROACHES = ("동", "서", "남", "북")
APPROACH_COLUMN_WIDTH = 30
VEHICLE_COLUMN_WIDTH = 10
TRAFFIC_COLUMN_WIDTH = 8
RATIO_COLUMN_WIDTH = 8


@dataclass(frozen=True)
class TrafficRecord:
    """원본 워크북에서 읽은 한 건의 차종별·회전방향별 교통량."""

    observed_at: datetime
    approach: str
    movement: str
    vehicle_type: str
    traffic_volume: int


Summary = dict[datetime, dict[str, dict[str, int]]]


def parse_observed_at(value: object) -> datetime:
    """엑셀 시간 셀을 시간대 집계에 사용할 datetime으로 변환한다."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value.strip())
        except ValueError as exc:
            raise ValueError(f"시간 값이 올바르지 않습니다: {value!r}") from exc
    raise ValueError(f"시간 값이 올바르지 않습니다: {value!r}")


def read_traffic_records(input_path: Path) -> list[TrafficRecord]:
    """원본 엑셀을 읽어 필요한 열만 TrafficRecord로 변환한다."""
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        rows = worksheet.iter_rows(values_only=True)
        try:
            headers = next(rows)
        except StopIteration as exc:
            raise ValueError(f"빈 워크북입니다: {input_path}") from exc

        header_indexes = {
            str(header).strip(): index for index, header in enumerate(headers) if header is not None
        }
        missing_columns = [column for column in REQUIRED_COLUMNS if column not in header_indexes]
        if missing_columns:
            raise ValueError("필수 열이 없습니다: " + ", ".join(missing_columns))

        records = []
        for row in rows:
            if not any(value is not None for value in row):
                continue
            records.append(
                TrafficRecord(
                    observed_at=parse_observed_at(row[header_indexes["시간"]]),
                    approach=str(row[header_indexes["교차로 방향"]]).strip(),
                    movement=str(row[header_indexes["방향"]]).strip(),
                    vehicle_type=str(row[header_indexes["차종"]]).strip(),
                    traffic_volume=int(row[header_indexes["교통량"]] or 0),
                )
            )
    finally:
        workbook.close()

    if not records:
        raise ValueError(f"교통량 데이터가 없습니다: {input_path}")
    return records


def aggregate_target_vehicles(records: Iterable[TrafficRecord]) -> Summary:
    """회전 방향을 합산하고 대상 차종만 시간·접근로별로 집계한다."""
    summary: defaultdict[datetime, defaultdict[str, defaultdict[str, int]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(int))
    )
    for record in records:
        if record.vehicle_type not in TARGET_VEHICLES:
            continue
        summary[record.observed_at][record.approach][record.vehicle_type] += record.traffic_volume

    return {
        observed_at: {
            approach: dict(vehicle_totals) for approach, vehicle_totals in approach_totals.items()
        }
        for observed_at, approach_totals in summary.items()
    }


def approach_direction(approach: str) -> str | None:
    """접근로 표기에서 남·동·북·서 방향을 추출한다."""
    for direction in EXPECTED_APPROACHES:
        if f"-{direction}" in approach:
            return direction
    return None


def sort_approaches(approaches: Iterable[str]) -> list[str]:
    """동·서·남·북 순서로 접근로를 정렬한다."""
    direction_order = {direction: index for index, direction in enumerate(EXPECTED_APPROACHES)}
    return sorted(
        approaches,
        key=lambda approach: (
            direction_order.get(approach_direction(approach) or "", len(direction_order)),
            approach,
        ),
    )


def validate_summary(summary: Summary) -> None:
    """요청한 4개 접근로·2개 시간대와 대상 차종 합계를 검증한다."""
    periods = sorted(summary)
    if len(periods) != 2:
        raise ValueError(f"시간대가 2개가 아닙니다: {len(periods)}개")

    approaches = {approach for period in summary.values() for approach in period}
    directions = {approach_direction(approach) for approach in approaches}
    if directions != set(EXPECTED_APPROACHES) or len(approaches) != len(EXPECTED_APPROACHES):
        raise ValueError("접근로가 남·동·북·서 4개인지 확인할 수 없습니다.")


def display_width(value: str) -> int:
    """동아시아 문자 폭을 반영한 콘솔 표시 너비를 계산한다."""
    return sum(
        2 if unicodedata.east_asian_width(character) in {"F", "W"} else 1 for character in value
    )


def pad_right(value: str, width: int) -> str:
    """콘솔에서 지정한 표시 너비만큼 문자열 오른쪽을 공백으로 채운다."""
    return value + " " * max(0, width - display_width(value))


def format_period_label(observed_at: datetime) -> str:
    """시간대를 표 머리글에 표시할 날짜·시간 범위 문자열로 변환한다."""
    return f"{observed_at.month}월 {observed_at.day}일_{observed_at:%H}~{(observed_at.hour + 1) % 24:02d}시"


def format_ratio(traffic_volume: int, subtotal: int) -> str:
    """접근로·시간대 합계 대비 차종 교통량 구성비를 표시한다."""
    if subtotal == 0:
        return "0.0%"
    return f"{traffic_volume / subtotal:.1%}"


def format_summary(summary: Summary) -> str:
    """접근로·차종 행과 시간대 열로 된 콘솔 표 문자열을 구성한다."""
    periods = sorted(summary)
    approaches = sort_approaches(
        approach for approach_totals in summary.values() for approach in approach_totals
    )
    approaches = list(dict.fromkeys(approaches))
    header = " ".join(
        (
            pad_right("접근로명", APPROACH_COLUMN_WIDTH),
            pad_right("차종", VEHICLE_COLUMN_WIDTH),
            *(
                pad_right(
                    format_period_label(period),
                    TRAFFIC_COLUMN_WIDTH + RATIO_COLUMN_WIDTH + 1,
                )
                for period in periods
            ),
        )
    )
    subheader = " ".join(
        (
            " " * APPROACH_COLUMN_WIDTH,
            " " * VEHICLE_COLUMN_WIDTH,
            *(
                f"{'교통량':>{TRAFFIC_COLUMN_WIDTH}} {'구성비':>{RATIO_COLUMN_WIDTH}}"
                for _ in periods
            ),
        )
    )
    lines = [header, subheader, "-" * display_width(header)]

    for approach in approaches:
        subtotals = {
            period: sum(
                summary[period].get(approach, {}).get(vehicle, 0) for vehicle in TARGET_VEHICLES
            )
            for period in periods
        }
        for row_index, vehicle in enumerate(TARGET_VEHICLES):
            period_cells = []
            for period in periods:
                traffic_volume = summary[period].get(approach, {}).get(vehicle, 0)
                period_cells.append(
                    f"{traffic_volume:>{TRAFFIC_COLUMN_WIDTH},} "
                    f"{format_ratio(traffic_volume, subtotals[period]):>{RATIO_COLUMN_WIDTH}}"
                )
            lines.append(
                " ".join(
                    (
                        pad_right(approach if row_index == 0 else "", APPROACH_COLUMN_WIDTH),
                        pad_right(vehicle, VEHICLE_COLUMN_WIDTH),
                        *period_cells,
                    )
                )
            )

        total_cells = [
            f"{subtotals[period]:>{TRAFFIC_COLUMN_WIDTH},} {'100.0%':>{RATIO_COLUMN_WIDTH}}"
            for period in periods
        ]
        lines.append(
            " ".join(
                (
                    " " * APPROACH_COLUMN_WIDTH,
                    pad_right("합계", VEHICLE_COLUMN_WIDTH),
                    *total_cells,
                )
            )
        )
    return "\n".join(lines)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT_PATH,
        help="입력 워크북 경로 (기본값: 산업길사거리_260714.xlsx)",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """원본 파일을 수정하지 않고 집계 표만 표준 출력으로 보낸다."""
    args = parse_args(argv)
    try:
        records = read_traffic_records(args.input)
        summary = aggregate_target_vehicles(records)
        validate_summary(summary)
    except (OSError, ValueError) as exc:
        print(f"집계를 중단합니다: {exc}")
        return 1

    print(format_summary(summary))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
