"""Aggregate direction-level traffic CSV rows by time, intersection, and vehicle type."""

import csv
from collections import defaultdict
from decimal import Decimal
from pathlib import Path


INPUT_PATH = Path("02_Result/차종별_교통량_추출/대장동공영차고지사거리외10개_260904.csv")
OUTPUT_PATH = INPUT_PATH.with_name("대장동공영차고지사거리외10개_260904_교차로별_차종_통합.csv")
OUTPUT_COLUMNS = ("시간", "교차로", "차종", "교통량")


def format_quantity(quantity: Decimal) -> str:
    return str(int(quantity)) if quantity == quantity.to_integral() else format(quantity, "f")


def main() -> None:
    with INPUT_PATH.open(encoding="utf-8-sig", newline="") as source_file:
        rows = list(csv.DictReader(source_file))

    required_columns = {"시간", "교차로", "교차로 방향", "방향", "차종", "교통량"}
    if not rows or set(rows[0]) != required_columns:
        raise ValueError("원본 CSV 헤더가 예상 형식과 다릅니다.")

    totals: dict[tuple[str, str, str], Decimal] = defaultdict(Decimal)
    for row in rows:
        totals[(row["시간"], row["교차로"], row["차종"])] += Decimal(row["교통량"])

    result_rows = [
        {"시간": time, "교차로": intersection, "차종": vehicle, "교통량": format_quantity(quantity)}
        for (time, intersection, vehicle), quantity in sorted(totals.items())
    ]
    with OUTPUT_PATH.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.DictWriter(output_file, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(result_rows)

    source_total = sum((Decimal(row["교통량"]) for row in rows), Decimal())
    result_total = sum((Decimal(row["교통량"]) for row in result_rows), Decimal())
    sample_key = next(key for key in totals if key[1] == "고강지하차도사거리" and key[2] == "SUV")
    source_sample = sum(
        (
            Decimal(row["교통량"])
            for row in rows
            if (row["시간"], row["교차로"], row["차종"]) == sample_key
        ),
        Decimal(),
    )
    assert source_total == result_total
    assert source_sample == totals[sample_key]
    assert len(result_rows) == len(
        {(row["시간"], row["교차로"], row["차종"]) for row in result_rows}
    )
    assert tuple(result_rows[0]) == OUTPUT_COLUMNS
    print(f"생성: {OUTPUT_PATH}")
    print(f"원본/결과 교통량 합계: {format_quantity(source_total)}")
    print(f"고강지하차도사거리 SUV 검증: {sample_key[0]} = {format_quantity(source_sample)}")
    print(f"결과 행 수: {len(result_rows)}")


if __name__ == "__main__":
    main()
