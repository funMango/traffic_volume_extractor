"""삼정교사거리 2026-10-01 07~09시 원천 수집 확인 (읽기 전용)."""

import argparse
import csv
import importlib
import json
import sys
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "01_Program" / "00_main"))
vehicle = importlib.import_module("차종별_교차로_추출")
START = datetime(2026, 10, 1, 7)
END = datetime(2026, 10, 1, 9)
OUTPUT = ROOT / "02_Result/차종별_교통량_추출/삼정교사거리_261001_07-09_원천.csv"
HEADERS = [
    "시간",
    "교차로",
    "NODE_ID",
    "접근로 ID",
    "접근로 이름",
    "방향 코드",
    "방향 이름",
    "차종 코드",
    "차종 이름",
    "교통량",
]


def code(value):
    return "" if value is None else str(value).strip()


def mapping(rows):
    result = {}
    for key, name in rows:
        key, name = code(key), code(name)
        if key in result and result[key] != name:
            raise RuntimeError(f"기준정보 중복 매핑: {key}")
        result[key] = name
    return result


def guard_check():
    vehicle.ensure_select_sql("SELECT NODE_ID FROM M_CRSRD_INF WHERE NODE_ID = :node_id")
    for sql in (
        "DELETE FROM T",
        "UPDATE T SET X=1",
        "INSERT INTO T VALUES (1)",
        "DROP TABLE T",
        "CREATE TABLE T (X NUMBER)",
        "TRUNCATE TABLE T",
        "BEGIN NULL; END;",
        "SELECT 1 FROM DUAL; SELECT 2 FROM DUAL",
        "SELECT 1 FROM DUAL; DELETE FROM T",
        "COMMIT",
    ):
        try:
            vehicle.ensure_select_sql(sql)
        except RuntimeError:
            continue
        raise AssertionError(f"가드가 SQL을 허용함: {sql}")
    print("SELECT 가드 검증 통과 (DB 미접속)")


def extract():
    connection = vehicle.connect_db()  # 연결 직후 SET TRANSACTION READ ONLY, 실패 시 중단
    try:
        with connection.cursor() as cursor:

            def select(sql, binds=None):
                return vehicle.execute_select(cursor, sql, binds)

            columns = select(
                "SELECT COLUMN_NAME, DATA_TYPE FROM USER_TAB_COLUMNS WHERE TABLE_NAME = :table_name",
                {"table_name": "S_CRSRD_VKND_TRF_1HH"},
            )
            types = dict(columns)
            required = {"TOT_DT", "NODE_ID", "ACSR_ID", "DRCT_CD", "VKND_CD", "TRF_QNTY"}
            if not required.issubset(types) or types["TOT_DT"] not in ("DATE", "TIMESTAMP(6)"):
                raise RuntimeError(f"원천 테이블 컬럼 검증 실패: {columns}")
            targets = select(
                "SELECT NODE_ID, CRSRD_NM FROM M_CRSRD_INF "
                "WHERE REPLACE(CRSRD_NM, ' ', '') = :name ORDER BY NODE_ID",
                {"name": "삼정교사거리"},
            )
            if len(targets) != 1:
                raise RuntimeError(f"교차로 매칭이 유일하지 않음: {targets}")
            node, name = targets[0]
            approaches = mapping(
                select(
                    "SELECT ACSR_ID, ACSR_NM FROM M_CRSRD_ACSR_INF WHERE NODE_ID = :node_id",
                    {"node_id": node},
                )
            )
            directions = mapping(
                select(
                    "SELECT CD, CD_NM FROM M_CD_INF WHERE GRP_CD = :grp",
                    {"grp": "DRCT_CD"},
                )
            )
            vehicles = mapping(
                select(
                    "SELECT CD, CD_NM FROM M_CD_INF WHERE GRP_CD = :grp",
                    {"grp": "VHCL_ATTR_CD"},
                )
            )
            rows = select(
                "SELECT TOT_DT, NODE_ID, ACSR_ID, DRCT_CD, VKND_CD, TRF_QNTY "
                "FROM S_CRSRD_VKND_TRF_1HH WHERE NODE_ID = :node_id "
                "AND TOT_DT >= :start_at AND TOT_DT < :end_at "
                "ORDER BY TOT_DT, ACSR_ID, DRCT_CD, VKND_CD",
                {"node_id": node, "start_at": START, "end_at": END},
            )
    finally:
        connection.close()
    return code(node), code(name), approaches, directions, vehicles, rows


def save_and_verify(node, name, approaches, directions, vehicles, rows):
    local = json.loads((ROOT / "00_Data/VKND_KIND.json").read_text(encoding="utf-8-sig"))
    output_rows = []
    summary = defaultdict(lambda: [0, Decimal(0), 0, 0])
    for when, row_node, approach, direction, kind, volume in rows:
        if code(row_node) != node or not START <= when < END:
            raise AssertionError("원천 조회 범위 불일치")
        a, d, k = code(approach), code(direction), code(kind)
        output_rows.append(
            [
                when.isoformat(sep=" "),
                name,
                node,
                a,
                approaches.get(a, "매핑에 없는 코드"),
                d,
                directions.get(d, directions.get(d.zfill(2), "매핑에 없는 코드")),
                k,
                vehicles.get(k, local.get(k, "매핑에 없는 코드")),
                volume,
            ]
        )
        item = summary[(when.hour, k)]
        item[0] += 1
        item[1] += Decimal(str(volume)) if volume is not None else Decimal(0)
        item[2] += volume == 0
        item[3] += volume is None
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(HEADERS)
        writer.writerows(output_rows)
    with OUTPUT.open(encoding="utf-8-sig", newline="") as stream:
        loaded = list(csv.reader(stream))
    assert loaded[0] == HEADERS and len(loaded) - 1 == len(rows)
    assert loaded[1:] == [["" if v is None else str(v) for v in row] for row in output_rows]
    raw_total = sum((Decimal(str(r[5])) for r in rows if r[5] is not None), Decimal(0))
    csv_total = sum((Decimal(r[9]) for r in loaded[1:] if r[9]), Decimal(0))
    assert raw_total == csv_total == sum((s[1] for s in summary.values()), Decimal(0))
    report = [
        "# 삼정교사거리 차종별 수집 확인",
        "",
        f"- NODE_ID: `{node}`; 원천: `S_CRSRD_VKND_TRF_1HH`",
        "- 한국시간 2026-10-01 07:00 이상, 09:00 미만 (Oracle TOT_DT 현지시각 기준)",
        f"- 원천/CSV: {len(rows)}행; 교통량 합계: {raw_total}대",
        "- SELECT 가드, 범위, CSV 전체 행, 원천·CSV·요약 합계 검증 통과",
        "- NULL 교통량은 CSV 빈 값으로 보존하며 합계에서 제외",
        "",
    ]
    for hour in (7, 8):
        count = sum(s[0] for (h, _), s in summary.items() if h == hour)
        total = sum((s[1] for (h, _), s in summary.items() if h == hour), Decimal(0))
        report.append(
            f"- {hour:02}시: {count}행, {total}대 — " + ("데이터 존재" if count else "행 없음")
        )
    unknown = sorted({k for _, k in summary if k not in local})
    for hour in (7, 8):
        for kind in local:
            summary[(hour, kind)]  # 행 없는 차종도 0행으로 명시 (0대 행과 구분)
    report += [
        "",
        f"로컬 VKND_KIND.json 매핑에 없는 코드: {', '.join(unknown) or '없음'}.",
        "미등록은 매핑에 없는 코드라는 의미이며 신규 도입 차종 여부는 판단하지 않음.",
        "",
        "| 시간 | 차종 코드 | 로컬 차종 이름 | 행 수 | 교통량 합계 | 0대 행 | NULL 행 |",
        "|---|---|---|---:|---:|---:|---:|",
    ]
    for (hour, kind), (count, total, zeros, nulls) in sorted(summary.items()):
        report.append(
            f"| {hour:02}시 | {kind} | {local.get(kind, '매핑에 없는 코드')} | {count} | {total} | {zeros} | {nulls} |"
        )
    report_path = OUTPUT.with_name(OUTPUT.stem.replace("원천", "수집확인") + ".md")
    report_path.write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))
    print(f"CSV: {OUTPUT}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check-only", action="store_true", help="DB 없이 SELECT 가드 검증")
    args = parser.parse_args()
    guard_check()
    if not args.check_only:
        save_and_verify(*extract())
