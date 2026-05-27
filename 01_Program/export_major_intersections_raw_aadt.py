#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""2025년 연간보고서 주요교차로 2024/2023 연간 1일 평균 교통량 추출."""

from __future__ import annotations

import argparse
import calendar
import os
import re
import sys
import zipfile
from datetime import date
from pathlib import Path
from typing import Any, Iterable
from xml.sax.saxutils import escape

try:
    import oracledb
except ImportError:  # pragma: no cover - runtime dependency check
    oracledb = None

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - runtime dependency check
    load_dotenv = None


YEARS = [2024, 2023]
ENV_PATH = Path(__file__).resolve().parent.parent / "00_Data" / ".env"
TRAFFIC_TABLE = "S_CRSRD_TRF_1DD"
INTERSECTION_TABLE = "M_CRSRD_INF"
AVERAGE_CRITERIA = "조회된 일별 교통량이 있는 날짜만 포함하여 연간 1일 평균 산정"

MAJOR_INTERSECTIONS = [
    ("BC001N0161", "중동IC사거리"),
    ("BC001N0122", "문예사거리"),
    ("BC001N0121", "석천사거리"),
    ("BC001N0135", "춘의사거리"),
    ("BC001N0119", "상동역사거리"),
    ("BC001N0009", "송내대로사거리"),
    ("BC001N0117", "넘말사거리"),
    ("BC001N0115", "부천체육관사거리"),
    ("BC001N0130", "심곡고가사거리"),
    ("BC001N0042", "꿈마을사거리"),
    ("BC001N0005", "지역난방공사삼거리"),
    ("BC001N0001", "봉오고가교사거리"),
    ("BC001N0028", "계남고가사거리"),
    ("BC001N0151", "까치울사거리"),
    ("BC001N0020", "송내사거리"),
    ("BC001N0147", "동남삼거리"),
    ("BC001N0143", "멀뫼사거리"),
    ("BC001N0133", "봉오대로사거리"),
    ("BC001N0113", "박촌교삼거리"),
    ("BC001N0144", "소사삼거리"),
    ("BC001N0141", "종합운동장사거리"),
    ("BC001N0123", "사단사거리"),
    ("BC001N0128", "약대교회사거리"),
    ("BC001N0156", "역곡고가사거리"),
    ("BC001N0127", "내동사거리"),
    ("BC001N0104", "복사골아파트사거리"),
    ("BC001N0160", "약대오거리"),
    ("BC001N0067", "여월삼거리"),
    ("BC001N0132", "내촌고가삼거리"),
    ("BC001N0124", "중동사거리"),
    ("BC001N0071", "역곡남부역삼거리"),
    ("BC001N0150", "작동사거리"),
    ("BC001N0114", "산업길사거리"),
    ("BC001N0062", "원종교사거리"),
    ("BC001N0137", "부천북부역사거리"),
    ("BC001N0136", "심곡천사거리"),
    ("BC001N0134", "내촌사거리"),
    ("BC002N0169", "소방서사거리"),
    ("BC001N0058", "원종IC사거리"),
    ("BC002N0250", "양지초교사거리"),
    ("BC001N0041", "도당공구상가사거리"),
    ("BC001N0155", "옥길스타필드사거리"),
    ("BC002N0213", "퀸즈파크사거리"),
    ("BC002N0171", "소명여고사거리"),
    ("BC001N0091", "범박사거리"),
    ("BC002N0183", "신흥시장사거리"),
    ("BC001N0154", "동남사거리"),
    ("BC001N0118", "송내북부역삼거리"),
    ("BC001N0108", "역곡초교입구사거리"),
    ("BC001N0004", "삼정교사거리"),
]

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_PATH = (
    BASE_DIR
    / "02_Result"
    / "교통량_추출"
    / "교통량"
    / "2025년_연간보고서_주요교차로_2024_2023_연간1일평균교통량.xlsx"
)

BLOCKED_SQL_PATTERN = re.compile(
    r"\b(INSERT|UPDATE|DELETE|MERGE|DROP|TRUNCATE|ALTER|CREATE|COMMIT|ROLLBACK|GRANT|REVOKE)\b",
    re.IGNORECASE,
)


def _progress_bar(pct: float, width: int = 20) -> str:
    bounded = min(max(pct, 0.0), 100.0)
    filled = int(round(width * bounded / 100))
    return "[" + "#" * filled + "-" * (width - filled) + "]"


def _stage(stage_no: int, title: str, pct: float, detail: str) -> None:
    print(f"단계({stage_no}/4) {title} {_progress_bar(pct)} {pct:.2f}% | {detail}", flush=True)


def _strip_sql_comments(sql: str) -> str:
    sql = re.sub(r"/\*.*?\*/", " ", sql, flags=re.DOTALL)
    return re.sub(r"--[^\n\r]*", " ", sql)


def _ensure_select_sql(sql: str) -> None:
    normalized = _strip_sql_comments(sql).strip()
    if not normalized:
        raise RuntimeError("빈 SQL은 실행할 수 없습니다.")
    if not normalized.upper().startswith("SELECT"):
        raise RuntimeError("읽기 전용 보호: SELECT 문만 실행할 수 있습니다.")
    if ";" in normalized.rstrip(";"):
        raise RuntimeError("읽기 전용 보호: 복수 SQL 문은 실행할 수 없습니다.")
    if BLOCKED_SQL_PATTERN.search(normalized):
        raise RuntimeError("읽기 전용 보호: 쓰기/DDL/트랜잭션 SQL은 실행할 수 없습니다.")


def _execute_select(cursor: Any, sql: str, params: dict[str, Any] | None = None) -> list[tuple[Any, ...]]:
    _ensure_select_sql(sql)
    cursor.execute(sql, params or {})
    return list(cursor.fetchall())


def _connect_db(env_path: Path) -> Any:
    if oracledb is None:
        raise RuntimeError("oracledb 패키지가 설치되어 있지 않습니다.")
    if load_dotenv is None:
        raise RuntimeError("python-dotenv 패키지가 설치되어 있지 않습니다.")

    load_dotenv(env_path)
    required = ["DB_USER", "DB_PASSWORD", "DB_HOST", "DB_SERVICE"]
    missing = [name for name in required if not os.getenv(name)]
    if missing:
        raise RuntimeError("DB 접속 환경변수가 없습니다: " + ", ".join(missing))

    conn = oracledb.connect(
        user=os.getenv("DB_USER", "").strip(),
        password=os.getenv("DB_PASSWORD", "").strip(),
        host=os.getenv("DB_HOST", "").strip(),
        port=int(os.getenv("DB_PORT", "1521").strip()),
        service_name=os.getenv("DB_SERVICE", "").strip(),
    )
    with conn.cursor() as cur:
        cur.execute("SET TRANSACTION READ ONLY")
    return conn


def _get_table_columns(cursor: Any, table_name: str) -> set[str]:
    rows = _execute_select(
        cursor,
        """
        SELECT COLUMN_NAME
        FROM ALL_TAB_COLUMNS
        WHERE TABLE_NAME = :table_name
        """,
        {"table_name": table_name.upper()},
    )
    return {str(row[0]).upper() for row in rows}


def _validate_tables(cursor: Any) -> None:
    required_traffic_columns = {"NODE_ID", "TOT_DT", "TRF_QNTY"}
    traffic_columns = _get_table_columns(cursor, TRAFFIC_TABLE)
    missing_traffic_columns = sorted(required_traffic_columns - traffic_columns)
    if missing_traffic_columns:
        raise RuntimeError(
            f"{TRAFFIC_TABLE} 필수 컬럼이 없습니다: {', '.join(missing_traffic_columns)}"
        )

    required_intersection_columns = {"NODE_ID", "CRSRD_NM"}
    intersection_columns = _get_table_columns(cursor, INTERSECTION_TABLE)
    missing_intersection_columns = sorted(required_intersection_columns - intersection_columns)
    if missing_intersection_columns:
        raise RuntimeError(
            f"{INTERSECTION_TABLE} 필수 컬럼이 없습니다: {', '.join(missing_intersection_columns)}"
        )


def _resolve_intersections(cursor: Any, limit: int | None = None) -> list[dict[str, Any]]:
    selected = MAJOR_INTERSECTIONS[:limit] if limit else MAJOR_INTERSECTIONS
    total = len(selected)
    bind_names = [f"id{idx}" for idx in range(total)]
    placeholders = ", ".join(f":{name}" for name in bind_names)
    params = {name: node_id for name, (node_id, _name) in zip(bind_names, selected)}
    rows = _execute_select(
        cursor,
        f"""
        SELECT NODE_ID, CRSRD_NM
        FROM {INTERSECTION_TABLE}
        WHERE NODE_ID IN ({placeholders})
        """,
        params,
    )
    by_id = {str(node_id): str(name) for node_id, name in rows}

    resolved = []
    missing = []
    mismatches = []
    for idx, (node_id, requested_name) in enumerate(selected, start=1):
        db_name = by_id.get(node_id)
        if db_name is None:
            missing.append(f"{node_id} {requested_name}")
            continue
        if db_name != requested_name:
            mismatches.append(f"{node_id}: 요청명={requested_name}, DB명={db_name}")
        resolved.append({"node_id": node_id, "name": requested_name, "db_name": db_name})
        _stage(2, "테이블/교차로 검증", idx / total * 100, f"{idx}/{total} {requested_name}")

    if missing:
        raise RuntimeError("교차로를 찾을 수 없습니다: " + ", ".join(missing))
    if mismatches:
        print("[주의] DB 교차로명과 요청 교차로명이 다른 항목:", flush=True)
        for message in mismatches:
            print(f"  - {message}", flush=True)

    return resolved


def _year_bounds(year: int) -> tuple[date, date]:
    return date(year, 1, 1), date(year + 1, 1, 1)


def _fetch_daily_traffic(cursor: Any, node_id: str, year: int) -> list[tuple[Any, Any]]:
    start, end_next = _year_bounds(year)
    return _execute_select(
        cursor,
        f"""
        SELECT TRUNC(TOT_DT) AS TRAFFIC_DATE,
               SUM(TRF_QNTY) AS DAILY_TRAFFIC
        FROM {TRAFFIC_TABLE}
        WHERE NODE_ID = :node_id
          AND TOT_DT >= :date_start
          AND TOT_DT < :date_end_next
          AND TRF_QNTY IS NOT NULL
        GROUP BY TRUNC(TOT_DT)
        ORDER BY TRAFFIC_DATE
        """,
        {"node_id": node_id, "date_start": start, "date_end_next": end_next},
    )


def _calc_year_result(rows: Iterable[tuple[Any, Any]], year: int) -> dict[str, Any]:
    daily_totals = [float(daily_traffic) for _traffic_date, daily_traffic in rows]
    collected_days = len(daily_totals)
    total_days = 366 if calendar.isleap(year) else 365
    traffic_sum = sum(daily_totals)
    average_raw = traffic_sum / collected_days if collected_days else None
    average_rounded = round(average_raw) if average_raw is not None else None
    return {
        "aadt": average_rounded,
        "collected_days": collected_days,
        "excluded_days": total_days - collected_days,
        "traffic_sum": round(traffic_sum, 6),
        "average_raw": round(average_raw, 6) if average_raw is not None else None,
    }


def _extract_aadt(cursor: Any, intersections: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    total_tasks = len(intersections) * len(YEARS)
    completed = 0
    summary_by_id = {
        inter["node_id"]: {
            "교차로ID": inter["node_id"],
            "교차로이름": inter["name"],
        }
        for inter in intersections
    }
    info_rows: list[dict[str, Any]] = []

    for node_idx, inter in enumerate(intersections, start=1):
        for year_idx, year in enumerate(YEARS, start=1):
            before_pct = completed / total_tasks * 100 if total_tasks else 100.0
            detail = f"교차로({node_idx}/{len(intersections)}) {inter['name']} | 연도({year_idx}/{len(YEARS)}) {year}"
            _stage(3, "일교통량 조회 및 평균 산정", before_pct, detail)

            rows = _fetch_daily_traffic(cursor, inter["node_id"], year)
            result = _calc_year_result(rows, year)
            completed += 1

            summary_by_id[inter["node_id"]][str(year)] = result["aadt"]
            info_rows.append(
                {
                    "교차로ID": inter["node_id"],
                    "교차로이름": inter["name"],
                    "연도": year,
                    "수집일 수": result["collected_days"],
                    "제외일 수": result["excluded_days"],
                    "일교통량 합계": result["traffic_sum"],
                    "평균 산정값": result["average_raw"],
                    "DB table": TRAFFIC_TABLE,
                    "평균 기준": AVERAGE_CRITERIA,
                }
            )

            after_pct = completed / total_tasks * 100 if total_tasks else 100.0
            _stage(3, "일교통량 조회 및 평균 산정", after_pct, detail + " 완료")

    return [summary_by_id[inter["node_id"]] for inter in intersections], info_rows


def _excel_col_name(col_idx: int) -> str:
    name = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _display_width(value: Any) -> int:
    text = "" if value is None else str(value)
    return sum(2 if ord(ch) > 127 else 1 for ch in text)


def _worksheet_xml(rows: list[list[Any]]) -> str:
    col_count = max((len(row) for row in rows), default=0)
    widths = []
    for col_idx in range(col_count):
        max_width = max((_display_width(row[col_idx]) for row in rows if col_idx < len(row)), default=0)
        width = min(max_width + 2, 50)
        widths.append(f'<col min="{col_idx + 1}" max="{col_idx + 1}" width="{width}" customWidth="1"/>')

    xml_rows = []
    for row_idx, row in enumerate(rows, start=1):
        cells = []
        for col_idx, value in enumerate(row, start=1):
            if value is None:
                continue
            ref = f"{_excel_col_name(col_idx)}{row_idx}"
            style = ' s="1"' if row_idx == 1 else ""
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append(f'<c r="{ref}"{style}><v>{value}</v></c>')
            else:
                cells.append(
                    f'<c r="{ref}" t="inlineStr"{style}><is><t>{escape(str(value))}</t></is></c>'
                )
        xml_rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    dimension = f"A1:{_excel_col_name(max(col_count, 1))}{max(len(rows), 1)}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        f'<cols>{"".join(widths)}</cols>'
        f'<sheetData>{"".join(xml_rows)}</sheetData>'
        '</worksheet>'
    )


def _write_xlsx(path: Path, sheets: list[tuple[str, list[list[Any]]]]) -> None:
    workbook_sheets = []
    workbook_rels = []
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]

    for idx, (sheet_name, _rows) in enumerate(sheets, start=1):
        workbook_sheets.append(
            f'<sheet name="{escape(sheet_name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        )
        workbook_rels.append(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
        )
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    styles_rel_id = len(sheets) + 1
    workbook_rels.append(
        f'<Relationship Id="rId{styles_rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'{"".join(overrides)}'
        '</Types>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{"".join(workbook_sheets)}</sheets>'
        '</workbook>'
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(workbook_rels)}'
        '</Relationships>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center"/>'
        '</xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/styles.xml", styles)
        for idx, (_sheet_name, rows) in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", _worksheet_xml(rows))


def _save_excel(summary_rows: list[dict[str, Any]], info_rows: list[dict[str, Any]], path: Path) -> None:
    _stage(4, "XLSX 저장", 0.00, str(path))
    path.parent.mkdir(parents=True, exist_ok=True)

    traffic_headers = ["교차로이름", "2024", "2023"]
    traffic_sheet_rows = [traffic_headers]
    for row in summary_rows:
        traffic_sheet_rows.append([row.get(header) for header in traffic_headers])

    info_headers = [
        "교차로ID",
        "교차로이름",
        "연도",
        "수집일 수",
        "제외일 수",
        "일교통량 합계",
        "평균 산정값",
        "DB table",
        "평균 기준",
    ]
    info_sheet_rows = [info_headers]
    for row in info_rows:
        info_sheet_rows.append([row.get(header) for header in info_headers])

    _write_xlsx(path, [("교통량", traffic_sheet_rows), ("산정정보", info_sheet_rows)])
    _stage(4, "XLSX 저장", 100.00, "저장 완료")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="주요교차로 2024/2023 연간 1일 평균 교통량 XLSX 생성")
    parser.add_argument("--output", default=str(OUTPUT_PATH))
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="검증용으로 앞 N개 교차로만 처리합니다. 기본값 0은 전체 50개입니다.",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    output_path = Path(args.output)
    limit = args.limit if args.limit and args.limit > 0 else None
    conn = None

    try:
        _stage(1, "DB 연결 및 읽기전용 보호", 0.00, "DB 연결 시작")
        conn = _connect_db(ENV_PATH)
        _stage(1, "DB 연결 및 읽기전용 보호", 100.00, "SET TRANSACTION READ ONLY 완료")

        with conn.cursor() as cur:
            _stage(2, "테이블/교차로 검증", 0.00, "필수 테이블/컬럼 검증 시작")
            _validate_tables(cur)
            intersections = _resolve_intersections(cur, limit=limit)
            _stage(2, "테이블/교차로 검증", 100.00, f"{len(intersections)}개 교차로 검증 완료")

            _stage(3, "일교통량 조회 및 평균 산정", 0.00, f"총 {len(intersections) * len(YEARS)}건 조회 시작")
            summary_rows, info_rows = _extract_aadt(cur, intersections)
            _stage(3, "일교통량 조회 및 평균 산정", 100.00, "연간 1일 평균 교통량 산정 완료")

        _save_excel(summary_rows, info_rows, output_path)
    except RuntimeError as exc:
        print(f"실행을 중단합니다: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - DB driver/runtime errors
        print(f"실행을 중단합니다: {exc}", file=sys.stderr)
        return 1
    finally:
        if conn is not None:
            conn.close()

    print("전체 진행률 100.00%", flush=True)
    print(f"저장 경로: {output_path.resolve()}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
