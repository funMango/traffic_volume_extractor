#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""삼정동 일대 대형트럭 2026년 3/4월 월 일평균 교통량 추출."""

from __future__ import annotations

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

try:
    import oracledb
except ImportError:  # pragma: no cover - runtime dependency check
    oracledb = None

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - runtime dependency check
    load_dotenv = None


BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / "00_Data" / ".env"
OUTPUT_PATH = (
    BASE_DIR / "02_Result" / "교통량_추출" / "교통량" / "삼정동일대_교통량.xlsx"
)

TRAFFIC_TABLE = "S_CRSRD_VKND_TRF_1HH"
INTERSECTION_TABLE = "M_CRSRD_INF"
CODE_TABLE = "M_CD_INF"
VEHICLE_GROUP_CODE = "VHCL_ATTR_CD"
VEHICLE_CODE = "20000"
VEHICLE_NAME = "대형트럭"
MONTHS = [(2026, 3), (2026, 4)]
AVERAGE_CRITERIA = "조회된 일별 교통량이 있는 날짜만 포함하여 월 일평균 산정"

TARGET_INTERSECTIONS = [
    ("신흥로", "BC001N0024", "부천IC입구"),
    ("신흥로", "BC001N0114", "산업길사거리"),
    ("신흥로", "BC001N0126", "오정산업단지사거리"),
    ("오정로", "BC001N0003", "삼정고가교삼거리"),
    ("오정로", "BC001N0004", "삼정교사거리"),
    ("오정로", "BC001N0132", "내촌고가삼거리"),
    ("석천로", "BC002N0282", "삼정교삼거리"),
    ("석천로", "BC002N0283", "오정물류단지상류1"),
    ("석천로", "BC002N0284", "오정물류단지상류2"),
    ("석천로", "BC001N0002", "대장동공영차고지사거리"),
]

BLOCKED_SQL_PATTERN = re.compile(
    r"\b(INSERT|UPDATE|DELETE|MERGE|DROP|TRUNCATE|ALTER|CREATE|COMMIT|ROLLBACK|GRANT|REVOKE)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Intersection:
    road: str
    node_id: str
    name: str
    db_name: str


def _progress_bar(pct: float, width: int = 20) -> str:
    bounded = min(max(pct, 0.0), 100.0)
    filled = int(round(width * bounded / 100))
    return "[" + "#" * filled + "-" * (width - filled) + "]"


def _stage(stage_no: int, title: str, pct: float, detail: str) -> None:
    print(f"단계({stage_no}/5) {title} {_progress_bar(pct)} {pct:.2f}% | {detail}", flush=True)


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
    required_traffic_columns = {"TOT_DT", "NODE_ID", "VKND_CD", "TRF_QNTY"}
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


def _validate_vehicle_code(cursor: Any) -> None:
    rows = _execute_select(
        cursor,
        f"""
        SELECT TRIM(TO_CHAR(CD)) AS CD,
               CD_NM
        FROM {CODE_TABLE}
        WHERE GRP_CD = :group_code
          AND TRIM(TO_CHAR(CD)) = :vehicle_code
        """,
        {"group_code": VEHICLE_GROUP_CODE, "vehicle_code": VEHICLE_CODE},
    )
    if not rows:
        raise RuntimeError(f"차량 코드가 없습니다: {VEHICLE_GROUP_CODE}/{VEHICLE_CODE}")
    db_vehicle_name = str(rows[0][1])
    if db_vehicle_name != VEHICLE_NAME:
        raise RuntimeError(
            f"차량 코드명이 다릅니다: 코드={VEHICLE_CODE}, 요청명={VEHICLE_NAME}, DB명={db_vehicle_name}"
        )


def _resolve_intersections(cursor: Any) -> list[Intersection]:
    bind_names = [f"id{idx}" for idx in range(len(TARGET_INTERSECTIONS))]
    placeholders = ", ".join(f":{name}" for name in bind_names)
    params = {
        bind_name: node_id
        for bind_name, (_road, node_id, _name) in zip(bind_names, TARGET_INTERSECTIONS)
    }
    rows = _execute_select(
        cursor,
        f"""
        SELECT NODE_ID,
               CRSRD_NM
        FROM {INTERSECTION_TABLE}
        WHERE NODE_ID IN ({placeholders})
        """,
        params,
    )
    by_id = {str(node_id): str(name) for node_id, name in rows}

    missing = []
    mismatches = []
    intersections: list[Intersection] = []
    for idx, (road, node_id, requested_name) in enumerate(TARGET_INTERSECTIONS, start=1):
        db_name = by_id.get(node_id)
        if db_name is None:
            missing.append(f"{node_id} {requested_name}")
            continue
        if db_name != requested_name:
            mismatches.append(f"{node_id}: 요청명={requested_name}, DB명={db_name}")
        intersections.append(Intersection(road=road, node_id=node_id, name=requested_name, db_name=db_name))
        _stage(2, "테이블/교차로/차종 검증", idx / len(TARGET_INTERSECTIONS) * 100, requested_name)

    if missing:
        raise RuntimeError("교차로를 찾을 수 없습니다: " + ", ".join(missing))
    if mismatches:
        raise RuntimeError("교차로명이 DB와 다릅니다: " + "; ".join(mismatches))
    if len(intersections) != 10:
        raise RuntimeError(f"대상 교차로 수가 10개가 아닙니다: {len(intersections)}개")
    return intersections


def _next_month(year: int, month: int) -> tuple[int, int]:
    if month == 12:
        return year + 1, 1
    return year, month + 1


def _month_bounds(year: int, month: int) -> tuple[date, date]:
    next_year, next_month = _next_month(year, month)
    return date(year, month, 1), date(next_year, next_month, 1)


def _fetch_monthly_daily_rows(cursor: Any, node_id: str, year: int, month: int) -> list[tuple[Any, ...]]:
    start, end_next = _month_bounds(year, month)
    return _execute_select(
        cursor,
        f"""
        SELECT TRUNC(TOT_DT) AS TRAFFIC_DATE,
               SUM(NVL(TRF_QNTY, 0)) AS DAILY_TRAFFIC
        FROM {TRAFFIC_TABLE}
        WHERE NODE_ID = :node_id
          AND TRIM(TO_CHAR(VKND_CD)) = :vehicle_code
          AND TOT_DT >= :date_start
          AND TOT_DT < :date_end_next
          AND TRF_QNTY IS NOT NULL
        GROUP BY TRUNC(TOT_DT)
        ORDER BY TRAFFIC_DATE
        """,
        {
            "node_id": node_id,
            "vehicle_code": VEHICLE_CODE,
            "date_start": start,
            "date_end_next": end_next,
        },
    )


def _round_half_up(value: float | None) -> int | None:
    if value is None:
        return None
    return int(value + 0.5)


def _calc_month_result(rows: list[tuple[Any, ...]]) -> dict[str, Any]:
    daily_totals = [float(daily_traffic) for _traffic_date, daily_traffic in rows]
    collected_days = len(daily_totals)
    traffic_sum = sum(daily_totals)
    average_raw = traffic_sum / collected_days if collected_days else None
    return {
        "collected_days": collected_days,
        "traffic_sum": round(traffic_sum, 6),
        "average_raw": round(average_raw, 6) if average_raw is not None else None,
        "average_rounded": _round_half_up(average_raw),
    }


def _month_label(year: int, month: int) -> str:
    return f"{year}년 {month}월"


def _summary_header(year: int, month: int) -> str:
    return f"{year}년 {month}월 일평균"


def _extract_monthly_average(
    cursor: Any, intersections: list[Intersection]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    summary_rows = [
        {
            "도로": item.road,
            "교차로ID": item.node_id,
            "교차로이름": item.name,
        }
        for item in intersections
    ]
    summary_by_id = {row["교차로ID"]: row for row in summary_rows}
    info_rows: list[dict[str, Any]] = []

    total_tasks = len(intersections) * len(MONTHS)
    completed = 0
    for inter_idx, inter in enumerate(intersections, start=1):
        for month_idx, (year, month) in enumerate(MONTHS, start=1):
            detail = (
                f"교차로({inter_idx}/{len(intersections)}) {inter.name} | "
                f"월({month_idx}/{len(MONTHS)}) {_month_label(year, month)}"
            )
            _stage(3, "월별 일교통량 조회", completed / total_tasks * 100, detail)
            result = _calc_month_result(_fetch_monthly_daily_rows(cursor, inter.node_id, year, month))
            completed += 1

            summary_by_id[inter.node_id][_summary_header(year, month)] = result["average_rounded"]
            info_rows.append(
                {
                    "도로": inter.road,
                    "교차로ID": inter.node_id,
                    "교차로이름": inter.name,
                    "월": _month_label(year, month),
                    "수집일 수": result["collected_days"],
                    "월 합계": result["traffic_sum"],
                    "평균 산정값": result["average_raw"],
                    "반올림값": result["average_rounded"],
                    "차량코드": VEHICLE_CODE,
                    "차량명": VEHICLE_NAME,
                    "DB table": TRAFFIC_TABLE,
                    "평균 기준": AVERAGE_CRITERIA,
                }
            )
            _stage(3, "월별 일교통량 조회", completed / total_tasks * 100, detail + " 완료")

    return summary_rows, info_rows


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
        max_width = max(
            (_display_width(row[col_idx]) for row in rows if col_idx < len(row)), default=0
        )
        widths.append(
            f'<col min="{col_idx + 1}" max="{col_idx + 1}" '
            f'width="{min(max_width + 2, 50)}" customWidth="1"/>'
        )

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
                cells.append(f'<c r="{ref}" t="inlineStr"{style}><is><t>{escape(str(value))}</t></is></c>')
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
        workbook_sheets.append(f'<sheet name="{escape(sheet_name)}" sheetId="{idx}" r:id="rId{idx}"/>')
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
    _stage(4, "XLSX 저장", 0.0, str(path))
    path.parent.mkdir(parents=True, exist_ok=True)

    traffic_headers = [
        "도로",
        "교차로ID",
        "교차로이름",
        _summary_header(2026, 3),
        _summary_header(2026, 4),
    ]
    traffic_sheet_rows = [traffic_headers]
    for row in summary_rows:
        traffic_sheet_rows.append([row.get(header) for header in traffic_headers])

    info_headers = [
        "도로",
        "교차로ID",
        "교차로이름",
        "월",
        "수집일 수",
        "월 합계",
        "평균 산정값",
        "반올림값",
        "차량코드",
        "차량명",
        "DB table",
        "평균 기준",
    ]
    info_sheet_rows = [info_headers]
    for row in info_rows:
        info_sheet_rows.append([row.get(header) for header in info_headers])

    _write_xlsx(path, [("교통량", traffic_sheet_rows), ("산정정보", info_sheet_rows)])
    _stage(4, "XLSX 저장", 100.0, "저장 완료")


def _read_xlsx_rows(path: Path) -> dict[str, list[list[Any]]]:
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

    def _cell_col_idx(cell_ref: str) -> int:
        letters = re.sub(r"[^A-Z]", "", cell_ref.upper())
        col_idx = 0
        for letter in letters:
            col_idx = col_idx * 26 + ord(letter) - 64
        return col_idx

    with zipfile.ZipFile(path) as zf:
        workbook_xml = zf.read("xl/workbook.xml")
        workbook_rels_xml = zf.read("xl/_rels/workbook.xml.rels")
        workbook_root = ET.fromstring(workbook_xml)
        rels_root = ET.fromstring(workbook_rels_xml)

        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root
            if rel.attrib["Type"].endswith("/worksheet")
        }
        sheets: dict[str, list[list[Any]]] = {}
        sheets_elem = workbook_root.find("a:sheets", ns)
        if sheets_elem is None:
            return sheets
        for sheet in sheets_elem:
            sheet_name = sheet.attrib["name"]
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rel_targets[rel_id]
            target_path = "xl/" + target.lstrip("/")
            sheet_root = ET.fromstring(zf.read(target_path))
            rows = []
            for row in sheet_root.findall(".//a:row", ns):
                values_by_col: dict[int, Any] = {}
                for cell in row.findall("a:c", ns):
                    col_idx = _cell_col_idx(cell.attrib["r"])
                    value_elem = cell.find("a:v", ns)
                    inline_text = cell.find("a:is/a:t", ns)
                    if inline_text is not None:
                        values_by_col[col_idx] = inline_text.text or ""
                    elif value_elem is not None:
                        raw = value_elem.text or ""
                        values_by_col[col_idx] = float(raw) if "." in raw else int(raw)
                    else:
                        values_by_col[col_idx] = None
                max_col = max(values_by_col, default=0)
                values = [values_by_col.get(col_idx) for col_idx in range(1, max_col + 1)]
                rows.append(values)
            sheets[sheet_name] = rows
        return sheets


def _verify_workbook(path: Path) -> None:
    _stage(5, "XLSX 검증", 0.0, str(path))
    if not path.exists():
        raise RuntimeError(f"XLSX 파일이 생성되지 않았습니다: {path}")
    sheets = _read_xlsx_rows(path)
    expected_sheets = {"교통량", "산정정보"}
    if set(sheets) != expected_sheets:
        raise RuntimeError(f"시트 구성이 다릅니다: {sorted(sheets)}")

    traffic_rows = sheets["교통량"]
    info_rows = sheets["산정정보"]
    if len(traffic_rows) != len(TARGET_INTERSECTIONS) + 1:
        raise RuntimeError(f"교통량 시트 행 수가 다릅니다: {len(traffic_rows)}")
    if len(info_rows) != len(TARGET_INTERSECTIONS) * len(MONTHS) + 1:
        raise RuntimeError(f"산정정보 시트 행 수가 다릅니다: {len(info_rows)}")

    traffic_header = traffic_rows[0]
    info_header = info_rows[0]
    normalized_traffic_rows = [row + [None] * (len(traffic_header) - len(row)) for row in traffic_rows[1:]]
    traffic_by_key = {
        (row[1], _summary_header(2026, 3)): row[3]
        for row in normalized_traffic_rows
    } | {
        (row[1], _summary_header(2026, 4)): row[4]
        for row in normalized_traffic_rows
    }
    month_to_header = {
        _month_label(2026, 3): _summary_header(2026, 3),
        _month_label(2026, 4): _summary_header(2026, 4),
    }
    node_idx = info_header.index("교차로ID")
    month_idx = info_header.index("월")
    rounded_idx = info_header.index("반올림값")
    for row in info_rows[1:]:
        key = (row[node_idx], month_to_header[row[month_idx]])
        if traffic_by_key.get(key) != row[rounded_idx]:
            raise RuntimeError(f"교통량/산정정보 반올림값이 다릅니다: {key}")

    if traffic_header != [
        "도로",
        "교차로ID",
        "교차로이름",
        _summary_header(2026, 3),
        _summary_header(2026, 4),
    ]:
        raise RuntimeError("교통량 시트 헤더가 다릅니다.")
    _stage(5, "XLSX 검증", 100.0, "검증 완료")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="삼정동 일대 대형트럭 월 일평균 교통량 XLSX 생성")
    parser.add_argument("--output", default=str(OUTPUT_PATH))
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    output_path = Path(args.output)
    conn = None
    try:
        _stage(1, "DB 연결 및 읽기전용 보호", 0.0, "DB 연결 시작")
        conn = _connect_db(ENV_PATH)
        _stage(1, "DB 연결 및 읽기전용 보호", 100.0, "SET TRANSACTION READ ONLY 완료")

        with conn.cursor() as cur:
            _stage(2, "테이블/교차로/차종 검증", 0.0, "필수 컬럼 검증 시작")
            _validate_tables(cur)
            _validate_vehicle_code(cur)
            intersections = _resolve_intersections(cur)
            _stage(2, "테이블/교차로/차종 검증", 100.0, f"10개 교차로 및 {VEHICLE_NAME} 코드 검증 완료")

            _stage(3, "월별 일교통량 조회", 0.0, "조회 시작")
            summary_rows, info_rows = _extract_monthly_average(cur, intersections)
            _stage(3, "월별 일교통량 조회", 100.0, "월 일평균 산정 완료")

        _save_excel(summary_rows, info_rows, output_path)
        _verify_workbook(output_path)
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
