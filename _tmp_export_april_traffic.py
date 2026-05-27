from __future__ import annotations

import datetime as dt
import json
import math
import sys
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape


API_BASE_URL = "http://localhost:8000"
DATE_START = "2026-04-01"
DATE_END = "2026-04-30"
HOURS = list(range(24))
BATCH_SIZE = 16
OUTPUT_PATH = Path("02_Result") / "교통량_추출" / "교통량" / "26년_04월.xlsx"


class ExportError(RuntimeError):
    pass


def http_json(method: str, path: str, payload: dict[str, Any] | None = None, timeout: int = 300) -> Any:
    url = f"{API_BASE_URL}{path}"
    data = None
    headers = {"Accept": "application/json"}
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            body = response.read().decode("utf-8-sig")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise ExportError(f"API 요청 실패 {exc.code}: {url}\n{detail}") from exc
    except urllib.error.URLError as exc:
        raise ExportError(f"API 서버에 연결할 수 없습니다: {url}\n{exc}") from exc
    return json.loads(body) if body else {}


def load_intersections() -> list[dict[str, Any]]:
    payload = http_json("GET", "/intersections", timeout=60)
    if isinstance(payload, list):
        rows = payload
    elif isinstance(payload, dict):
        rows = payload.get("intersections") or payload.get("items") or payload.get("data") or []
    else:
        rows = []
    rows = [row for row in rows if isinstance(row, dict) and row.get("node_id")]
    rows.sort(key=lambda row: str(row.get("name") or row.get("node_name") or row.get("node_id")))
    if not rows:
        raise ExportError("/intersections 결과가 비어 있습니다.")
    return rows


def assert_openapi_paths() -> None:
    payload = http_json("GET", "/openapi.json", timeout=60)
    paths = payload.get("paths") if isinstance(payload, dict) else None
    if not isinstance(paths, dict):
        raise ExportError("/openapi.json 응답에서 paths를 찾을 수 없습니다.")
    missing = [path for path in ("/intersections", "/raw-traffic") if path not in paths]
    if missing:
        raise ExportError(f"서버 버전이 오래됨: OpenAPI에서 {', '.join(missing)}를 찾을 수 없습니다.")


def batches(items: list[dict[str, Any]], size: int) -> list[list[dict[str, Any]]]:
    return [items[index : index + size] for index in range(0, len(items), size)]


def cell_xml(value: Any) -> str:
    if value is None:
        return "<c/>"
    if isinstance(value, bool):
        return f"<c t=\"b\"><v>{1 if value else 0}</v></c>"
    if isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
        return f"<c><v>{value}</v></c>"
    return f"<c t=\"inlineStr\"><is><t>{escape(str(value))}</t></is></c>"


def row_xml(values: list[Any], row_number: int) -> str:
    return f"<row r=\"{row_number}\">{''.join(cell_xml(value) for value in values)}</row>"


def workbook_parts(sheet_rows_xml: str) -> dict[str, str]:
    return {
        "_rels/.rels": (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
            "</Relationships>"
        ),
        "xl/workbook.xml": (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            "<sheets><sheet name=\"교통량\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"
        ),
        "xl/_rels/workbook.xml.rels": (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
            "</Relationships>"
        ),
        "[Content_Types].xml": (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
            "</Types>"
        ),
        "xl/worksheets/sheet1.xml": (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            "<sheetData>"
            f"{sheet_rows_xml}"
            "</sheetData></worksheet>"
        ),
    }


def write_xlsx(path: Path, row_xml_parts: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in workbook_parts("".join(row_xml_parts)).items():
            archive.writestr(name, content)


def date_range(start: str, end: str) -> list[str]:
    current = dt.date.fromisoformat(start)
    final = dt.date.fromisoformat(end)
    days: list[str] = []
    while current <= final:
        days.append(current.isoformat())
        current += dt.timedelta(days=1)
    return days


def export() -> dict[str, Any]:
    assert_openapi_paths()
    intersections = load_intersections()
    id_to_name = {str(row["node_id"]): str(row.get("name") or row.get("node_name") or row["node_id"]) for row in intersections}
    row_parts = [row_xml(["시간", "교차로이름", "교통량"], 1)]
    excel_row = 2
    node_slot_count = 0
    min_time: str | None = None
    max_time: str | None = None
    batch_count = 0

    intersection_batches = batches(intersections, BATCH_SIZE)
    for date_value in date_range(DATE_START, DATE_END):
        for batch_index, batch in enumerate(intersection_batches, start=1):
            batch_count += 1
            node_ids = [str(row["node_id"]) for row in batch]
            request_body = {
                "node_ids": node_ids,
                "date_start": date_value,
                "date_end": date_value,
                "hours": HOURS,
            }
            for attempt in range(1, 4):
                try:
                    payload = http_json("POST", "/raw-traffic", request_body, timeout=180)
                    break
                except ExportError:
                    if attempt == 3:
                        raise
                    time.sleep(attempt * 2)
            else:
                raise ExportError("raw-traffic 호출에 실패했습니다.")

            slots = payload.get("node_slots") if isinstance(payload, dict) else None
            if not isinstance(slots, list):
                raise ExportError(f"{date_value} {batch_index}번째 배치 응답에서 node_slots를 찾을 수 없습니다.")
            for slot in slots:
                if not isinstance(slot, dict):
                    continue
                slot_date = str(slot.get("date", ""))
                hour_value = int(slot.get("hour", 0))
                time_value = f"{slot_date} {hour_value:02d}:00"
                if time_value < f"{DATE_START} 00:00" or time_value > f"{DATE_END} 23:00":
                    raise ExportError(f"응답 시간이 요청 범위를 벗어났습니다: {time_value}")
                node_id = str(slot.get("node_id", ""))
                node_name = str(slot.get("node_name") or id_to_name.get(node_id, node_id))
                volume = slot.get("traffic_volume")
                row_parts.append(row_xml([time_value, node_name, volume], excel_row))
                excel_row += 1
                node_slot_count += 1
                min_time = time_value if min_time is None or time_value < min_time else min_time
                max_time = time_value if max_time is None or time_value > max_time else max_time
            print(
                f"{date_value} batch {batch_index}/{len(intersection_batches)}: "
                f"nodes={len(node_ids)} node_slots={len(slots)}",
                flush=True,
            )

    write_xlsx(OUTPUT_PATH, row_parts)
    return {
        "output": str(OUTPUT_PATH),
        "intersection_count": len(intersections),
        "batch_count": batch_count,
        "data_rows": node_slot_count,
        "sheet_rows": excel_row - 1,
        "min_time": min_time,
        "max_time": max_time,
    }


def main() -> int:
    try:
        result = export()
    except ExportError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
