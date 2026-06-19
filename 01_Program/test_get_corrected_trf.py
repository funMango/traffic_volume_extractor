import importlib.util
import sys
from datetime import date
from pathlib import Path

import pytest


MODULE_PATH = Path(__file__).resolve().parent / "00_main" / "get_corrected_trf.py"
SPEC = importlib.util.spec_from_file_location("get_corrected_trf", MODULE_PATH)
gct = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = gct
SPEC.loader.exec_module(gct)


def _slot(
    timestamp="2026-04-12T08:00:00",
    node_key="1",
    node_name="계남고가사거리",
    node_order=0,
    approach_id=10,
    approach_name="계남고가사거리-동(서향)",
    drct_cd="01",
    drct_name="좌",
    value=10,
):
    ts_date, ts_time = timestamp.split("T")
    hour, minute, _second = ts_time.split(":")
    return {
        "interval": "5m",
        "timestamp": timestamp,
        "date": ts_date,
        "hour": int(hour),
        "minute": int(minute),
        "node_id": int(node_key),
        "node_key": node_key,
        "node_name": node_name,
        "node_order": node_order,
        "approach_id": approach_id,
        "approach_name": approach_name,
        "drct_cd": drct_cd,
        "drct_name": drct_name,
        "traffic_volume": value,
        "corrected_value": value,
        "value": value,
    }


class RecordingProgress:
    def __init__(self, total=0):
        self.total = total
        self.starts = []
        self.phases = []
        self.advances = []
        self.finishes = []

    def start(self, phase=""):
        self.starts.append(phase)

    def update_phase(self, phase):
        self.phases.append(phase)

    def advance(self, step=1):
        self.advances.append(step)

    def finish(self, success=True):
        self.finishes.append(success)


def test_console_progress_line_includes_spinner_percent_count_and_phase():
    progress = gct.ConsoleProgress(total=2)

    progress.update_phase("API 조회: 260412~260415")
    progress.advance()
    line = progress.format_line()

    assert line.startswith("| [##########----------]")
    assert "50.00%" in line
    assert "(1/2)" in line
    assert "API 조회: 260412~260415" in line


def test_fetch_corrected_direction_slots_advances_after_each_period(monkeypatch):
    periods = gct.parse_periods("260412,260413")
    intersections = [gct.Intersection(1, "계남고가사거리", 0)]
    interval = gct.parse_interval("5분")
    time_range = gct.parse_time_range("24시간", interval)
    progress = RecordingProgress()
    calls = []

    def fake_api_post(path, payload, base_url=gct.API_BASE_URL):
        calls.append((path, payload, base_url))
        return {
            "drct_slots": [
                {
                    "interval": "5m",
                    "timestamp": f"{payload['date_start']}T08:00:00",
                    "node_id": 1,
                    "node_name": "계남고가사거리",
                    "approach_id": 10,
                    "approach_name": "계남고가사거리-동(서향)",
                    "drct_cd": "01",
                    "drct_name": "좌",
                    "corrected_value": 10,
                }
            ]
        }

    monkeypatch.setattr(gct, "api_post", fake_api_post)

    rows = gct.fetch_corrected_direction_slots(
        intersections,
        periods,
        interval,
        time_range,
        progress=progress,
    )

    assert [phase for phase in progress.phases] == ["API 조회: 260412", "API 조회: 260413"]
    assert progress.advances == [1, 1]
    assert len(calls) == 2
    assert len(rows) == 2


def test_extract_progress_total_includes_excel_save(monkeypatch):
    periods = gct.parse_periods("260412,260413")
    intersections = [gct.Intersection(1, "계남고가사거리", 0)]
    interval = gct.parse_interval("5분")
    time_range = gct.parse_time_range("24시간", interval)
    created_progress = []

    class FakeConsoleProgress(RecordingProgress):
        def __init__(self, total):
            super().__init__(total)
            created_progress.append(self)

    monkeypatch.setattr(gct, "ConsoleProgress", FakeConsoleProgress)
    monkeypatch.setattr(gct, "fetch_corrected_direction_slots", lambda *args, **kwargs: [])
    monkeypatch.setattr(gct, "make_output_path", lambda *args, **kwargs: Path("out.xlsx"))
    monkeypatch.setattr(gct, "save_workbook", lambda *args, **kwargs: None)

    output_path, row_count = gct.extract_and_save_corrected_direction_slots(
        intersections,
        periods,
        interval,
        time_range,
        gct.OUTPUT_MODE_DIRECTION,
    )

    progress = created_progress[0]
    assert output_path == Path("out.xlsx")
    assert row_count == 0
    assert progress.total == 3
    assert progress.starts == ["API 조회 준비"]
    assert progress.phases == ["Excel 저장"]
    assert progress.advances == [1]
    assert progress.finishes == [True]


def test_extract_finishes_progress_and_reraises_api_exception(monkeypatch):
    periods = gct.parse_periods("260412")
    intersections = [gct.Intersection(1, "계남고가사거리", 0)]
    interval = gct.parse_interval("5분")
    time_range = gct.parse_time_range("24시간", interval)
    created_progress = []

    class FakeConsoleProgress(RecordingProgress):
        def __init__(self, total):
            super().__init__(total)
            created_progress.append(self)

    def fail_fetch(*args, **kwargs):
        raise RuntimeError("API failure")

    monkeypatch.setattr(gct, "ConsoleProgress", FakeConsoleProgress)
    monkeypatch.setattr(gct, "fetch_corrected_direction_slots", fail_fetch)

    with pytest.raises(RuntimeError, match="API failure"):
        gct.extract_and_save_corrected_direction_slots(
            intersections,
            periods,
            interval,
            time_range,
            gct.OUTPUT_MODE_DIRECTION,
        )

    progress = created_progress[0]
    assert progress.total == 2
    assert progress.advances == []
    assert progress.finishes == [False]


def test_resolve_intersection_input_all_and_dedupes_comma_order():
    items = [
        {"node_id": 1, "name": "계남고가사거리"},
        {"node_id": 2, "name": "중동IC사거리"},
        {"node_id": 3, "name": "송내사거리"},
    ]

    all_items = gct.resolve_intersection_input("*", items)
    selected = gct.resolve_intersection_input("중동IC사거리,계남고가사거리,중동IC사거리", items)

    assert [item.node_id for item in all_items] == [1, 2, 3]
    assert [item.name for item in selected] == ["중동IC사거리", "계남고가사거리"]
    assert [item.order for item in selected] == [0, 1]


def test_parse_periods_accepts_range_day_month_and_year():
    periods = gct.parse_periods("260412~260415,260412,2605,2026")

    assert (periods[0].start, periods[0].end) == (
        date(2026, 4, 12),
        date(2026, 4, 15),
    )
    assert (periods[1].start, periods[1].end) == (
        date(2026, 4, 12),
        date(2026, 4, 12),
    )
    assert (periods[2].start, periods[2].end) == (
        date(2026, 5, 1),
        date(2026, 5, 31),
    )
    assert (periods[3].start, periods[3].end) == (
        date(2026, 1, 1),
        date(2026, 12, 31),
    )


def test_parse_time_range_validates_interval_alignment_and_24_hour_rule():
    time_range = gct.parse_time_range("17:15~17:45", "15분")

    assert (time_range.start_minute, time_range.end_minute) == (17 * 60 + 15, 17 * 60 + 45)
    assert gct.hours_for_api(time_range) == [17]

    with pytest.raises(ValueError):
        gct.parse_time_range("17:15~17:45", "1시간")
    with pytest.raises(ValueError):
        gct.parse_time_range("24:00~24:00", "5분")
    with pytest.raises(ValueError):
        gct.parse_time_range("18:00~24:00", "1일")


def test_aggregates_5m_rows_to_15m_and_30m():
    rows = [
        _slot(timestamp="2026-04-12T17:00:00", value=1),
        _slot(timestamp="2026-04-12T17:05:00", value=2),
        _slot(timestamp="2026-04-12T17:10:00", value=3),
        _slot(timestamp="2026-04-12T17:15:00", value=4),
    ]

    rows_15m = gct.aggregate_slots_by_interval(rows, gct.parse_interval("15분"))
    rows_30m = gct.aggregate_slots_by_interval(rows, gct.parse_interval("30분"))

    assert [row["timestamp"] for row in rows_15m] == [
        "2026-04-12T17:00:00",
        "2026-04-12T17:15:00",
    ]
    assert [row["value"] for row in rows_15m] == [6, 4]
    assert [row["timestamp"] for row in rows_30m] == ["2026-04-12T17:00:00"]
    assert rows_30m[0]["value"] == 10


def test_sort_slots_uses_approach_and_drct_code_order():
    rows = [
        _slot(approach_id=20, drct_cd="02", drct_name="직"),
        _slot(approach_id=10, drct_cd="03", drct_name="우"),
        _slot(approach_id=10, drct_cd="01", drct_name="좌"),
    ]

    sorted_rows = gct.sort_slots(rows)

    assert [(row["approach_id"], row["drct_cd"]) for row in sorted_rows] == [
        (10, "01"),
        (10, "03"),
        (20, "02"),
    ]


def test_invalid_direction_rows_are_excluded():
    valid = _slot(drct_cd="01", drct_name="좌")
    drct_00 = _slot(drct_cd="00", drct_name="좌")
    empty_name = _slot(drct_cd="02", drct_name="")
    numeric_name = _slot(drct_cd="03", drct_name="03")
    english_name = _slot(drct_cd="04", drct_name="Left")

    assert gct.is_valid_direction_slot(valid)
    assert not gct.is_valid_direction_slot(drct_00)
    assert not gct.is_valid_direction_slot(empty_name)
    assert not gct.is_valid_direction_slot(numeric_name)
    assert not gct.is_valid_direction_slot(english_name)


def test_intersection_sheet_header_structure():
    rows = [
        _slot(drct_cd="01", drct_name="좌", value=10),
        _slot(drct_cd="02", drct_name="직", value=20),
    ]

    workbook = gct.create_workbook(rows, gct.parse_interval("5분"), gct.OUTPUT_MODE_INTERSECTION)
    sheet = workbook["교차로"]
    merged = {str(item) for item in sheet.merged_cells.ranges}

    assert sheet["A1"].value == "시간대"
    assert sheet["B1"].value == "교차로 이름"
    assert sheet["C1"].value == "동(서향)"
    assert [sheet.cell(2, col).value for col in range(3, 6)] == ["좌", "직", "합계"]
    assert "A1:A2" in merged
    assert "B1:B2" in merged
    assert "C1:E1" in merged
    assert [sheet.cell(3, col).value for col in range(1, 6)] == [
        "2026-04-12 08:00~08:05",
        "계남고가사거리",
        10,
        20,
        30,
    ]


def test_direction_sheet_header_structure_uses_full_approach_name():
    rows = [
        _slot(drct_cd="01", drct_name="좌", value=0),
        _slot(drct_cd="02", drct_name="직", value=20),
    ]

    workbook = gct.create_workbook(rows, gct.parse_interval("5분"), gct.OUTPUT_MODE_DIRECTION)
    sheet = workbook["방향"]

    assert [sheet.cell(1, col).value for col in range(1, 5)] == [
        "시간대",
        "방향 이름",
        "좌",
        "직",
    ]
    assert sheet["E1"].value == "합계"
    assert sheet["B2"].value == "계남고가사거리-동(서향)"
    assert [sheet.cell(2, col).value for col in range(3, 6)] == ["-", 20, 20]


def test_verify_api_contract_requires_corrected_endpoint(monkeypatch):
    monkeypatch.setattr(
        gct,
        "api_get",
        lambda path, base_url=gct.API_BASE_URL: {"paths": {"/intersections": {}}},
    )

    with pytest.raises(RuntimeError):
        gct.verify_api_contract()
