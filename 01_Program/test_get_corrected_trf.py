import importlib.util
import sqlite3
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
    anomaly_type=None,
    correction_method=None,
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
        "anomaly_type": anomaly_type,
        "correction_method": correction_method,
        "is_corrected": gct.is_corrected_slot(
            {"anomaly_type": anomaly_type, "correction_method": correction_method}
        ),
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


def test_api_request_json_waits_without_timeout(monkeypatch):
    calls = []

    class FakeResponse:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, traceback):
            return False

        def read(self):
            return b'{"ok": true}'

    def fake_urlopen(req, timeout=None):
        calls.append((req, timeout))
        return FakeResponse()

    monkeypatch.setattr(gct.urllib.request, "urlopen", fake_urlopen)

    response = gct._api_request_json("GET", "/health", base_url="http://api.local")

    assert response == {"ok": True}
    assert calls == [("http://api.local/health", None)]


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


def test_parse_time_ranges_supports_multiple_ranges_and_api_hour_union():
    time_ranges = gct.parse_time_ranges("09:00~10:00, 12:00~14:00", "1시간")

    assert [(item.start_minute, item.end_minute) for item in time_ranges] == [
        (9 * 60, 10 * 60),
        (12 * 60, 14 * 60),
    ]
    assert gct.hours_for_api(time_ranges) == [9, 12, 13]


def test_parse_time_ranges_preserves_single_range_and_full_day_behavior():
    single = gct.parse_time_ranges("17:15~17:45", "15분")
    full_day = gct.parse_time_ranges("24시간", "5분")

    assert [(item.start_minute, item.end_minute) for item in single] == [
        (17 * 60 + 15, 17 * 60 + 45)
    ]
    assert len(full_day) == 1
    assert full_day[0].is_full_day
    assert gct.hours_for_api(full_day) == list(range(24))


def test_parse_time_ranges_rejects_invalid_multiple_range_inputs():
    invalid_inputs = [
        "",
        "09:00",
        "10:00~09:00",
        "24시간, 09:00~10:00",
        "09:15~10:00",
    ]

    for raw in invalid_inputs:
        with pytest.raises(ValueError):
            gct.parse_time_ranges(raw, "1시간")


def test_parse_output_mode_accepts_approach_aliases():
    assert gct.parse_output_mode("1") == gct.OUTPUT_MODE_INTERSECTION
    assert gct.parse_output_mode("교차로 교통량") == gct.OUTPUT_MODE_INTERSECTION
    assert gct.parse_output_mode("2") == gct.OUTPUT_MODE_DIRECTION
    assert gct.parse_output_mode("방향별 교통량") == gct.OUTPUT_MODE_DIRECTION
    assert gct.parse_output_mode("3") == gct.OUTPUT_MODE_APPROACH
    assert gct.parse_output_mode("접근로") == gct.OUTPUT_MODE_APPROACH
    assert gct.parse_output_mode("접근로 교통량") == gct.OUTPUT_MODE_APPROACH


def test_make_output_path_uses_db_extension_only_for_approach(tmp_path):
    intersections = [gct.Intersection(1, "계남고가사거리", 0)]
    periods = gct.parse_periods("260412")
    interval = gct.parse_interval("5분")

    intersection_path = gct.make_output_path(
        intersections,
        periods,
        interval,
        gct.OUTPUT_MODE_INTERSECTION,
        output_dir=tmp_path,
    )
    direction_path = gct.make_output_path(
        intersections,
        periods,
        interval,
        gct.OUTPUT_MODE_DIRECTION,
        output_dir=tmp_path,
    )
    approach_path = gct.make_output_path(
        intersections,
        periods,
        interval,
        gct.OUTPUT_MODE_APPROACH,
        output_dir=tmp_path,
    )

    assert intersection_path.suffix == ".xlsx"
    assert "_교차로_" in intersection_path.name
    assert direction_path.suffix == ".xlsx"
    assert "_방향_" in direction_path.name
    assert approach_path.suffix == ".db"
    assert "_접근로_" in approach_path.name


def test_slot_in_time_ranges_includes_only_configured_ranges():
    time_ranges = gct.parse_time_ranges("09:00~10:00, 12:00~14:00", "1시간")

    assert gct.slot_in_time_ranges(_slot(timestamp="2026-04-12T09:00:00"), time_ranges)
    assert not gct.slot_in_time_ranges(_slot(timestamp="2026-04-12T10:00:00"), time_ranges)
    assert not gct.slot_in_time_ranges(_slot(timestamp="2026-04-12T11:00:00"), time_ranges)
    assert gct.slot_in_time_ranges(_slot(timestamp="2026-04-12T13:00:00"), time_ranges)
    assert not gct.slot_in_time_ranges(_slot(timestamp="2026-04-12T14:00:00"), time_ranges)


def test_fetch_corrected_direction_slots_uses_hour_union_and_filters_time_ranges(monkeypatch):
    periods = gct.parse_periods("260412")
    intersections = [gct.Intersection(1, "계남고가사거리", 0)]
    interval = gct.parse_interval("1시간")
    time_ranges = gct.parse_time_ranges("09:00~10:00, 12:00~14:00", interval)
    calls = []

    def fake_api_post(path, payload, base_url=gct.API_BASE_URL):
        calls.append((path, payload, base_url))
        return {
            "drct_slots": [
                {
                    "interval": "1h",
                    "timestamp": f"{payload['date_start']}T{hour:02d}:00:00",
                    "node_id": 1,
                    "node_name": "계남고가사거리",
                    "approach_id": 10,
                    "approach_name": "계남고가사거리(상향)",
                    "drct_cd": "01",
                    "drct_name": "좌",
                    "corrected_value": hour,
                }
                for hour in [9, 10, 11, 12, 13, 14]
            ]
        }

    monkeypatch.setattr(gct, "api_post", fake_api_post)

    rows = gct.fetch_corrected_direction_slots(
        intersections,
        periods,
        interval,
        time_ranges,
    )

    assert calls[0][1]["hours"] == [9, 12, 13]
    assert [row["hour"] for row in rows] == [9, 12, 13]


def test_aggregates_5m_rows_to_15m_and_30m():
    rows = [
        _slot(timestamp="2026-04-12T17:00:00", value=1),
        _slot(timestamp="2026-04-12T17:05:00", value=2, anomaly_type="이상치"),
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
    assert [row["is_corrected"] for row in rows_15m] == [True, False]
    assert [row["timestamp"] for row in rows_30m] == ["2026-04-12T17:00:00"]
    assert rows_30m[0]["value"] == 10
    assert rows_30m[0]["is_corrected"] is True


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


def test_create_workbook_splits_multiple_intersections_into_sheets():
    rows = [
        _slot(node_key="1", node_name="Alpha", node_order=0, approach_id=10, value=11),
        _slot(node_key="2", node_name="Beta", node_order=1, approach_id=20, value=22),
    ]

    workbook = gct.create_workbook(
        rows,
        gct.IntervalSpec("5m", "5m", 5),
        gct.OUTPUT_MODE_INTERSECTION,
    )

    assert workbook.sheetnames == ["Alpha", "Beta"]
    assert workbook["Alpha"]["B3"].value == "Alpha"
    assert workbook["Alpha"]["C3"].value == 11
    assert workbook["Beta"]["B3"].value == "Beta"
    assert workbook["Beta"]["C3"].value == 22


def test_create_direction_workbook_splits_multiple_intersections_into_sheets():
    rows = [
        _slot(
            node_key="1",
            node_name="Alpha",
            node_order=0,
            approach_id=10,
            approach_name="Alpha-North",
            value=11,
        ),
        _slot(
            node_key="2",
            node_name="Beta",
            node_order=1,
            approach_id=20,
            approach_name="Beta-South",
            value=22,
        ),
    ]

    workbook = gct.create_workbook(
        rows,
        gct.IntervalSpec("5m", "5m", 5),
        gct.OUTPUT_MODE_DIRECTION,
    )

    assert workbook.sheetnames == ["Alpha", "Beta"]
    assert workbook["Alpha"]["B2"].value == "Alpha-North"
    assert workbook["Alpha"]["C2"].value == 11
    assert workbook["Beta"]["B2"].value == "Beta-South"
    assert workbook["Beta"]["C2"].value == 22


def test_safe_sheet_title_handles_invalid_and_duplicate_names():
    used_titles = set()
    long_name = "A" * 40

    assert gct.safe_sheet_title("Bad[]:*?/\\Name", used_titles) == "Bad_______Name"
    assert gct.safe_sheet_title(long_name, used_titles) == "A" * 31
    assert gct.safe_sheet_title(long_name, used_titles) == f"{'A' * 29}_2"


def test_save_approach_database_writes_vertical_rows_and_correction_flags(tmp_path):
    output_path = tmp_path / "approach.db"
    rows = [
        _slot(
            timestamp="2026-04-12T08:00:00",
            approach_name="계남고가사거리-동(서향)",
            drct_name="좌",
            value=10,
            anomaly_type="정상",
        ),
        _slot(
            timestamp="2026-04-12T08:05:00",
            approach_name="계남고가사거리-동(서향)",
            drct_cd="02",
            drct_name="직",
            value=20,
            correction_method="선형보간",
        ),
    ]

    gct.save_approach_database(rows, gct.parse_interval("5분"), output_path)

    with sqlite3.connect(output_path) as conn:
        table_names = [
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY name"
            )
        ]
        columns = [row[1] for row in conn.execute('PRAGMA table_info("접근로")')]
        saved_rows = conn.execute(
            """
            SELECT
                "시간대",
                "교차로이름",
                "방향이름",
                "접근로",
                "교통량",
                "보정"
            FROM "접근로"
            ORDER BY "시간대", "접근로"
            """
        ).fetchall()

    assert table_names == ["접근로"]
    assert columns == ["시간대", "교차로이름", "방향이름", "접근로", "교통량", "보정"]
    assert saved_rows == [
        ("2026-04-12 08:00~08:05", "계남고가사거리", "동(서향)", "좌", 10, 0),
        ("2026-04-12 08:05~08:10", "계남고가사거리", "동(서향)", "직", 20, 1),
    ]


def test_verify_api_contract_requires_corrected_endpoint(monkeypatch):
    monkeypatch.setattr(
        gct,
        "api_get",
        lambda path, base_url=gct.API_BASE_URL: {"paths": {"/intersections": {}}},
    )

    with pytest.raises(RuntimeError):
        gct.verify_api_contract()
