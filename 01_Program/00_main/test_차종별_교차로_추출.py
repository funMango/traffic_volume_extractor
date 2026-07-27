from __future__ import annotations

import importlib.util
import re
import sys
from pathlib import Path

import pytest


MODULE_PATH = Path(__file__).with_name("차종별_교차로_추출.py")
SPEC = importlib.util.spec_from_file_location("vehicle_intersection_extract", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
vehicle = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = vehicle
SPEC.loader.exec_module(vehicle)


def test_parse_time_range_supports_hourly_and_quarter_hour_slots():
    assert vehicle.parse_time_range("07~09") == [420, 480]
    assert vehicle.parse_time_range("07:00~09:00") == [420, 480]
    assert vehicle.parse_time_range("07:15~09:30", vehicle.FIFTEEN_MINUTE_AGGREGATION) == [
        435,
        450,
        465,
        480,
        495,
        510,
        525,
        540,
        555,
    ]
    assert vehicle.parse_time_range("24시간", vehicle.FIFTEEN_MINUTE_AGGREGATION) == list(
        range(0, 24 * 60, 15)
    )


@pytest.mark.parametrize(
    ("raw", "aggregation_unit", "message"),
    [
        ("07:15~09:00", vehicle.ONE_HOUR_AGGREGATION, ":00 경계"),
        ("07:10~09:00", vehicle.FIFTEEN_MINUTE_AGGREGATION, ":15, :30, :45 경계"),
    ],
)
def test_parse_time_range_rejects_unaligned_boundaries(raw, aggregation_unit, message):
    with pytest.raises(ValueError, match=re.escape(message)):
        vehicle.parse_time_range(raw, aggregation_unit)


def test_fifteen_minute_sql_uses_selected_table_and_minute_slots():
    sql, params = vehicle.build_fetch_rows_sql_params(
        vehicle.Intersection(node_id=7, name="A교차로"),
        vehicle.parse_periods("260501")[0],
        vehicle.parse_time_range("07:15~09:30", vehicle.FIFTEEN_MINUTE_AGGREGATION),
        {"2": "세단"},
        vehicle.FIFTEEN_MINUTE_AGGREGATION,
    )
    compact_sql = re.sub(r"\s+", " ", sql)

    assert "FROM S_CRSRD_VKND_TRF_15MI v" in compact_sql
    assert "TO_CHAR(v.TOT_DT, 'HH24:MI') IN" in compact_sql
    assert [params[f"time_slot{index}"] for index in range(9)] == [
        "07:15",
        "07:30",
        "07:45",
        "08:00",
        "08:15",
        "08:30",
        "08:45",
        "09:00",
        "09:15",
    ]


def test_hourly_sql_keeps_hour_source_and_hour_filter():
    sql, params = vehicle.build_fetch_all_intersections_periods_sql_params(
        vehicle.parse_periods("260501"),
        vehicle.parse_time_range("07~09"),
        {"2": "세단"},
        vehicle.ONE_HOUR_AGGREGATION,
    )
    compact_sql = re.sub(r"\s+", " ", sql)

    assert f"FROM {vehicle.TRAFFIC_TABLE} v" in compact_sql
    assert "TO_NUMBER(TO_CHAR(v.TOT_DT, 'HH24')) IN" in compact_sql
    assert params["hour0"] == 7
    assert params["hour1"] == 8


def test_required_column_validation_uses_selected_traffic_table(monkeypatch):
    checked_tables: list[str] = []

    def fake_columns(_cursor, table_name):
        checked_tables.append(table_name)
        return {
            "NODE_ID",
            "TOT_DT",
            "ACSR_ID",
            "DRCT_CD",
            "VKND_CD",
            "TRF_QNTY",
            "CRSRD_NM",
            "ACSR_NM",
            "GRP_CD",
            "CD",
            "CD_NM",
        }

    monkeypatch.setattr(vehicle, "get_table_columns", fake_columns)

    vehicle.validate_vehicle_extract_tables(object(), vehicle.FIFTEEN_MINUTE_AGGREGATION)

    assert checked_tables[0] == "S_CRSRD_VKND_TRF_15MI"
