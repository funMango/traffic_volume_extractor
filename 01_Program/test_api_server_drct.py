#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""api_server DRCT 엔드포인트 단위 테스트."""

import os
import sys
import types
import unittest
from collections import defaultdict
from datetime import date
from unittest.mock import MagicMock, patch


def _make_module(name: str) -> types.ModuleType:
    m = types.ModuleType(name)
    sys.modules[name] = m
    return m


sys.path.insert(0, os.path.dirname(__file__))

if "oracledb" not in sys.modules:
    _make_module("oracledb")

if "dotenv" not in sys.modules:
    _dotenv = _make_module("dotenv")
    _dotenv.load_dotenv = lambda *a, **kw: None  # type: ignore

for _pmod in ("plotly", "plotly.graph_objects"):
    if _pmod not in sys.modules:
        _m = _make_module(_pmod)
        _m.Figure = type("Figure", (), {
            "__init__": lambda s, **kw: None,
            "add_trace": lambda s, *a, **kw: None,
            "update_layout": lambda s, **kw: None,
            "write_html": lambda s, *a, **kw: None,
            "to_html": lambda s, **kw: "<div></div>",
        })
        _m.Scatter = type("Scatter", (), {"__init__": lambda s, **kw: None})

if "plotly.subplots" not in sys.modules:
    _subplots = _make_module("plotly.subplots")
    _subplots.make_subplots = lambda *a, **kw: None  # type: ignore

if "openpyxl" not in sys.modules:
    _openpyxl = _make_module("openpyxl")
    _openpyxl.Workbook = type("Workbook", (), {
        "__init__": lambda s: None,
        "active": property(lambda s: type("WS", (), {
            "title": "",
            "append": lambda s, r: None,
            "column_dimensions": defaultdict(lambda: type("CD", (), {"width": 0})()),
        })()),
        "create_sheet": lambda s, t: None,
        "save": lambda s, p: None,
        "remove": lambda s, ws: None,
    })
if "openpyxl.styles" not in sys.modules:
    _styles = _make_module("openpyxl.styles")
    for _cls in ("Font", "PatternFill", "Alignment", "Border", "Side"):
        setattr(_styles, _cls, type(_cls, (), {"__init__": lambda s, **kw: None}))
if "openpyxl.utils" not in sys.modules:
    _utils = _make_module("openpyxl.utils")
    _utils.get_column_letter = lambda i: chr(64 + i)


import api_server  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

_client = TestClient(api_server.app, raise_server_exceptions=True)


class TestFetchCorrectedTrafficDrct(unittest.TestCase):
    def setUp(self):
        api_server._holiday_dates = set()

    def _make_request(self):
        return api_server.JobRequest(
            node_ids=[260322],
            date_start="2026-03-22",
            date_end="2026-03-22",
            hours=[7, 8],
        )

    def test_returns_drct_slots_and_acsr_aggregates(self):
        d = date(2026, 3, 22)
        node_results = [
            {
                "날짜": "2026.03.22",
                "시간": "08:00",
                "교차로": "송내사거리",
                "방향": "좌회전",
                "교통량": 0,
                "판정": "A형",
                "_acsr_id": (2001, "01"),
                "_date": d,
                "_hour": 8,
                "_cell": ((2001, "01"), 1, 8, 3),
                "보정값": 120,
                "보정방법": "median",
                "신뢰도": "OK",
            }
        ]
        baselines = {
            ((2001, "01"), 1, 8, 3): {
                "zero_rate": 0.1,
                "expansion_stage": 0,
                "stats": {"median": 120.0, "q1": 110.0, "q3": 130.0, "n_clean": 14},
            }
        }
        target_data = {
            ((2001, "01"), d, 7): 100,
            ((2001, "02"), d, 7): 80,
        }
        drct_meta = {
            (2001, "01"): {
                "approach_id": 2001,
                "approach_name": "북향",
                "drct_cd": "01",
                "drct_name": "좌회전",
            },
            (2001, "02"): {
                "approach_id": 2001,
                "approach_name": "북향",
                "drct_cd": "02",
                "drct_name": "직진",
            },
        }

        with patch.object(api_server.ad, "connect_db", return_value=MagicMock()), \
             patch.object(api_server, "_get_id_to_name_map", return_value={260322: "송내사거리"}), \
             patch.object(api_server.ad, "select_baseline_years", return_value=[2024, 2025]), \
             patch.object(api_server.ad, "select_fallback_year", return_value=2025), \
             patch.object(api_server.ad, "get_adjacent_month_periods", return_value=[]), \
             patch.object(
                 api_server.ad,
                 "analyse_node_drct",
                 return_value=(node_results, baselines, target_data, drct_meta),
             ):
            result = api_server._fetch_corrected_traffic_drct(self._make_request())

        self.assertIn("slots", result)
        self.assertIn("drct_slots", result)
        self.assertEqual(len(result["drct_slots"]), 3)  # target 2 + 결측보정 1

        h8_left = next(
            s for s in result["drct_slots"]
            if s["hour"] == 8 and s["drct_cd"] == "01"
        )
        self.assertEqual(h8_left["traffic_volume"], 0)
        self.assertEqual(h8_left["corrected_value"], 120)
        self.assertEqual(h8_left["drct_name"], "좌회전")

        acsr_h7 = next(s for s in result["slots"] if s["hour"] == 7 and s["approach_id"] == 2001)
        self.assertEqual(acsr_h7["traffic_volume"], 180)
        self.assertEqual(acsr_h7["corrected_value"], 180)
        self.assertIsNone(acsr_h7["anomaly_type"])

        acsr_h8 = next(s for s in result["slots"] if s["hour"] == 8 and s["approach_id"] == 2001)
        self.assertEqual(acsr_h8["traffic_volume"], 0)
        self.assertEqual(acsr_h8["corrected_value"], 120)
        self.assertEqual(acsr_h8["anomaly_type"], "A형")

    def test_aggregate_marks_mixed_and_ignores_drct_00(self):
        rows = api_server._aggregate_acsr_slots_from_drct([
            {
                "date": "2026-03-22",
                "hour": 8,
                "node_id": 260322,
                "node_name": "송내사거리",
                "approach_id": 2001,
                "approach_name": "북향",
                "drct_cd": "00",
                "traffic_volume": 999,
                "anomaly_type": "B형",
                "corrected_value": 999,
            },
            {
                "date": "2026-03-22",
                "hour": 8,
                "node_id": 260322,
                "node_name": "송내사거리",
                "approach_id": 2001,
                "approach_name": "북향",
                "drct_cd": "01",
                "traffic_volume": 0,
                "anomaly_type": "A형",
                "corrected_value": 120,
            },
            {
                "date": "2026-03-22",
                "hour": 8,
                "node_id": 260322,
                "node_name": "송내사거리",
                "approach_id": 2001,
                "approach_name": "북향",
                "drct_cd": "02",
                "traffic_volume": 30,
                "anomaly_type": "B형",
                "corrected_value": 95,
            },
        ])
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["traffic_volume"], 30)
        self.assertEqual(row["corrected_value"], 215)
        self.assertEqual(row["anomaly_type"], "A+B혼합")

    def test_fetch_drct_excludes_drct_00_from_outputs(self):
        d = date(2026, 3, 22)
        node_results = [
            {
                "날짜": "2026.03.22",
                "시간": "08:00",
                "교차로": "송내사거리",
                "방향": "미분류",
                "교통량": 0,
                "판정": "A형",
                "_acsr_id": (2001, "00"),
                "_date": d,
                "_hour": 8,
                "_cell": ((2001, "00"), 1, 8, 3),
                "보정값": 100,
                "보정방법": "median",
                "신뢰도": "OK",
            }
        ]
        baselines = {
            ((2001, "00"), 1, 8, 3): {
                "zero_rate": 0.1,
                "expansion_stage": 0,
                "stats": {"median": 100.0, "q1": 90.0, "q3": 110.0, "n_clean": 14},
            }
        }
        target_data = {
            ((2001, "00"), d, 8): 77,
        }
        drct_meta = {
            (2001, "00"): {
                "approach_id": 2001,
                "approach_name": "북향",
                "drct_cd": "00",
                "drct_name": "미분류",
            }
        }

        with patch.object(api_server.ad, "connect_db", return_value=MagicMock()), \
             patch.object(api_server, "_get_id_to_name_map", return_value={260322: "송내사거리"}), \
             patch.object(api_server.ad, "select_baseline_years", return_value=[2024, 2025]), \
             patch.object(api_server.ad, "select_fallback_year", return_value=2025), \
             patch.object(api_server.ad, "get_adjacent_month_periods", return_value=[]), \
             patch.object(
                 api_server.ad,
                 "analyse_node_drct",
                 return_value=(node_results, baselines, target_data, drct_meta),
             ):
            result = api_server._fetch_corrected_traffic_drct(self._make_request())

        self.assertEqual(result["drct_slots"], [])
        self.assertEqual(result["slots"], [])


class TestCorrectedTrafficDrctEndpoint(unittest.TestCase):
    def setUp(self):
        api_server._holiday_dates = set()

    def test_endpoint_returns_slots_and_drct_slots(self):
        with patch.object(
            api_server,
            "_fetch_corrected_traffic_drct",
            return_value={"slots": [], "drct_slots": []},
        ) as mock_fetch:
            resp = _client.post(
                "/corrected-traffic-drct",
                json={
                    "node_ids": [260322],
                    "date_start": "2026-03-22",
                    "date_end": "2026-03-22",
                },
            )

        self.assertEqual(resp.status_code, 200)
        body = resp.json()
        self.assertIn("slots", body)
        self.assertIn("drct_slots", body)
        self.assertEqual(body["slots"], [])
        self.assertEqual(body["drct_slots"], [])
        self.assertEqual(mock_fetch.call_count, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
