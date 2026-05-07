#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for /raw-traffic-vknd without touching a real database."""

import sys
import types
import unittest
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from unittest.mock import MagicMock, patch


def _make_module(name: str) -> types.ModuleType:
    module = types.ModuleType(name)
    sys.modules[name] = module
    return module


PROGRAM_DIR = Path(__file__).resolve().parents[1] / "01_Program"
sys.path.insert(0, str(PROGRAM_DIR))

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


class _FakeCursor:
    def __init__(self, rows):
        self.rows = rows
        self.sql = None
        self.params = None
        self.arraysize = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def execute(self, sql, **params):
        self.sql = sql
        self.params = params

    def fetchall(self):
        return self.rows


class _FakeConn:
    def __init__(self, rows):
        self.cursor_obj = _FakeCursor(rows)

    def cursor(self):
        return self.cursor_obj


class TestRawTrafficVkndHelpers(unittest.TestCase):
    def test_interval_1h_uses_hourly_table(self):
        conn = _FakeConn([])
        api_server._load_raw_vknd_target_rows(
            conn,
            260322,
            date(2026, 3, 22),
            date(2026, 3, 22),
            [8],
            "1h",
        )

        self.assertIn("S_CRSRD_VKND_TRF_1HH", conn.cursor_obj.sql)
        self.assertNotIn("S_CRSRD_VKND_TRF_15MI", conn.cursor_obj.sql)
        self.assertEqual(conn.cursor_obj.params["h0"], 8)

    def test_interval_15m_uses_15minute_table_and_vknd_filter(self):
        conn = _FakeConn([])
        api_server._load_raw_vknd_target_rows(
            conn,
            260322,
            date(2026, 3, 22),
            date(2026, 3, 22),
            list(range(24)),
            "15m",
            ["1", "2"],
        )

        self.assertIn("S_CRSRD_VKND_TRF_15MI", conn.cursor_obj.sql)
        self.assertNotIn("TO_CHAR(TOT_DT, 'HH24')) IN", conn.cursor_obj.sql)
        self.assertEqual(conn.cursor_obj.params["vknd0"], "1")
        self.assertEqual(conn.cursor_obj.params["vknd1"], "2")

    def test_aggregates_are_split_by_vehicle_kind_and_preserve_null(self):
        vknd_slots = [
            {
                "interval": "1h",
                "timestamp": "2026-03-22T08:00:00",
                "date": "2026-03-22",
                "hour": 8,
                "minute": 0,
                "node_id": 260322,
                "node_name": "Node",
                "approach_id": 2001,
                "approach_name": "North",
                "vknd_cd": "1",
                "vknd_name": "Car",
                "traffic_volume": 100,
            },
            {
                "interval": "1h",
                "timestamp": "2026-03-22T08:00:00",
                "date": "2026-03-22",
                "hour": 8,
                "minute": 0,
                "node_id": 260322,
                "node_name": "Node",
                "approach_id": 2001,
                "approach_name": "North",
                "vknd_cd": "2",
                "vknd_name": "Bus",
                "traffic_volume": None,
            },
            {
                "interval": "1h",
                "timestamp": "2026-03-22T08:00:00",
                "date": "2026-03-22",
                "hour": 8,
                "minute": 0,
                "node_id": 260322,
                "node_name": "Node",
                "approach_id": 2002,
                "approach_name": "South",
                "vknd_cd": "1",
                "vknd_name": "Car",
                "traffic_volume": 50,
            },
        ]

        slots = api_server._aggregate_raw_slots_from_vknd(vknd_slots)
        node_slots = api_server._aggregate_raw_node_slots_from_vknd(slots)

        self.assertEqual(len(slots), 3)
        self.assertEqual(next(s for s in node_slots if s["vknd_cd"] == "1")["traffic_volume"], 150)
        self.assertIsNone(next(s for s in node_slots if s["vknd_cd"] == "2")["traffic_volume"])


class TestFetchRawTrafficVknd(unittest.TestCase):
    def _make_request(self, interval="1h", vknd_codes=None):
        return api_server.RawTrafficVkndRequest(
            node_ids=[260322],
            date_start="2026-03-22",
            date_end="2026-03-22",
            hours=[8],
            interval=interval,
            vknd_codes=vknd_codes,
        )

    def test_fetch_returns_all_vehicle_kinds_when_filter_is_omitted(self):
        mock_conn = MagicMock()
        rows = [
            (2001, "1", datetime(2026, 3, 22, 8, 0), 100),
            (2001, "2", datetime(2026, 3, 22, 8, 0), 20),
        ]

        with patch.object(api_server.ad, "connect_db", return_value=mock_conn), \
             patch.object(api_server, "_get_id_to_name_map", return_value={260322: "Node"}), \
             patch.object(api_server, "_load_vknd_code_map", return_value={"1": "Car", "2": "Bus"}), \
             patch.object(api_server.ad, "load_approaches", return_value=[(2001, "North")]), \
             patch.object(api_server, "_load_raw_vknd_target_rows", return_value=rows) as mock_rows:
            result = api_server._fetch_raw_traffic_vknd(self._make_request())

        self.assertEqual(len(result["vknd_slots"]), 2)
        self.assertIsNone(mock_rows.call_args.args[6])
        self.assertEqual({s["vknd_cd"] for s in result["vknd_slots"]}, {"1", "2"})
        mock_conn.close.assert_called_once()

    def test_fetch_filters_vehicle_kinds_and_preserves_15minute_minute(self):
        mock_conn = MagicMock()

        def fake_rows(*args):
            vknd_codes = args[6]
            self.assertEqual(vknd_codes, ["2"])
            return [(2001, "2", datetime(2026, 3, 22, 8, 15), 20)]

        with patch.object(api_server.ad, "connect_db", return_value=mock_conn), \
             patch.object(api_server, "_get_id_to_name_map", return_value={260322: "Node"}), \
             patch.object(api_server, "_load_vknd_code_map", return_value={"2": "Bus"}), \
             patch.object(api_server.ad, "load_approaches", return_value=[(2001, "North")]), \
             patch.object(api_server, "_load_raw_vknd_target_rows", side_effect=fake_rows):
            result = api_server._fetch_raw_traffic_vknd(self._make_request("15m", [2]))

        slot = result["vknd_slots"][0]
        self.assertEqual(slot["interval"], "15m")
        self.assertEqual(slot["minute"], 15)
        self.assertEqual(slot["vknd_cd"], "2")
        self.assertEqual(result["slots"][0]["traffic_volume"], 20)
        self.assertEqual(result["node_slots"][0]["traffic_volume"], 20)


class TestRawTrafficVkndEndpoint(unittest.TestCase):
    def test_endpoint_returns_vknd_payload(self):
        payload = {"vknd_slots": [], "slots": [], "node_slots": []}
        with patch.object(api_server, "_fetch_raw_traffic_vknd", return_value=payload) as mock_fetch:
            resp = _client.post(
                "/raw-traffic-vknd",
                json={
                    "node_ids": [260322],
                    "date_start": "2026-03-22",
                    "date_end": "2026-03-22",
                    "interval": "1h",
                },
            )

        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.json(), payload)
        self.assertEqual(mock_fetch.call_count, 1)

    def test_invalid_interval_returns_422(self):
        resp = _client.post(
            "/raw-traffic-vknd",
            json={
                "node_ids": [260322],
                "date_start": "2026-03-22",
                "date_end": "2026-03-22",
                "interval": "5m",
            },
        )

        self.assertEqual(resp.status_code, 422)


class TestVehicleKindsEndpoint(unittest.TestCase):
    def test_fetch_vehicle_kinds_uses_code_map_and_closes_connection(self):
        mock_conn = MagicMock()
        with patch.object(api_server.ad, "connect_db", return_value=mock_conn), \
             patch.object(api_server, "_load_vknd_code_map", return_value={"2": "Bus", "1": "Car"}):
            result = api_server._fetch_vehicle_kinds()

        self.assertEqual(result, [{"code": "1", "name": "Car"}, {"code": "2", "name": "Bus"}])
        mock_conn.close.assert_called_once()

    def test_endpoint_returns_vehicle_kind_payload(self):
        items = [{"code": "1", "name": "Car"}]
        with patch.object(api_server, "_fetch_vehicle_kinds", return_value=items):
            resp = _client.get("/vehicle-kinds")

        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.json(), {"count": 1, "items": items})


class TestCorrectedTrafficVkndHelpers(unittest.TestCase):
    def test_hourly_aggregation_uses_corrected_15minute_values(self):
        slots_15m = [
            {
                "interval": "15m",
                "timestamp": "2026-03-22T08:00:00",
                "date": "2026-03-22",
                "hour": 8,
                "minute": 0,
                "node_id": 260322,
                "node_name": "Node",
                "approach_id": 2001,
                "approach_name": "North",
                "vknd_cd": "1",
                "vknd_name": "Car",
                "traffic_volume": 0,
                "anomaly_type": "A형",
                "corrected_value": 100,
                "correction_method": "median",
                "confidence": "OK",
                "baseline": {"median": 100},
            },
            {
                "interval": "15m",
                "timestamp": "2026-03-22T08:15:00",
                "date": "2026-03-22",
                "hour": 8,
                "minute": 15,
                "node_id": 260322,
                "node_name": "Node",
                "approach_id": 2001,
                "approach_name": "North",
                "vknd_cd": "1",
                "vknd_name": "Car",
                "traffic_volume": 50,
                "anomaly_type": None,
                "corrected_value": None,
                "correction_method": None,
                "confidence": None,
                "baseline": None,
            },
            {
                "interval": "15m",
                "timestamp": "2026-03-22T08:30:00",
                "date": "2026-03-22",
                "hour": 8,
                "minute": 30,
                "node_id": 260322,
                "node_name": "Node",
                "approach_id": 2001,
                "approach_name": "North",
                "vknd_cd": "1",
                "vknd_name": "Car",
                "traffic_volume": 5,
                "anomaly_type": "B형",
                "corrected_value": 80,
                "correction_method": "선형보간",
                "confidence": "LOW",
                "baseline": {"median": 80},
            },
            {
                "interval": "15m",
                "timestamp": "2026-03-22T08:45:00",
                "date": "2026-03-22",
                "hour": 8,
                "minute": 45,
                "node_id": 260322,
                "node_name": "Node",
                "approach_id": 2001,
                "approach_name": "North",
                "vknd_cd": "1",
                "vknd_name": "Car",
                "traffic_volume": 60,
                "anomaly_type": None,
                "corrected_value": None,
                "correction_method": None,
                "confidence": None,
                "baseline": None,
            },
        ]

        hourly = api_server._aggregate_vknd_15m_to_1h(slots_15m)

        self.assertEqual(len(hourly), 1)
        self.assertEqual(hourly[0]["interval"], "1h")
        self.assertIsNone(hourly[0]["minute"])
        self.assertEqual(hourly[0]["traffic_volume"], 115)
        self.assertEqual(hourly[0]["corrected_value"], 290)
        self.assertEqual(hourly[0]["anomaly_type"], "A+B혼합")
        self.assertEqual(hourly[0]["confidence"], "LOW")

    def test_detects_missing_15minute_slot_and_preserves_minute(self):
        d = date(2026, 3, 22)
        cell = (2001, "1", 2, 8, 15, 3)
        baselines = {
            cell: {
                "zero_rate": 0.0,
                "stats": {"median": 100.0, "q1": 90.0, "q3": 110.0, "n_clean": 12},
                "expansion_stage": 0,
                "n_fallback": 0,
            }
        }

        rows = api_server._detect_vknd_15m_anomalies(
            "Node",
            [(2001, "North")],
            ["1"],
            {"1": "Car"},
            baselines,
            {},
            d,
            d,
            [8],
            {"2026-03-22"},
        )
        api_server._apply_vknd_15m_corrections(rows, baselines, {})

        slot = api_server._serialize_vknd_slot(
            rows[0],
            260322,
            "Node",
            {2001: "North"},
            {"1": "Car"},
            baselines,
        )
        self.assertEqual(slot["minute"], 15)
        self.assertEqual(slot["anomaly_type"], "A형")
        self.assertEqual(slot["corrected_value"], 100)
        self.assertEqual(slot["baseline"]["median"], 100.0)


class TestCorrectedTrafficVkndEndpoint(unittest.TestCase):
    def test_openapi_contains_corrected_vknd_endpoint(self):
        resp = _client.get("/openapi.json")

        self.assertEqual(resp.status_code, 200)
        self.assertIn("/corrected-traffic-vknd", resp.json()["paths"])

    def test_endpoint_returns_corrected_vknd_payload(self):
        payload = {"vknd_slots": [], "slots": [], "node_slots": []}
        with patch.object(api_server, "_fetch_corrected_traffic_vknd", return_value=payload) as mock_fetch:
            resp = _client.post(
                "/corrected-traffic-vknd",
                json={
                    "node_ids": [260322],
                    "date_start": "2026-03-22",
                    "date_end": "2026-03-22",
                    "interval": "15m",
                },
            )

        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.json(), payload)
        self.assertEqual(mock_fetch.call_count, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
