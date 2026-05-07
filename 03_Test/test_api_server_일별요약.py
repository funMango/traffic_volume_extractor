#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import os
import sys
import types
import unittest
from collections import defaultdict
from datetime import date
from unittest.mock import MagicMock, patch


sys.path.insert(0, os.path.dirname(__file__))


def _make_module(name: str) -> types.ModuleType:
    m = types.ModuleType(name)
    sys.modules[name] = m
    return m


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
        })
        _m.Scatter = type("Scatter", (), {"__init__": lambda s, **kw: None})

if "plotly.subplots" not in sys.modules:
    _subplots = _make_module("plotly.subplots")
    _subplots.make_subplots = lambda *a, **kw: None

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

_styles = _make_module("openpyxl.styles")
for _cls in ("Font", "PatternFill", "Alignment", "Border", "Side"):
    setattr(_styles, _cls, type(_cls, (), {"__init__": lambda s, **kw: None}))

_utils = _make_module("openpyxl.utils")
_utils.get_column_letter = lambda i: chr(64 + i)

if "fastapi" not in sys.modules:
    _fastapi = _make_module("fastapi")

    class _HTTPException(Exception):
        def __init__(self, status_code=None, detail=None):
            super().__init__(detail)
            self.status_code = status_code
            self.detail = detail

    class _FastAPI:
        def __init__(self, *args, **kwargs):
            pass

        def post(self, *args, **kwargs):
            def _decorator(fn):
                return fn
            return _decorator

        def get(self, *args, **kwargs):
            def _decorator(fn):
                return fn
            return _decorator

    _fastapi.FastAPI = _FastAPI
    _fastapi.HTTPException = _HTTPException


import api_server  # noqa: E402


class TestAggregateDailySummaryRows(unittest.TestCase):
    def test_aggregate_daily_summary_rows_node_and_approach(self):
        d = date(2026, 1, 1)
        target_data = {
            (1, d, 8): 100,
            (2, d, 8): 120,
            (1, d, 9): 80,
            (2, d, 9): 90,
        }
        node_results = [
            {"판정": "A형", "_acsr_id": 1, "_date": d, "_hour": 8, "방향": "북"},
            {"판정": "A형", "_acsr_id": 2, "_date": d, "_hour": 8, "방향": "남"},
            {"판정": "A형", "_acsr_id": 1, "_date": d, "_hour": 9, "방향": "북"},
        ]
        rows = api_server._aggregate_daily_summary_rows(1001, "테스트교차로", node_results, target_data)
        self.assertEqual(len(rows), 2)

        node_row = next(r for r in rows if r["approach_id"] is None)
        approach_row = next(r for r in rows if r["approach_id"] == 1)
        self.assertEqual(node_row["missing_count"], 1)
        self.assertEqual(approach_row["missing_count"], 1)

    def test_aggregate_daily_summary_rows_without_a_type(self):
        d = date(2026, 1, 1)
        target_data = {(1, d, 8): 100}
        node_results = [{"판정": "B형", "_acsr_id": 1, "_date": d, "_hour": 8, "방향": "북"}]
        rows = api_server._aggregate_daily_summary_rows(1, "N", node_results, target_data)
        self.assertEqual(rows, [])


class TestFetchAnomalyDailySummary(unittest.TestCase):
    def _req(self):
        return api_server.JobRequest(
            node_ids=[1001],
            date_start="2026-01-01",
            date_end="2026-01-01",
            hours=[8, 9],
        )

    def test_fetch_anomaly_daily_summary_success(self):
        d = date(2026, 1, 1)
        node_results = [
            {"판정": "A형", "_acsr_id": 1, "_date": d, "_hour": 8, "방향": "북"},
            {"판정": "A형", "_acsr_id": 2, "_date": d, "_hour": 8, "방향": "남"},
        ]
        target_data = {
            (1, d, 8): 100,
            (2, d, 8): 110,
        }

        api_server._holiday_dates = set()
        mock_conn = MagicMock()
        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024, 2025]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[]), \
             patch("api_server.ad.load_intersections", return_value=[(1001, "테스트교차로")]), \
             patch("api_server.ad.analyse_node", return_value=(node_results, {}, target_data)):
            result = api_server._fetch_anomaly_daily_summary(self._req())

        self.assertIn("rows", result)
        self.assertEqual(len(result["rows"]), 1)
        self.assertEqual(result["rows"][0]["missing_count"], 1)
        mock_conn.close.assert_called_once()

    def test_fetch_anomaly_daily_summary_accepts_string_node_id(self):
        d = date(2026, 1, 1)
        node_results = [
            {"판정": "A형", "_acsr_id": 1, "_date": d, "_hour": 8, "방향": "북"},
        ]
        target_data = {
            (1, d, 8): 100,
        }
        req = api_server.JobRequest(
            node_ids=["BC001N0044"],
            date_start="2026-01-01",
            date_end="2026-01-01",
            hours=[8],
        )

        api_server._holiday_dates = set()
        mock_conn = MagicMock()
        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024, 2025]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[]), \
             patch("api_server.ad.load_intersections", return_value=[("BC001N0044", "테스트교차로")]), \
             patch("api_server.ad.analyse_node", return_value=(node_results, {}, target_data)):
            result = api_server._fetch_anomaly_daily_summary(req)

        self.assertEqual(len(result["rows"]), 1)
        self.assertEqual(result["rows"][0]["node_id"], "BC001N0044")
        mock_conn.close.assert_called_once()

    def test_fetch_anomaly_daily_summary_requires_server_init(self):
        api_server._holiday_dates = None
        with self.assertRaises(RuntimeError):
            api_server._fetch_anomaly_daily_summary(self._req())


class TestAnomalyDailySummaryEndpoint(unittest.TestCase):
    def test_endpoint_function_returns_rows(self):
        rows = [{
            "date": "2026-01-01",
            "node_id": 1001,
            "node_name": "테스트교차로",
            "approach_id": None,
            "approach_name": None,
            "missing_count": 1,
        }]
        req = api_server.JobRequest(
            node_ids=[1001],
            date_start="2026-01-01",
            date_end="2026-01-01",
        )
        with patch.object(api_server, "_fetch_anomaly_daily_summary", return_value={"rows": rows}) as mock_fetch:
            result = asyncio.run(api_server.anomaly_daily_summary(req))
        self.assertIn("rows", result)
        self.assertEqual(len(result["rows"]), 1)
        mock_fetch.assert_called_once()

    def test_job_request_validation_error_for_missing_node_ids(self):
        with self.assertRaises(Exception):
            api_server.JobRequest(
                node_ids=[],
                date_start="2026-01-01",
                date_end="2026-01-01",
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
