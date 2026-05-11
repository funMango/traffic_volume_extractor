#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for the new traffic_api Clean Architecture entry point."""

import os
import sys
import types
import unittest
from collections import defaultdict
from datetime import date
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


from fastapi.testclient import TestClient  # noqa: E402

from traffic_api.application.job_use_cases import JobUseCase  # noqa: E402
from traffic_api.domain.aggregation import aggregate_acsr_slots_from_drct  # noqa: E402
from traffic_api.infrastructure.legacy_analysis_gateway import LegacyTrafficAnalysisGateway  # noqa: E402
from traffic_api.infrastructure.memory_job_store import InMemoryJobStore  # noqa: E402
from traffic_api.main import create_app  # noqa: E402
from traffic_api.presentation.schemas import JobRequest, RawTrafficDrctRequest, RawTrafficVkndRequest  # noqa: E402


class _ImmediateRunner:
    def submit(self, fn, *args):
        fn(*args)

    def shutdown(self):
        pass


class _Gateway:
    def __init__(self):
        self.jobs = []

    def run_job(self, job):
        self.jobs.append(job["job_id"])
        job["status"] = "done"
        job["result"] = {"slots": []}


class TestRequestModels(unittest.TestCase):
    def test_job_request_resolves_peak_hours(self):
        req = JobRequest(
            node_ids=[1],
            date_start="2026-03-01",
            date_end="2026-03-01",
            hours_preset="peak",
        )
        self.assertEqual(req.resolve_hours(), [7, 8, 12, 13, 17, 18])

    def test_job_request_rejects_hours_and_preset_together(self):
        with self.assertRaises(Exception):
            JobRequest(
                node_ids=[1],
                date_start="2026-03-01",
                date_end="2026-03-01",
                hours=[8],
                hours_preset="all",
            )

    def test_raw_vknd_request_defaults_to_hourly(self):
        req = RawTrafficVkndRequest(
            node_ids=[1],
            date_start="2026-03-01",
            date_end="2026-03-01",
        )
        self.assertEqual(req.interval, "1h")

    def test_raw_drct_request_supports_filters_and_daily_interval(self):
        req = RawTrafficDrctRequest(
            node_ids=[1],
            date_start="2026-03-01",
            date_end="2026-03-01",
            interval="1d",
            approach_ids=[10],
            drct_codes=["01"],
        )
        self.assertEqual(req.interval, "1d")
        self.assertEqual(req.approach_ids, [10])
        self.assertEqual(req.drct_codes, ["01"])


class TestDomainAggregation(unittest.TestCase):
    def test_drct_aggregation_is_pure_and_ignores_00(self):
        rows = aggregate_acsr_slots_from_drct([
            {
                "date": "2026-03-22",
                "hour": 8,
                "node_id": 1,
                "node_name": "N",
                "approach_id": 10,
                "approach_name": "A",
                "drct_cd": "00",
                "traffic_volume": 999,
                "anomaly_type": "B형",
                "corrected_value": 999,
            },
            {
                "date": "2026-03-22",
                "hour": 8,
                "node_id": 1,
                "node_name": "N",
                "approach_id": 10,
                "approach_name": "A",
                "drct_cd": "01",
                "traffic_volume": 0,
                "anomaly_type": "A형",
                "corrected_value": 100,
            },
            {
                "date": "2026-03-22",
                "hour": 8,
                "node_id": 1,
                "node_name": "N",
                "approach_id": 10,
                "approach_name": "A",
                "drct_cd": "02",
                "traffic_volume": 20,
                "anomaly_type": "B형",
                "corrected_value": 30,
            },
        ])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["traffic_volume"], 20)
        self.assertEqual(rows[0]["corrected_value"], 130)
        self.assertEqual(rows[0]["anomaly_type"], "A+B혼합")


class TestJobUseCase(unittest.TestCase):
    def test_create_runs_background_job_and_serializes_status(self):
        store = InMemoryJobStore()
        gateway = _Gateway()
        usecase = JobUseCase(store, _ImmediateRunner(), gateway)
        created = usecase.create(JobRequest(
            node_ids=[1],
            date_start="2026-03-01",
            date_end="2026-03-01",
            hours=[8],
        ))

        job = usecase.get(created["job_id"])
        self.assertEqual(job["status"], "done")
        self.assertEqual(job["result"], {"slots": []})
        self.assertEqual(gateway.jobs, [created["job_id"]])


class TestFastApiEntryPoint(unittest.TestCase):
    def test_public_endpoints_are_registered(self):
        client = TestClient(create_app())
        paths = client.get("/openapi.json").json()["paths"]
        for path in (
            "/jobs",
            "/jobs/{job_id}",
            "/corrected-traffic",
            "/corrected-traffic-drct",
            "/raw-traffic",
            "/raw-traffic-drct",
            "/raw-traffic-vknd",
            "/corrected-traffic-vknd",
            "/vehicle-kinds",
            "/anomaly-daily-summary",
            "/intersections",
        ):
            self.assertIn(path, paths)

    def test_corrected_traffic_response_shape_comes_from_use_case(self):
        app = create_app()
        app.state.services["traffic"].corrected_traffic = MagicMock(
            return_value={"slots": [{"hour": 8}]}
        )
        client = TestClient(app)

        resp = client.post(
            "/corrected-traffic",
            json={
                "node_ids": [1],
                "date_start": "2026-03-01",
                "date_end": "2026-03-01",
                "hours_preset": "peak",
            },
        )

        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.json(), {"slots": [{"hour": 8}]})
        call_req = app.state.services["traffic"].corrected_traffic.call_args.args[0]
        self.assertEqual(call_req.resolve_hours(), [7, 8, 12, 13, 17, 18])


class TestLegacyGateway(unittest.TestCase):
    def test_run_job_closes_connection_on_failure(self):
        gateway = LegacyTrafficAnalysisGateway()
        job = {
            "job_id": "j1",
            "status": "pending",
            "request": {
                "node_ids": [1001],
                "date_start": "2026-03-01",
                "date_end": "2026-03-01",
                "hours": [8],
            },
            "created_at": None,
            "started_at": None,
            "finished_at": None,
            "progress": None,
            "result": None,
            "error": None,
        }
        mock_conn = MagicMock()

        with patch("traffic_api.infrastructure.legacy_analysis_gateway.legacy._holiday_dates", set()), \
             patch("traffic_api.infrastructure.legacy_analysis_gateway.legacy.ad.connect_db", return_value=mock_conn), \
             patch("traffic_api.infrastructure.legacy_analysis_gateway.legacy.ad.select_baseline_years", side_effect=RuntimeError("boom")):
            gateway.run_job(job)

        self.assertEqual(job["status"], "failed")
        self.assertIn("boom", job["error"])
        mock_conn.close.assert_called_once()


if __name__ == "__main__":
    unittest.main(verbosity=2)
