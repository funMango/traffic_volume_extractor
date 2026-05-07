#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VKND CLI parsing, selection, and visualization helper tests."""

import os
import sys
import types
from collections import defaultdict
from datetime import date
from pathlib import Path
from unittest.mock import patch


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
            "__init__": lambda s, **kw: setattr(s, "traces", []),
            "add_trace": lambda s, *a, **kw: s.traces.append((a, kw)),
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
        "create_sheet": lambda s, title=None: None,
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


import 이상탐지 as ad  # noqa: E402


def test_parse_vknd_response_preserves_minute_and_mixed_anomaly():
    slots = [
        {
            "interval": "15m",
            "date": "2026-03-22",
            "hour": 8,
            "minute": 15,
            "approach_id": 2001,
            "approach_name": "North",
            "vknd_cd": "1",
            "vknd_name": "Car",
            "traffic_volume": 0,
            "anomaly_type": "A+B혼합",
            "corrected_value": 120,
            "correction_method": "median",
            "confidence": "LOW",
        }
    ]

    rows, target_data, approaches, approaches_by_vknd = ad._parse_api_response_vknd(slots, "Node")

    assert target_data[(2001, "1", date(2026, 3, 22), 8, 15)] == 0
    assert rows[0]["시간"] == "08:15"
    assert rows[0]["차종"] == "Car"
    assert rows[0]["판정"] == "A+B혼합"
    assert rows[0]["_minute"] == 15
    assert approaches == [(2001, "North")]
    assert approaches_by_vknd[2001]["vknds"] == [("1", "Car")]


def test_vknd_time_format_uses_hour_for_1h_and_minute_for_15m():
    assert ad._format_vknd_time(8, None, "1h") == "08:00"
    assert ad._format_vknd_time(8, 45, "15m") == "08:45"


def test_input_vknd_codes_accepts_enter_all_number_and_code():
    items = [{"code": "1", "name": "Car"}, {"code": "2", "name": "Bus"}]
    with patch("builtins.input", return_value=""):
        assert ad.input_vknd_codes(items) == ["1", "2"]
    with patch("builtins.input", return_value="1,2"):
        assert ad.input_vknd_codes(items) == ["1", "2"]
    with patch("builtins.input", return_value="2"):
        assert ad.input_vknd_codes(items) == ["2"]


def test_vknd_exporter_builds_approach_and_vehicle_files(tmp_path):
    written = []

    def fake_write(fig, out_path, verbose=True):
        written.append(out_path)

    approaches_by_vknd = {
        2001: {
            "approach_name": "North",
            "vknds": [("1", "Car"), ("2", "Bus")],
        }
    }
    target_data = {
        (2001, "1", date(2026, 3, 22), 8, 0): 10,
        (2001, "2", date(2026, 3, 22), 8, 0): 20,
    }

    with patch.object(ad, "_write_plotly_html", side_effect=fake_write):
        ad.export_vknd_visualizations(
            node_name="Node",
            approaches_by_vknd=approaches_by_vknd,
            target_data=target_data,
            node_results=[],
            date_start=date(2026, 3, 22),
            date_end=date(2026, 3, 22),
            hours=[8],
            interval="1h",
            viz_dir=tmp_path,
            verbose=False,
        )

    rel_paths = {str(path.relative_to(tmp_path)) for path in written}
    assert os.path.join("Node_260322~260322", "North", "Car.html") in rel_paths
    assert os.path.join("Node_260322~260322", "North", "Bus.html") in rel_paths
    assert os.path.join("Node_260322~260322", "North", "_전체.html") in rel_paths
