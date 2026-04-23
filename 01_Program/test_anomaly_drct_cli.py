#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""이상탐지.py DRCT 선택형 CLI 확장 테스트."""

import os
import sys
import types
from datetime import date
from collections import defaultdict
from unittest.mock import patch


def _make_module(name: str) -> types.ModuleType:
    mod = types.ModuleType(name)
    sys.modules[name] = mod
    return mod


sys.path.insert(0, os.path.dirname(__file__))

if "oracledb" not in sys.modules:
    _make_module("oracledb")

if "dotenv" not in sys.modules:
    _dotenv = _make_module("dotenv")
    _dotenv.load_dotenv = lambda *a, **kw: None  # type: ignore

for _pmod in ("plotly", "plotly.graph_objects", "plotly.subplots"):
    if _pmod not in sys.modules:
        _m = _make_module(_pmod)
        if _pmod == "plotly.subplots":
            _m.make_subplots = lambda *a, **kw: None
        else:
            _m.Figure = type("Figure", (), {})
            _m.Scatter = type("Scatter", (), {})

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


import 이상탐지 as ad  # noqa: E402


def test_parse_api_response_drct_splits_anomaly_and_target_data():
    drct_slots = [
        {
            "date": "2026-03-22",
            "hour": 7,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "00",
            "drct_name": "미분류",
            "traffic_volume": 999,
            "anomaly_type": "B형",
            "corrected_value": 999,
            "correction_method": "median",
            "confidence": "LOW",
        },
        {
            "date": "2026-03-22",
            "hour": 7,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "01",
            "drct_name": "좌회전",
            "traffic_volume": 80,
            "anomaly_type": None,
            "corrected_value": None,
            "correction_method": None,
            "confidence": None,
        },
        {
            "date": "2026-03-22",
            "hour": 8,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "01",
            "drct_name": "좌회전",
            "traffic_volume": 0,
            "anomaly_type": "A형",
            "corrected_value": 120,
            "correction_method": "median",
            "confidence": "OK",
        },
        {
            "date": "2026-03-22",
            "hour": 8,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "02",
            "drct_name": "직진",
            "traffic_volume": 30,
            "anomaly_type": "B형",
            "corrected_value": 95,
            "correction_method": "선형보간",
            "confidence": "LOW",
        },
    ]

    node_results, target_data, approaches, drct_options = ad._parse_api_response_drct(
        drct_slots, "송내사거리"
    )

    assert len(node_results) == 2
    assert len(target_data) == 3
    assert approaches == [(1001, "북(남)")]
    assert drct_options[1001]["drcts"] == [("01", "좌회전"), ("02", "직진")]
    assert {r["_drct_cd"] for r in node_results} == {"01", "02"}
    assert all(k[1] != "00" for k in target_data.keys())


def test_parse_api_response_acsr_from_drct_aggregates_mixed_and_excludes_00():
    drct_slots = [
        {
            "date": "2026-03-22",
            "hour": 8,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "00",
            "drct_name": "미분류",
            "traffic_volume": 999,
            "anomaly_type": "B형",
            "corrected_value": 999,
        },
        {
            "date": "2026-03-22",
            "hour": 8,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "01",
            "drct_name": "좌회전",
            "traffic_volume": 0,
            "anomaly_type": "A형",
            "corrected_value": 120,
        },
        {
            "date": "2026-03-22",
            "hour": 8,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "02",
            "drct_name": "직진",
            "traffic_volume": 30,
            "anomaly_type": "B형",
            "corrected_value": 95,
        },
        {
            "date": "2026-03-22",
            "hour": 9,
            "approach_id": 1001,
            "approach_name": "북(남)",
            "drct_cd": "01",
            "drct_name": "좌회전",
            "traffic_volume": 50,
            "anomaly_type": None,
            "corrected_value": None,
        },
    ]

    node_results, target_data, approaches = ad._parse_api_response_acsr_from_drct(
        drct_slots, "송내사거리"
    )

    assert approaches == [(1001, "북(남)")]
    assert len(node_results) == 1
    row = node_results[0]
    assert row["판정"] == "A+B혼합"
    assert row["교통량"] == 30
    assert row["보정값"] == 215
    assert target_data[(1001, date(2026, 3, 22), 8)] == 30
    assert target_data[(1001, date(2026, 3, 22), 9)] == 50


def test_input_drct_selection_enter_means_select_all():
    drct_options = {
        1001: {
            "approach_name": "북(남)",
            "drcts": [("01", "좌회전"), ("02", "직진")],
        }
    }
    with patch("builtins.input", side_effect=[""]):
        selected = ad.input_drct_selection_by_approach(drct_options)
    assert selected == {1001: {"01", "02"}}


def test_filter_by_drct_selection_keeps_only_selected_codes():
    node_results = [
        {"_acsr_id": 1001, "_drct_cd": "01", "판정": "A형"},
        {"_acsr_id": 1001, "_drct_cd": "02", "판정": "B형"},
        {"_acsr_id": 1002, "_drct_cd": "01", "판정": "A형"},
    ]
    target_data = {
        (1001, "01", date(2026, 3, 22), 8): 10,
        (1001, "02", date(2026, 3, 22), 8): 20,
        (1002, "01", date(2026, 3, 22), 8): 30,
    }
    selected = {1001: {"02"}, 1002: {"01"}}

    filtered_results, filtered_target_data = ad.filter_by_drct_selection(
        node_results, target_data, selected
    )

    assert len(filtered_results) == 2
    assert {(r["_acsr_id"], r["_drct_cd"]) for r in filtered_results} == {(1001, "02"), (1002, "01")}
    assert set(filtered_target_data.keys()) == {
        (1001, "02", date(2026, 3, 22), 8),
        (1002, "01", date(2026, 3, 22), 8),
    }


def test_build_excel_sheet_groups_for_multi_intersections():
    rows = [
        {"교차로": "A교차로", "방향": "북(남)", "판정": "A형"},
        {"교차로": "B교차로", "방향": "동(서)", "판정": "B형"},
        {"교차로": "A교차로", "방향": "남(북)", "판정": "정상"},
    ]
    grouped = ad._build_excel_sheet_groups(rows, intersections_count=2, has_drct=False)
    assert list(grouped.keys()) == ["A교차로", "B교차로"]


def test_build_excel_sheet_groups_for_single_intersection():
    rows = [
        {"교차로": "A교차로", "방향": "북(남)", "판정": "A형"},
        {"교차로": "A교차로", "방향": "동(서)", "판정": "B형"},
    ]
    grouped = ad._build_excel_sheet_groups(rows, intersections_count=1, has_drct=False)
    assert list(grouped.keys()) == ["북(남)", "동(서)"]


def test_build_excel_sheet_groups_for_single_intersection_drct():
    rows = [
        {"교차로": "A교차로", "방향": "북(남)", "판정": "A형", "접근로방향": "좌"},
        {"교차로": "A교차로", "방향": "동(서)", "판정": "B형", "접근로방향": "직"},
    ]
    grouped = ad._build_excel_sheet_groups(rows, intersections_count=1, has_drct=True)
    assert list(grouped.keys()) == ["A교차로-북(남)", "A교차로-동(서)"]


def test_build_excel_sheet_groups_includes_mixed_type():
    rows = [
        {"교차로": "A교차로", "방향": "북(남)", "판정": "A+B혼합"},
        {"교차로": "A교차로", "방향": "동(서)", "판정": "정상"},
    ]
    grouped = ad._build_excel_sheet_groups(rows, intersections_count=1, has_drct=False)
    assert list(grouped.keys()) == ["북(남)"]


class _DummyFigure:
    def __init__(self):
        self.traces = []
        self.layout = {}

    def add_trace(self, trace):
        self.traces.append(trace)

    def update_layout(self, **kwargs):
        self.layout.update(kwargs)

    def to_html(self, **kwargs):
        return "<div>dummy-chart</div>"


def _dummy_scatter(**kwargs):
    return kwargs


def test_export_drct_visualizations_creates_individual_and_combined_files(tmp_path):
    drct_options = {
        1001: {
            "approach_name": "동(서)",
            "drcts": [("01", "좌회전"), ("02", "직진")],
        }
    }
    selected = {1001: {"01", "02"}}
    target_data = {
        (1001, "01", date(2026, 3, 22), 8): 10,
        (1001, "02", date(2026, 3, 22), 8): 20,
    }
    node_results = [
        {
            "_acsr_id": 1001,
            "_drct_cd": "01",
            "_date": date(2026, 3, 22),
            "_hour": 8,
            "판정": "A형",
            "보정값": 15,
        }
    ]

    with patch.object(ad.go, "Figure", _DummyFigure), patch.object(ad.go, "Scatter", _dummy_scatter):
        ad.export_drct_visualizations(
            node_name="송내사거리",
            drct_options_by_acsr=drct_options,
            drct_selection_by_acsr=selected,
            target_data=target_data,
            node_results=node_results,
            date_start=date(2026, 3, 22),
            date_end=date(2026, 3, 22),
            hours=[8],
            viz_dir=tmp_path,
        )

    node_dir = tmp_path / "송내사거리_260322~260322"
    approach_dir = node_dir / "동(서)"
    assert (approach_dir / "좌회전.html").exists()
    assert (approach_dir / "직진.html").exists()
    assert (approach_dir / "_전체.html").exists()


def test_export_acsr_visualizations_marks_mixed_with_a_and_b_markers(tmp_path):
    captured_figs = []

    def _capture_html(fig, out_path, verbose=True):
        captured_figs.append(fig)

    target_data = {
        (1001, date(2026, 3, 22), 8): 30,
    }
    node_results = [
        {
            "_acsr_id": 1001,
            "_date": date(2026, 3, 22),
            "_hour": 8,
            "교통량": 30,
            "판정": "A+B혼합",
            "보정값": 40,
        }
    ]

    with patch.object(ad.go, "Figure", _DummyFigure), \
         patch.object(ad.go, "Scatter", _dummy_scatter), \
         patch.object(ad, "_write_plotly_html", side_effect=_capture_html), \
         patch.object(ad, "export_visualization", return_value=None):
        ad.export_acsr_visualizations(
            node_name="송내사거리",
            approaches=[(1001, "북(남)")],
            target_data=target_data,
            node_results=node_results,
            date_start=date(2026, 3, 22),
            date_end=date(2026, 3, 22),
            hours=[8],
            viz_dir=tmp_path,
            verbose=False,
        )

    assert captured_figs
    names = [trace.get("name") for trace in captured_figs[0].traces]
    assert "보정 전" in names
    assert "보정 후" in names
    assert "A형(결측)" in names
    assert "B형(이상저값)" in names


def test_input_view_mode_returns_expected_value():
    with patch("builtins.input", side_effect=["2"]):
        assert ad.input_view_mode() == "DRCT"
