#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(__file__))

import anomaly_core as core


def _dt(y: int, m: int, d: int, h: int = 8) -> datetime:
    return datetime(y, m, d, h)


def test_select_baseline_years_excludes_target():
    years = core.select_baseline_years(2025)
    assert 2025 not in years
    assert len(years) <= 2


def test_select_fallback_year_rule():
    assert core.select_fallback_year(2026) == 2025
    assert core.select_fallback_year(core.DATA_START_YEAR) == core.DATA_START_YEAR + 1


def test_build_baselines_stage_expansion_when_no_rows():
    approaches = [(1001, "북향")]
    baselines, summary = core.build_baselines(
        raw_rows=[],
        baseline_years=[2024],
        holiday_dates=set(),
        approaches=approaches,
        date_start=date(2026, 3, 15),
        date_end=date(2026, 3, 15),
        hours=[8],
    )
    cell = (1001, core.get_day_type(date(2026, 3, 15), set()), 8, 3)
    assert cell in baselines
    assert baselines[cell]["expansion_stage"] == 2
    assert summary["n_stage2"] >= 1


def test_detect_anomalies_missing_record_classification():
    d = date(2026, 3, 15)
    cell = (1001, core.get_day_type(d, set()), 8, 3)
    approaches = [(1001, "북향")]

    baselines_a = {cell: {"zero_rate": 0.1, "stats": {"median": 100, "q1": 80, "q3": 120}}}
    res_a = core.detect_anomalies("테스트", approaches, baselines_a, {}, d, d, [8], set())
    assert res_a[0]["판정"] == "A형"

    baselines_b = {cell: {"zero_rate": 0.4, "stats": {"median": 100, "q1": 80, "q3": 120}}}
    res_b = core.detect_anomalies("테스트", approaches, baselines_b, {}, d, d, [8], set())
    assert res_b[0]["판정"] == "B형"

    baselines_n = {cell: {"zero_rate": 0.9, "stats": {"median": 100, "q1": 80, "q3": 120}}}
    res_n = core.detect_anomalies("테스트", approaches, baselines_n, {}, d, d, [8], set())
    assert res_n[0]["판정"] == "정상"


def test_apply_corrections_prefers_linear_interpolation():
    d = date(2026, 3, 15)
    result = [
        {
            "판정": "A형",
            "_acsr_id": 1001,
            "_date": d,
            "_hour": 8,
            "_cell": (1001, core.get_day_type(d, set()), 8, 3),
        }
    ]
    baselines = {result[0]["_cell"]: {"stats": {"median": 200, "n_clean": 10}, "n_fallback": 0}}
    target_data = {
        (1001, date(2026, 3, 14), 8): 100,
        (1001, date(2026, 3, 16), 8): 300,
    }

    core.apply_corrections(result, baselines, target_data)
    assert result[0]["보정값"] == 200
    assert result[0]["보정방법"] == "선형보간"
    assert result[0]["신뢰도"] in {"OK", "LOW"}


def test_apply_corrections_falls_back_to_median():
    d = date(2026, 3, 15)
    cell = (1001, core.get_day_type(d, set()), 8, 3)
    result = [{"판정": "B형", "_acsr_id": 1001, "_date": d, "_hour": 8, "_cell": cell}]
    baselines = {cell: {"stats": {"median": 222, "n_clean": 10}, "n_fallback": 0}}

    core.apply_corrections(result, baselines, target_data={})
    assert result[0]["보정값"] == 222
    assert result[0]["보정방법"] == "median"
