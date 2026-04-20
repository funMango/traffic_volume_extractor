#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Core anomaly calculation logic.

This module intentionally contains pure calculation logic and related constants.
It does not perform DB I/O, API calls, or CLI interactions.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta

import numpy as np

AVAILABLE_YEARS = [2022, 2024, 2025, 2026]
DATA_START_YEAR = 2023
ZERO_RATE_THRESHOLD = 0.2
RATIO_THRESHOLD = 0.5
IQR_MULTIPLIER = 1.5
MEDIAN_RATIO = 0.3
MIN_CLEAN_SAMPLES = 5
MIN_CONFIDENCE_SAMPLES = 8
MAX_INTERP_DAYS = 14
MAX_INTERP_DAYS_OK = 7
ZERO_RATE_EXPANSION_THRESHOLD = 0.8


def get_day_type(d: date, holiday_dates: set) -> int:
    """0=weekday, 1=saturday, 2=sunday/holiday."""
    if d.strftime("%Y-%m-%d") in holiday_dates:
        return 2
    wd = d.weekday()
    if wd == 6:
        return 2
    if wd == 5:
        return 1
    return 0


def select_baseline_years(target_year: int) -> list:
    """Select up to two nearest baseline years excluding target year."""
    candidates = [y for y in AVAILABLE_YEARS if y != target_year]
    before = sorted([y for y in candidates if y < target_year], reverse=True)
    after = sorted([y for y in candidates if y > target_year])
    result = []
    if before:
        result.append(before[0])
    if after:
        result.append(after[0])
    if len(result) < 2:
        if not before and len(after) >= 2:
            result.append(after[1])
        if not after and len(before) >= 2:
            result.append(before[1])
    return sorted(result)


def select_fallback_year(target_year: int) -> int:
    """Select a fallback year when baseline data is insufficient."""
    if target_year > DATA_START_YEAR:
        return target_year - 1
    return target_year + 1


def compute_baseline(values: list) -> dict | None:
    clean = [v for v in values if v is not None and v >= 0]
    for _ in range(3):
        if len(clean) < MIN_CLEAN_SAMPLES:
            break
        arr = np.array(clean, dtype=float)
        median = float(np.median(arr))
        q1, q3 = float(np.percentile(arr, 25)), float(np.percentile(arr, 75))
        iqr = q3 - q1
        lower = max(0.0, q1 - IQR_MULTIPLIER * iqr, median * MEDIAN_RATIO)
        new_clean = [v for v in clean if v >= lower]
        if len(new_clean) == len(clean):
            break
        clean = new_clean

    if len(clean) < MIN_CLEAN_SAMPLES:
        return None
    arr = np.array(clean, dtype=float)
    return {
        "median": float(np.median(arr)),
        "q1": float(np.percentile(arr, 25)),
        "q3": float(np.percentile(arr, 75)),
        "n_clean": len(clean),
    }


def build_baselines(
    raw_rows: list,
    baseline_years: list,
    holiday_dates: set,
    approaches: list,
    date_start: date,
    date_end: date,
    hours: list,
    adj_month_rows: list = None,
    fallback_rows: list = None,
    expand_stages: bool = True,
) -> tuple:
    """Build per-cell baselines for the target period."""
    needed_cells = set()
    cur_d = date_start
    while cur_d <= date_end:
        day_type = get_day_type(cur_d, holiday_dates)
        month = cur_d.month
        for hour in hours:
            for acsr_id, _ in approaches:
                needed_cells.add((acsr_id, day_type, hour, month))
        cur_d += timedelta(days=1)

    cell_values = defaultdict(list)
    cell_actual_dates = defaultdict(set)
    cell_zero_counts: dict[tuple, int] = defaultdict(int)
    all_month_values: dict = defaultdict(list)
    all_month_actual_dates: dict = defaultdict(set)

    for acsr_id, tot_dt, trf_qnty in raw_rows:
        d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
        h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
        month = d.month
        day_type = get_day_type(d, holiday_dates)
        cell = (acsr_id, day_type, h, month)
        if trf_qnty is not None:
            all_month_values[cell].append(trf_qnty)
        all_month_actual_dates[cell].add(d)
        if cell in needed_cells:
            if trf_qnty is not None:
                cell_values[cell].append(trf_qnty)
                if trf_qnty == 0:
                    cell_zero_counts[cell] += 1
            cell_actual_dates[cell].add(d)

    cell_adj_counts: dict[tuple, int] = defaultdict(int)
    if adj_month_rows:
        target_months = sorted({month for _, _, _, month in needed_cells})
        needed_keys = {(a, dt, h) for a, dt, h, _ in needed_cells}
        for acsr_id, tot_dt, trf_qnty in adj_month_rows:
            if trf_qnty is None:
                continue
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            day_type = get_day_type(d, holiday_dates)
            if (acsr_id, day_type, h) not in needed_keys:
                continue
            adj_month = d.month
            nearest_month = min(target_months, key=lambda m: abs(m - adj_month))
            cell = (acsr_id, day_type, h, nearest_month)
            if cell in needed_cells:
                cell_values[cell].append(trf_qnty)
                cell_adj_counts[cell] += 1

    cell_fallback_counts: dict[tuple, int] = defaultdict(int)
    if fallback_rows:
        fallback_cells = {
            cell for cell in needed_cells if len(cell_values.get(cell, [])) < MIN_CLEAN_SAMPLES
        }
        if fallback_cells:
            fallback_by_key: dict[tuple, list] = defaultdict(list)
            for acsr_id, tot_dt, trf_qnty in fallback_rows:
                if trf_qnty is None:
                    continue
                d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
                h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
                day_type = get_day_type(d, holiday_dates)
                fallback_by_key[(acsr_id, day_type, h)].append((d, trf_qnty))

            for cell in fallback_cells:
                acsr_id, day_type, hour, _month = cell
                already_seen = cell_actual_dates.get(cell, set())
                for d, val in fallback_by_key.get((acsr_id, day_type, hour), []):
                    if d not in already_seen:
                        cell_values[cell].append(val)
                        cell_fallback_counts[cell] += 1

    expected_counts: dict[tuple, int] = defaultdict(int)
    today = date.today()
    for year in baseline_years:
        cur_d = date(year, 1, 1)
        end_y = date(year, 12, 31)
        while cur_d <= end_y:
            if cur_d >= today:
                cur_d += timedelta(days=1)
                continue
            day_type = get_day_type(cur_d, holiday_dates)
            month = cur_d.month
            for h in range(24):
                expected_counts[(day_type, h, month)] += 1
            cur_d += timedelta(days=1)

    baselines = {}
    for cell in needed_cells:
        acsr_id, day_type, hour, month = cell
        exp_count = expected_counts.get((day_type, hour, month), 0)
        if exp_count == 0:
            continue
        act_count = len(cell_actual_dates.get(cell, set()))
        zero_rate = (exp_count - act_count) / exp_count
        stats = compute_baseline(cell_values.get(cell, []))
        n_adj = cell_adj_counts.get(cell, 0)
        n_fallback = cell_fallback_counts.get(cell, 0)
        baselines[cell] = {
            "zero_rate": zero_rate,
            "stage0_zr": zero_rate,
            "stage1_zr": None,
            "stage2_zr": None,
            "stats": stats,
            "n_adj": n_adj,
            "n_fallback": n_fallback,
            "expansion_stage": 0,
            "exp_count": exp_count,
            "act_count": act_count,
            "n_zero_baseline": cell_zero_counts.get(cell, 0),
            "n_total_baseline": len(cell_values.get(cell, [])),
        }

    if expand_stages:
        for cell, bl in baselines.items():
            if bl["zero_rate"] < ZERO_RATE_EXPANSION_THRESHOLD:
                continue

            acsr_id, day_type, hour, month = cell
            prev_m = ((month - 2) % 12) + 1
            next_m = (month % 12) + 1
            stage1_months = [prev_m, month, next_m]

            exp_1 = sum(expected_counts.get((day_type, hour, m), 0) for m in stage1_months)
            if exp_1 > 0:
                act_dates_1: set = set()
                for m in stage1_months:
                    act_dates_1 |= all_month_actual_dates.get((acsr_id, day_type, hour, m), set())
                zr_1 = (exp_1 - len(act_dates_1)) / exp_1
                vals_1 = []
                for m in stage1_months:
                    vals_1.extend(all_month_values.get((acsr_id, day_type, hour, m), []))
                bl["zero_rate"] = zr_1
                bl["stage1_zr"] = zr_1
                bl["stats"] = compute_baseline(vals_1)
                bl["expansion_stage"] = 1
                if zr_1 < ZERO_RATE_EXPANSION_THRESHOLD and bl["stats"] is not None:
                    continue

            exp_2 = sum(expected_counts.get((day_type, hour, m), 0) for m in range(1, 13))
            if exp_2 > 0:
                act_dates_2: set = set()
                for m in range(1, 13):
                    act_dates_2 |= all_month_actual_dates.get((acsr_id, day_type, hour, m), set())
                zr_2 = (exp_2 - len(act_dates_2)) / exp_2
                vals_2 = []
                for m in range(1, 13):
                    vals_2.extend(all_month_values.get((acsr_id, day_type, hour, m), []))
                bl["zero_rate"] = zr_2
                bl["stage2_zr"] = zr_2
                bl["stats"] = compute_baseline(vals_2)
                bl["expansion_stage"] = 2

    fallback_applied = {c for c, v in cell_fallback_counts.items() if v > 0}
    n_stage1 = sum(1 for bl in baselines.values() if bl["expansion_stage"] >= 1)
    n_stage2 = sum(1 for bl in baselines.values() if bl["expansion_stage"] == 2)
    fallback_summary = {
        "n_cells": len(fallback_applied),
        "n_values": sum(cell_fallback_counts.values()),
        "n_stage1": n_stage1,
        "n_stage2": n_stage2,
    }
    return baselines, fallback_summary


def detect_anomalies(
    node_name: str,
    approaches: list,
    baselines: dict,
    target_data: dict,
    date_start: date,
    date_end: date,
    hours: list,
    holiday_dates: set,
) -> list:
    results = []
    cur_d = date_start
    while cur_d <= date_end:
        day_type = get_day_type(cur_d, holiday_dates)
        month = cur_d.month
        for hour in hours:
            for acsr_id, acsr_nm in approaches:
                cell = (acsr_id, day_type, hour, month)
                bl = baselines.get(cell)
                if bl is None:
                    continue

                zero_rate = bl["zero_rate"]
                stats = bl["stats"]
                has_record = (acsr_id, cur_d, hour) in target_data
                trf_qnty = target_data.get((acsr_id, cur_d, hour))

                anomaly_type = None
                trf_val = 0

                if not has_record:
                    if zero_rate < ZERO_RATE_THRESHOLD:
                        anomaly_type = "A형"
                        trf_val = 0
                    elif zero_rate >= ZERO_RATE_EXPANSION_THRESHOLD:
                        anomaly_type = "정상"
                        trf_val = 0
                    else:
                        anomaly_type = "B형"
                        trf_val = 0
                elif stats is not None:
                    median = stats["median"]
                    q1 = stats["q1"]
                    q3 = stats["q3"]
                    iqr = q3 - q1
                    lower_iqr = q1 - IQR_MULTIPLIER * iqr
                    ratio = (trf_qnty / median) if median > 0 else 1.0
                    if ratio < RATIO_THRESHOLD or trf_qnty < lower_iqr:
                        anomaly_type = "B형"
                        trf_val = trf_qnty

                if anomaly_type:
                    results.append(
                        {
                            "날짜": cur_d.strftime("%Y.%m.%d"),
                            "시간": f"{hour:02d}:00",
                            "교차로": node_name,
                            "방향": acsr_nm,
                            "교통량": trf_val,
                            "판정": anomaly_type,
                            "_acsr_id": acsr_id,
                            "_date": cur_d,
                            "_hour": hour,
                            "_cell": cell,
                        }
                    )
        cur_d += timedelta(days=1)
    return results


def apply_corrections(results: list, baselines: dict, target_data: dict) -> None:
    """Apply correction values in-place."""
    anomaly_set = {(r["_acsr_id"], r["_date"], r["_hour"]) for r in results}

    for r in results:
        if r["판정"] == "정상":
            r["보정값"] = None
            r["보정방법"] = None
            r["신뢰도"] = None
            continue

        acsr_id = r["_acsr_id"]
        cur_d = r["_date"]
        hour = r["_hour"]
        cell = r["_cell"]

        bl = baselines.get(cell)
        stats = bl["stats"] if bl else None

        prev_val, d_before = None, None
        for d in range(1, MAX_INTERP_DAYS + 1):
            key = (acsr_id, cur_d - timedelta(days=d), hour)
            if key in anomaly_set:
                continue
            if key in target_data:
                prev_val = target_data[key]
                d_before = d
                break

        next_val, d_after = None, None
        for d in range(1, MAX_INTERP_DAYS + 1):
            key = (acsr_id, cur_d + timedelta(days=d), hour)
            if key in anomaly_set:
                continue
            if key in target_data:
                next_val = target_data[key]
                d_after = d
                break

        corrected, method = None, None
        interp_dist = None
        if prev_val is not None and next_val is not None and d_before + d_after <= MAX_INTERP_DAYS:
            corrected = prev_val + (next_val - prev_val) * d_before / (d_before + d_after)
            method = "선형보간"
            interp_dist = d_before + d_after
        elif stats is not None:
            corrected = stats["median"]
            method = "median"

        if corrected is not None:
            r["보정값"] = round(corrected)
            r["보정방법"] = method
            is_low = (
                (stats and stats["n_clean"] < MIN_CONFIDENCE_SAMPLES)
                or bl.get("n_fallback", 0) > 0
                or (interp_dist is not None and interp_dist > MAX_INTERP_DAYS_OK)
            )
            r["신뢰도"] = "LOW" if is_low else "OK"
        else:
            r["보정값"] = None
            r["보정방법"] = None
            r["신뢰도"] = None
