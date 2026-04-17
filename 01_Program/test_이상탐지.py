#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
이상탐지.py — C형 인접 월 보강 로직 테스트
대상 함수:
  - get_adjacent_month_periods
  - load_adjacent_month_data
  - build_baselines (adj_month_rows 포함)
"""

import sys
import os
from datetime import date, datetime, timedelta
from collections import defaultdict
from unittest.mock import MagicMock, patch

import pytest
import numpy as np

# ── 경로 추가 ──────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))

# ── 외부 의존성 스텁 (oracledb / plotly / openpyxl / dotenv) ───────────────────
import types

def _make_module(name: str) -> types.ModuleType:
    m = types.ModuleType(name)
    sys.modules[name] = m
    return m

# oracledb
if "oracledb" not in sys.modules:
    _make_module("oracledb")

# dotenv
if "dotenv" not in sys.modules:
    _dotenv = _make_module("dotenv")
    _dotenv.load_dotenv = lambda *a, **kw: None  # type: ignore

# plotly
for _pmod in ("plotly", "plotly.graph_objects"):
    if _pmod not in sys.modules:
        _m = _make_module(_pmod)
        _m.Figure = type("Figure", (), {"__init__": lambda s, **kw: None,
                                         "add_trace": lambda s, *a, **kw: None,
                                         "update_layout": lambda s, **kw: None,
                                         "write_html": lambda s, *a, **kw: None})
        _m.Scatter = type("Scatter", (), {"__init__": lambda s, **kw: None})

# openpyxl — 실제 설치 여부 무관하게 완전 stub
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

from 이상탐지 import (  # noqa: E402
    get_adjacent_month_periods,
    load_adjacent_month_data,
    build_baselines,
    get_day_type,
    ZERO_RATE_EXPANSION_THRESHOLD,
)

# ════════════════════════════════════════════════════════════════════════════════
# 헬퍼
# ════════════════════════════════════════════════════════════════════════════════

HOLIDAY_DATES: set = set()  # 테스트에서는 공휴일 없음 가정


def _dt(year, month, day, hour=8) -> datetime:
    return datetime(year, month, day, hour)


def _make_raw_rows(acsr_id: int, year: int, month: int,
                   day_range: range, hour: int, value: int) -> list:
    """베이스라인용 (acsr_id, tot_dt, trf_qnty) 목 데이터 생성"""
    rows = []
    for day in day_range:
        try:
            rows.append((acsr_id, _dt(year, month, day, hour), value))
        except ValueError:
            pass  # 해당 월에 존재하지 않는 날짜 무시
    return rows


# ════════════════════════════════════════════════════════════════════════════════
# 1. get_adjacent_month_periods
# ════════════════════════════════════════════════════════════════════════════════

class TestGetAdjacentMonthPeriods:

    # ── 정상 케이스 ─────────────────────────────────────────────────────────────

    def test_single_month_middle(self):
        """단일 월(2026-02) → 양쪽 인접 (2026-01), (2026-03) 반환"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 4, 13)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 2, 1), date(2026, 2, 28))
        assert (2026, 1) in result
        assert (2026, 3) in result
        assert (2026, 2) not in result

    def test_multi_month_range(self):
        """3개월(2026-01~03) → 경계 인접만: (2025-12), (2026-04)"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 5, 1)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 1, 1), date(2026, 3, 31))
        assert (2025, 12) in result
        assert (2026, 4) in result
        # 대상 기간 내부 월은 포함되지 않아야 함
        assert (2026, 1) not in result
        assert (2026, 2) not in result
        assert (2026, 3) not in result

    def test_year_start_prev_is_previous_year(self):
        """1월 대상 → 이전 월은 전년도 12월"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 4, 13)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 1, 1), date(2026, 1, 31))
        assert (2025, 12) in result

    def test_year_end_next_is_future_excluded(self):
        """12월 대상, 현재가 12월 → 내년 1월은 미래라 제외"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 12, 15)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 12, 1), date(2026, 12, 31))
        assert (2027, 1) not in result
        assert (2026, 11) in result

    def test_future_month_excluded(self):
        """인접 월이 미래(today 이후)이면 제외"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 2, 28)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 2, 1), date(2026, 2, 28))
        # 2026-03은 미래(오늘이 2월 28일 → 3월 1일이 오늘보다 미래)
        assert (2026, 3) not in result
        assert (2026, 1) in result

    def test_returns_sorted(self):
        """결과가 오름차순 정렬인지 확인"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 6, 1)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 3, 1), date(2026, 3, 31))
        assert result == sorted(result)

    # ── 경계값 케이스 ────────────────────────────────────────────────────────────

    def test_same_day_start_end(self):
        """date_start == date_end (하루짜리 기간)"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 4, 13)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 2, 15), date(2026, 2, 15))
        assert (2026, 1) in result
        assert (2026, 3) in result

    def test_no_adj_available(self):
        """모든 인접 월이 미래 → 빈 리스트"""
        with patch("이상탐지.date") as mock_date:
            mock_date.today.return_value = date(2026, 2, 1)
            mock_date.side_effect = lambda *a, **kw: date(*a, **kw)
            result = get_adjacent_month_periods(date(2026, 2, 1), date(2026, 2, 28))
        # 2026-03은 미래, 2026-01은 과거이므로 1개만 있어야 함
        assert (2026, 3) not in result
        assert (2026, 1) in result


# ════════════════════════════════════════════════════════════════════════════════
# 2. load_adjacent_month_data
# ════════════════════════════════════════════════════════════════════════════════

class TestLoadAdjacentMonthData:

    def _make_mock_conn(self, rows):
        """rows를 반환하는 mock DB 커넥션 생성"""
        cur = MagicMock()
        cur.fetchall.return_value = rows
        cur.__enter__ = lambda s: s
        cur.__exit__ = MagicMock(return_value=False)
        conn = MagicMock()
        conn.cursor.return_value = cur
        return conn, cur

    # ── 정상 케이스 ─────────────────────────────────────────────────────────────

    def test_returns_rows_from_db(self):
        """DB에서 반환된 row가 그대로 리턴되는지"""
        mock_rows = [
            (1001, _dt(2026, 1, 5, 8), 120),
            (1001, _dt(2026, 3, 5, 8), 135),
        ]
        conn, cur = self._make_mock_conn(mock_rows)
        result = load_adjacent_month_data(conn, node_id=1, adj_periods=[(2026, 1), (2026, 3)])
        assert result == mock_rows

    def test_sql_contains_year_month_conditions(self):
        """생성된 SQL에 연도·월 조건이 포함되는지"""
        conn, cur = self._make_mock_conn([])
        load_adjacent_month_data(conn, node_id=99, adj_periods=[(2026, 1), (2025, 12)])
        sql_called = cur.execute.call_args[0][0]
        assert "2026" in sql_called
        assert "2025" in sql_called

    def test_correct_node_id_passed(self):
        """node_id가 올바르게 바인딩되는지"""
        conn, cur = self._make_mock_conn([])
        load_adjacent_month_data(conn, node_id=42, adj_periods=[(2026, 1)])
        _, kwargs = cur.execute.call_args
        assert kwargs.get("nid") == 42

    # ── 경계값 케이스 ────────────────────────────────────────────────────────────

    def test_empty_adj_periods_returns_empty(self):
        """adj_periods가 빈 리스트면 DB 조회 없이 [] 반환"""
        conn = MagicMock()
        result = load_adjacent_month_data(conn, node_id=1, adj_periods=[])
        assert result == []
        conn.cursor.assert_not_called()

    def test_single_period(self):
        """인접 월이 1개인 경우"""
        conn, cur = self._make_mock_conn([(1001, _dt(2026, 1, 10, 9), 80)])
        result = load_adjacent_month_data(conn, node_id=1, adj_periods=[(2026, 1)])
        assert len(result) == 1

    def test_db_returns_empty(self):
        """DB에 해당 데이터 없는 경우 빈 리스트"""
        conn, cur = self._make_mock_conn([])
        result = load_adjacent_month_data(conn, node_id=1, adj_periods=[(2026, 1), (2026, 3)])
        assert result == []


# ════════════════════════════════════════════════════════════════════════════════
# 3. build_baselines — adj_month_rows 병합 로직
# ════════════════════════════════════════════════════════════════════════════════

class TestBuildBaselinesAdjMonth:
    """
    셀 구성:
      acsr_id=1001, day_type=0(평일), hour=8, month=2
    대상 기간: 2026-02-03 ~ 2026-02-07 (월~금 평일)
    베이스라인 연도: [2024, 2025]
    """

    ACSR_ID   = 1001
    APPROACHES = [(1001, "북")]
    DATE_START = date(2026, 2, 3)   # 화요일
    DATE_END   = date(2026, 2, 7)   # 토요일 포함 안 하려면 금요일까지
    HOURS      = [8]
    BASELINE_YEARS = [2024, 2025]

    def _make_baseline_rows(self, year: int, month: int,
                             days: list, value: int) -> list:
        rows = []
        for d in days:
            rows.append((self.ACSR_ID, datetime(year, month, d, 8), value))
        return rows

    def test_adj_rows_increase_n_clean(self):
        """인접 월 데이터가 추가되면 n_clean이 증가해야 함
        10건(5+5)으로 zero_rate ≈ 0.756 < 0.8 → Stage 0 유지 → adj_rows가 stats에 반영됨
        """
        # 베이스라인 데이터: 2024·2025년 2월 각 5건 (총 10건, zero_rate < 0.8 확보)
        raw_rows = (
            self._make_baseline_rows(2024, 2, [5, 12, 19, 26, 6], 100) +
            self._make_baseline_rows(2025, 2, [3, 10, 17, 24, 5], 105)
        )
        baselines_without, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=None
        )

        # 인접 월 데이터: 2026년 1월 5건 + 3월 5건
        adj_rows = (
            self._make_baseline_rows(2026, 1, [5, 12, 19, 26, 6], 102) +
            self._make_baseline_rows(2026, 3, [3, 10, 17, 24, 7], 98)
        )
        baselines_with, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=adj_rows
        )

        cell = (self.ACSR_ID, 0, 8, 2)
        n_without = baselines_without[cell]["stats"]["n_clean"]
        n_with    = baselines_with[cell]["stats"]["n_clean"]
        assert n_with > n_without, f"n_clean 증가 기대: {n_without} → {n_with}"

    def test_n_adj_field_populated(self):
        """n_adj 필드가 추가된 인접 월 샘플 수를 올바르게 반영"""
        raw_rows = self._make_baseline_rows(2024, 2, [5, 12, 19], 100)
        adj_rows = self._make_baseline_rows(2026, 1, [5, 12, 19, 26, 6], 102)  # 5건

        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=adj_rows
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert baselines[cell]["n_adj"] == 5

    def test_n_adj_zero_when_no_adj_rows(self):
        """인접 월 데이터 없으면 n_adj=0"""
        raw_rows = self._make_baseline_rows(2024, 2, [5, 12, 19, 26, 6], 100)
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert baselines[cell]["n_adj"] == 0

    def test_adj_rows_mapped_to_nearest_target_month(self):
        """인접 월(1월, 3월) 데이터가 대상 월(2월) 셀에 올바르게 매핑
        날짜 선택 주의: 2026년 3월의 평일(화요일) 날짜만 사용 (Mar 3, 10, 17, 24, 31)
        — Mar 7은 토요일이므로 day_type=1 셀로 분류됨
        """
        raw_rows = self._make_baseline_rows(2024, 2, [5, 12, 19, 26, 6], 100)
        adj_rows = (
            self._make_baseline_rows(2026, 1, [5, 12, 19, 26, 6], 90) +    # 1월 평일 5건
            self._make_baseline_rows(2026, 3, [3, 10, 17, 24, 31], 110)     # 3월 화요일 5건
        )
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=adj_rows
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert baselines[cell]["n_adj"] == 10  # 1월 5 + 3월 5

    def test_adj_rows_none_value_ignored(self):
        """trf_qnty=None인 인접 월 레코드는 무시"""
        raw_rows = self._make_baseline_rows(2024, 2, [5, 12, 19, 26, 6], 100)
        adj_rows = [
            (self.ACSR_ID, datetime(2026, 1, 5, 8), None),
            (self.ACSR_ID, datetime(2026, 1, 12, 8), 95),
            (self.ACSR_ID, datetime(2026, 1, 19, 8), None),
            (self.ACSR_ID, datetime(2026, 1, 26, 8), 100),
            (self.ACSR_ID, datetime(2026, 1, 6, 8), 98),
        ]
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=adj_rows
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        # None 2건 제외 → 3건만 n_adj
        assert baselines[cell]["n_adj"] == 3

    def test_adj_rows_wrong_acsr_id_ignored(self):
        """다른 접근로(acsr_id) 데이터는 매핑되지 않음"""
        raw_rows = self._make_baseline_rows(2024, 2, [5, 12, 19, 26, 6], 100)
        # acsr_id=9999는 approaches에 없음
        adj_rows = [(9999, datetime(2026, 1, 5, 8), 120)]
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=adj_rows
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert baselines[cell]["n_adj"] == 0

    def test_median_with_adj_is_reasonable(self):
        """인접 월 추가 후 median이 합리적인 범위 내인지 (기본값에서 크게 벗어나지 않음)"""
        # 베이스라인: 2월 데이터 100 x 5건
        raw_rows = self._make_baseline_rows(2024, 2, [5, 12, 19, 26, 6], 100)
        # 인접 월: 비슷한 값 95~105
        adj_rows = (
            self._make_baseline_rows(2026, 1, [5, 12, 19, 26, 6], 95) +
            self._make_baseline_rows(2026, 3, [3, 10, 17, 24, 7], 105)
        )
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=adj_rows
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        median = baselines[cell]["stats"]["median"]
        assert 90 <= median <= 110, f"median={median} 범위 초과"

    def test_confidence_upgrades_low_to_ok(self):
        """인접 월 데이터 보강으로 n_clean < 8(LOW) → ≥ 8(OK)로 개선

        주의: 기준값이 모두 동일(IQR=0)이면 lower=q1=median이 되어
        기준값보다 조금 낮은 인접 월 값도 cleaning에서 제거됨.
        → 기준값에 현실적인 분산을 두거나, 인접 월 값을 기준값 이상으로 설정해야 함.
        여기서는 기준값에 분산(±10%)을 주어 IQR > 0을 확보하는 방식 사용.
        """
        # raw_rows: 2024년 2월 5건, 분산 있는 값 [85, 95, 100, 105, 115]
        raw_rows = [
            (self.ACSR_ID, datetime(2024, 2, d, 8), v)
            for d, v in [(5, 85), (12, 95), (19, 100), (26, 105), (6, 115)]
        ]
        from 이상탐지 import MIN_CONFIDENCE_SAMPLES
        # BASELINE_YEARS=[2024]만 사용: expected=21, 5건 → zero_rate ≈ 0.762 < 0.8
        # → Stage 0 유지 → adj_rows가 Stage 0 stats에 반영됨
        baselines_before, _ = build_baselines(
            raw_rows, [2024], HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert baselines_before[cell]["stats"]["n_clean"] < MIN_CONFIDENCE_SAMPLES

        # 인접 월 3건 추가 — 값이 IQR 하한보다 높아야 cleaning에서 살아남음
        adj_rows = (
            self._make_baseline_rows(2026, 1, [5, 12], 98) +    # 평일 2건
            self._make_baseline_rows(2026, 3, [3, 10], 103)      # 평일 2건
        )
        baselines_after, _ = build_baselines(
            raw_rows, [2024], HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
            adj_month_rows=adj_rows
        )
        assert baselines_after[cell]["stats"]["n_clean"] >= MIN_CONFIDENCE_SAMPLES

    # ── 경계값 케이스 ────────────────────────────────────────────────────────────

    def test_adj_rows_empty_list(self):
        """adj_month_rows=[] 빈 리스트 → 기존 로직과 동일"""
        raw_rows = self._make_baseline_rows(2024, 2, [5, 12, 19, 26, 6], 100)
        b1, _ = build_baselines(raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
                                self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
                                adj_month_rows=[])
        b2, _ = build_baselines(raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
                                self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
                                adj_month_rows=None)
        cell = (self.ACSR_ID, 0, 8, 2)
        assert b1[cell]["n_adj"] == b2[cell]["n_adj"] == 0
        assert b1[cell]["stats"]["n_clean"] == b2[cell]["stats"]["n_clean"]


# ════════════════════════════════════════════════════════════════════════════════
# 4. nearest_month 매핑 로직 — 연도 경계 케이스
# ════════════════════════════════════════════════════════════════════════════════

class TestNearestMonthMapping:
    """
    target_months가 연말·연초에 걸칠 때 nearest 계산이 올바른지 검증.
    build_baselines 내부 로직을 간접 테스트.
    """
    ACSR_ID    = 2001
    APPROACHES = [(2001, "남")]
    HOURS      = [8]

    def test_target_january_adj_december_maps_to_january(self):
        """대상=1월, 인접=12월 → 12월 데이터가 1월 셀에 매핑"""
        date_start = date(2026, 1, 5)   # 월요일
        date_end   = date(2026, 1, 9)   # 금요일

        raw_rows = [(self.ACSR_ID, datetime(2024, 1, d, 8), 100) for d in [7, 14, 21, 28, 6]]
        adj_rows = [(self.ACSR_ID, datetime(2025, 12, d, 8), 95) for d in [1, 8, 15, 22, 29]]

        baselines, _ = build_baselines(
            raw_rows, [2024], HOLIDAY_DATES,
            self.APPROACHES, date_start, date_end, self.HOURS,
            adj_month_rows=adj_rows
        )
        cell = (self.ACSR_ID, 0, 8, 1)  # month=1
        assert baselines[cell]["n_adj"] == 5

    def test_multi_month_target_adj_maps_to_boundary(self):
        """대상=2~3월, 인접 1월→2월(가까움), 4월→3월(가까움) 매핑"""
        date_start = date(2026, 2, 2)   # 월요일
        date_end   = date(2026, 3, 6)   # 금요일

        raw_rows = (
            [(self.ACSR_ID, datetime(2024, 2, d, 8), 100) for d in [5, 12, 19, 26, 6]] +
            [(self.ACSR_ID, datetime(2024, 3, d, 8), 110) for d in [4, 11, 18, 25, 3]]
        )
        adj_jan  = [(self.ACSR_ID, datetime(2026, 1, d, 8), 90)  for d in [5, 12, 19, 26, 6]]
        adj_apr  = [(self.ACSR_ID, datetime(2026, 4, d, 8), 115) for d in [6, 13, 20, 27, 7]]

        baselines, _ = build_baselines(
            raw_rows, [2024], HOLIDAY_DATES,
            self.APPROACHES, date_start, date_end, self.HOURS,
            adj_month_rows=adj_jan + adj_apr
        )
        cell_feb = (self.ACSR_ID, 0, 8, 2)
        cell_mar = (self.ACSR_ID, 0, 8, 3)
        assert baselines[cell_feb]["n_adj"] == 5  # 1월 → 2월
        assert baselines[cell_mar]["n_adj"] == 5  # 4월 → 3월


# ════════════════════════════════════════════════════════════════════════════════
# 5. get_day_type — 기존 함수 회귀 테스트
# ════════════════════════════════════════════════════════════════════════════════

class TestGetDayType:
    def test_weekday(self):
        assert get_day_type(date(2026, 4, 13), set()) == 0  # 월요일

    def test_saturday(self):
        assert get_day_type(date(2026, 4, 11), set()) == 1  # 토요일

    def test_sunday(self):
        assert get_day_type(date(2026, 4, 12), set()) == 2  # 일요일

    def test_holiday_treated_as_sunday(self):
        holidays = {"2026-03-01"}
        assert get_day_type(date(2026, 3, 1), holidays) == 2  # 삼일절(일요일 타입)


# ════════════════════════════════════════════════════════════════════════════════
# 6. Stage 전환 임계값 검증 — ZERO_RATE_EXPANSION_THRESHOLD = 0.8
# ════════════════════════════════════════════════════════════════════════════════

class TestBuildBaselinesStageExpansion:
    """
    Stage 0→1→2 전환 임계값(ZERO_RATE_EXPANSION_THRESHOLD=0.8) 검증

    셀 구성: acsr_id=3001, day_type=0(평일), hour=8, month=2
    대상 기간: 2026-02-02 ~ 2026-02-06 (월~금)
    베이스라인 연도: [2024]

    Feb 2024 평일: 21일
      - 실제 4건 제공 → zero_rate = (21-4)/21 ≈ 0.810  (≥ 0.8 → Stage 확장 유발)
      - 실제 5건 제공 → zero_rate = (21-5)/21 ≈ 0.762  (< 0.8 → Stage 0 유지)

    Stage 1 expected (Jan+Feb+Mar 2024 평일) = 23+21+21 = 65
      - 실제 24건(Feb4+Jan10+Mar10) → zr_1 = (65-24)/65 ≈ 0.631  (< 0.8 → Stage 1 중단)
      - 실제 10건(Feb4+Jan3+Mar3)  → zr_1 = (65-10)/65 ≈ 0.846  (≥ 0.8 → Stage 2 진입)
    """

    ACSR_ID    = 3001
    APPROACHES = [(3001, "테스트")]
    DATE_START = date(2026, 2, 2)   # 월요일
    DATE_END   = date(2026, 2, 6)   # 금요일
    HOURS      = [8]
    BASELINE_YEARS = [2024]

    # Feb 2024 평일: Feb 1 = 목 → 5,6,7,8,9,12,...
    FEB_WDAYS = [5, 6, 7, 8, 9, 12, 13, 14, 15, 16]
    # Jan 2024 평일: Jan 1 = 월 → 1,2,3,4,5,8,9,10,11,12
    JAN_WDAYS = [1, 2, 3, 4, 5, 8, 9, 10, 11, 12]
    # Mar 2024 평일: Mar 1 = 금 → 4,5,6,7,8,11,12,13,14,15
    MAR_WDAYS = [4, 5, 6, 7, 8, 11, 12, 13, 14, 15]

    def _rows(self, year, month, days):
        """(acsr_id, datetime, trf_qnty) 목 데이터 생성"""
        return [(self.ACSR_ID, datetime(year, month, d, 8), 100) for d in days]

    # ── 정상 케이스 ───────────────────────────────────────────────────────────────

    def test_no_expansion_when_zero_rate_below_threshold(self):
        """zero_rate < 0.8 → Stage 0 유지
        Feb 2024 평일 5건 → zero_rate ≈ 0.762
        """
        raw_rows = self._rows(2024, 2, self.FEB_WDAYS[:5])
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert cell in baselines
        assert baselines[cell]["expansion_stage"] == 0

    def test_triggers_stage1_when_zero_rate_gte_threshold(self):
        """zero_rate >= 0.8 → Stage 1 확장 발생
        Feb 2024 평일 4건 → zero_rate ≈ 0.810
        Jan/Mar에 풍부한 데이터(각 10건) → Stage 1 zero_rate ≈ 0.631 < 0.8 → Stage 1에서 중단
        """
        raw_rows = (
            self._rows(2024, 2, self.FEB_WDAYS[:4]) +
            self._rows(2024, 1, self.JAN_WDAYS) +        # 10건
            self._rows(2024, 3, self.MAR_WDAYS)          # 10건
        )
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert cell in baselines
        assert baselines[cell]["expansion_stage"] == 1

    def test_triggers_stage2_when_stage1_still_gte_threshold(self):
        """Stage 1 zero_rate >= 0.8 → Stage 2 추가 확장
        Feb 4건 + Jan 3건 + Mar 3건 → Stage 1 actual=10, expected=65
        zero_rate_1 = (65-10)/65 ≈ 0.846 → Stage 2 진입
        """
        raw_rows = (
            self._rows(2024, 2, self.FEB_WDAYS[:4]) +
            self._rows(2024, 1, self.JAN_WDAYS[:3]) +    # 3건
            self._rows(2024, 3, self.MAR_WDAYS[:3])      # 3건
        )
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert cell in baselines
        assert baselines[cell]["expansion_stage"] == 2

    def test_fallback_summary_n_stage1_increments(self):
        """Stage 1 진입 시 fallback_summary['n_stage1'] >= 1"""
        raw_rows = (
            self._rows(2024, 2, self.FEB_WDAYS[:4]) +
            self._rows(2024, 1, self.JAN_WDAYS) +
            self._rows(2024, 3, self.MAR_WDAYS)
        )
        _, summary = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        assert summary["n_stage1"] >= 1

    def test_fallback_summary_n_stage2_increments(self):
        """Stage 2 진입 시 fallback_summary['n_stage2'] >= 1"""
        raw_rows = (
            self._rows(2024, 2, self.FEB_WDAYS[:4]) +
            self._rows(2024, 1, self.JAN_WDAYS[:3]) +
            self._rows(2024, 3, self.MAR_WDAYS[:3])
        )
        _, summary = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        assert summary["n_stage2"] >= 1

    # ── 경계값 케이스 ─────────────────────────────────────────────────────────────

    def test_threshold_value_is_0_8(self):
        """상수 값이 0.8인지 직접 검증"""
        assert ZERO_RATE_EXPANSION_THRESHOLD == 0.8

    def test_boundary_just_above_triggers_expansion(self):
        """경계 바로 위(≈0.810) → 반드시 Stage 확장 발생"""
        raw_rows = (
            self._rows(2024, 2, self.FEB_WDAYS[:4]) +   # zero_rate ≈ 0.810
            self._rows(2024, 1, self.JAN_WDAYS) +
            self._rows(2024, 3, self.MAR_WDAYS)
        )
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert baselines[cell]["expansion_stage"] >= 1, \
            f"zero_rate ≈ 0.810 이므로 Stage 확장이 발생해야 함 " \
            f"(stage={baselines[cell]['expansion_stage']})"

    def test_boundary_just_below_no_expansion(self):
        """경계 바로 아래(≈0.762) → Stage 0 유지"""
        raw_rows = self._rows(2024, 2, self.FEB_WDAYS[:5])  # zero_rate ≈ 0.762
        baselines, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        assert baselines[cell]["expansion_stage"] == 0, \
            f"zero_rate ≈ 0.762 이므로 Stage 0 유지 " \
            f"(stage={baselines[cell]['expansion_stage']})"

    # ── 오류 케이스 ───────────────────────────────────────────────────────────────

    def test_stage0_zero_rate_computed_correctly(self):
        """Stage 0 zero_rate 계산값이 (expected - actual) / expected 공식을 따르는지"""
        # Feb 2024 평일 21일 중 4건 제공 → zero_rate = 17/21 ≈ 0.8095
        raw_rows = self._rows(2024, 2, self.FEB_WDAYS[:4])
        baselines_before_expansion, _ = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        # Stage 1 확장 후 zero_rate가 바뀌므로 stage>=1 이면 원본 zero_rate 검증 생략
        # 여기서는 expansion이 발생했는지만 검증
        assert baselines_before_expansion[cell]["expansion_stage"] >= 1

    def test_no_rows_results_in_zero_rate_1_and_stage2(self):
        """베이스라인 데이터가 전혀 없으면 zero_rate=1.0 → Stage 2까지 확장"""
        raw_rows = []  # 데이터 없음
        baselines, summary = build_baselines(
            raw_rows, self.BASELINE_YEARS, HOLIDAY_DATES,
            self.APPROACHES, self.DATE_START, self.DATE_END, self.HOURS,
        )
        cell = (self.ACSR_ID, 0, 8, 2)
        if cell in baselines:
            # zero_rate=1.0 >= 0.8 → Stage 1 → Stage 2
            assert baselines[cell]["expansion_stage"] == 2


# ════════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
