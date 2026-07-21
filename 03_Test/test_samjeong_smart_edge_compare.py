from __future__ import annotations

import importlib.util
import sys
from datetime import date
from pathlib import Path


MODULE_PATH = (
    Path(__file__).resolve().parents[1]
    / "01_Program"
    / "09_else"
    / "samjeong_smart_edge_compare.py"
)
SPEC = importlib.util.spec_from_file_location("samjeong_smart_edge_compare", MODULE_PATH)
comparison = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
sys.modules[SPEC.name] = comparison
SPEC.loader.exec_module(comparison)


def test_target_direction_mapping_includes_samjeong_bridge_west_substitution() -> None:
    targets = {target.label: target for target in comparison.TARGETS}
    assert targets["삼정교사거리"].smart_approaches == ("삼정교사거리-서 (동향)",)
    assert targets["박촌교 삼거리"].smart_approaches == (
        "박촌교삼거리-북 (남향)",
        "박촌교삼거리-남 (북향)",
    )


def test_weekday_excludes_weekends_and_configured_holidays() -> None:
    assert not comparison.is_weekday(date(2026, 5, 1))
    assert not comparison.is_weekday(date(2026, 5, 5))
    assert not comparison.is_weekday(date(2026, 5, 9))
    assert comparison.is_weekday(date(2026, 5, 6))


def test_edge_missing_day_remains_in_denominator_but_smart_missing_day_is_excluded() -> None:
    first = date(2026, 5, 1)
    second = date(2026, 5, 2)
    values = {
        first: comparison.DailyTraffic(100, True),
        second: comparison.DailyTraffic(0, False),
    }
    assert comparison.average_daily(values, (first, second), False) == (50, 2)
    assert comparison.average_daily(values, (first, second), True) == (100, 1)


def test_peak_hours_are_exactly_the_two_half_open_windows() -> None:
    assert comparison.PEAK_HOURS == {7, 8, 17, 18}


def test_rounding_is_half_up() -> None:
    assert comparison.round_half_up(comparison.Decimal("10.5")) == 11


def test_v2_smart_averages_exclude_zero_days_and_holidays() -> None:
    class FakeReader:
        def load(self):
            values = {}
            for target in comparison.TARGETS:
                daily = {}
                for day in comparison.calendar_days(date(2026, 5, 1), date(2026, 6, 30)):
                    daily[day] = (0, 0)
                daily[date(2026, 5, 1)] = (100, 100)  # Holiday: excluded from weekday metrics.
                daily[date(2026, 5, 6)] = (10, 5)
                daily[date(2026, 5, 7)] = (20, 10)
                values[target.label] = daily
            return values

    payload = comparison.smart_v2_payload(FakeReader())
    result = payload[(5, "박촌교 삼거리")]
    assert result == {
        "monthly": 43,
        "weekday": 15,
        "peak": 8,
        "monthly_days": 3,
        "weekday_days": 2,
        "peak_days": 2,
    }
