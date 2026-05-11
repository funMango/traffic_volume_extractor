from __future__ import annotations


PEAK_HOURS = [7, 8, 12, 13, 17, 18]
ALL_HOURS = list(range(24))


def resolve_hours(hours: list[int] | None, hours_preset: str | None) -> list[int]:
    if hours is not None:
        return sorted(set(hours))
    if hours_preset == "peak":
        return PEAK_HOURS.copy()
    return ALL_HOURS.copy()


def validate_hour_values(hours: list[int] | None) -> list[int]:
    if hours is None:
        return []
    return [hour for hour in hours if not (0 <= hour <= 23)]

