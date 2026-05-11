from __future__ import annotations

import re
from datetime import date
from typing import Literal, Optional

from pydantic import BaseModel, model_validator

from traffic_api.domain.time_rules import resolve_hours, validate_hour_values


class JobRequest(BaseModel):
    node_ids: list[str | int]
    date_start: str
    date_end: str
    hours: Optional[list[int]] = None
    hours_preset: Optional[str] = None

    @model_validator(mode="after")
    def validate_fields(self):
        if self.hours is not None and self.hours_preset is not None:
            raise ValueError("hours와 hours_preset은 동시에 지정할 수 없습니다.")
        if not self.node_ids:
            raise ValueError("node_ids는 최소 1개 이상이어야 합니다.")
        date_re = re.compile(r"^\d{4}-\d{2}-\d{2}$")
        if not date_re.match(self.date_start) or not date_re.match(self.date_end):
            raise ValueError("날짜 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        try:
            ds = date.fromisoformat(self.date_start)
            de = date.fromisoformat(self.date_end)
        except ValueError as exc:
            raise ValueError(f"날짜 형식이 올바르지 않습니다 (YYYY-MM-DD): {exc}") from exc
        if ds > de:
            raise ValueError("date_start가 date_end보다 늦습니다.")
        invalid = validate_hour_values(self.hours)
        if invalid:
            raise ValueError(f"hours 값은 0-23 범위여야 합니다: {invalid}")
        if self.hours_preset is not None and self.hours_preset not in ("all", "peak"):
            raise ValueError("hours_preset은 'all' 또는 'peak'만 허용됩니다.")
        return self

    def resolve_hours(self) -> list[int]:
        return resolve_hours(self.hours, self.hours_preset)


class RawTrafficVkndRequest(JobRequest):
    interval: Literal["15m", "1h"] = "1h"
    vknd_codes: Optional[list[str | int]] = None

