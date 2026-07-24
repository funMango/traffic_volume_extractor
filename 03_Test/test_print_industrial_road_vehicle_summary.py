from __future__ import annotations

from datetime import datetime
from pathlib import Path
import sys


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "01_Program" / "09_else"
sys.path.insert(0, str(SCRIPT_DIR))

from print_industrial_road_vehicle_summary import (  # noqa: E402
    TARGET_VEHICLES,
    TrafficRecord,
    aggregate_target_vehicles,
    format_summary,
    validate_summary,
)


def test_aggregate_target_vehicles_sums_movements_and_excludes_other_vehicles() -> None:
    records = []
    approaches = (
        "산업길사거리-남 (북향)",
        "산업길사거리-동 (서향)",
        "산업길사거리-북 (남향)",
        "산업길사거리-서 (동향)",
    )
    hours = (datetime(2026, 7, 14, 8), datetime(2026, 7, 14, 9))
    for approach_index, approach in enumerate(approaches):
        for hour_index, observed_at in enumerate(hours):
            for vehicle_index, vehicle in enumerate(TARGET_VEHICLES):
                for movement_volume in (3, 5):
                    records.append(
                        TrafficRecord(
                            observed_at=observed_at,
                            approach=approach,
                            movement="좌회전",
                            vehicle_type=vehicle,
                            traffic_volume=(
                                approach_index + hour_index + vehicle_index + movement_volume
                            ),
                        )
                    )
            records.append(
                TrafficRecord(
                    observed_at=observed_at,
                    approach=approach,
                    movement="우회전",
                    vehicle_type="SUV",
                    traffic_volume=999,
                )
            )

    summary = aggregate_target_vehicles(records)

    validate_summary(summary)
    assert summary[hours[0]][approaches[0]]["세단"] == 8
    assert summary[hours[1]][approaches[3]]["대형트럭"] == 24
    assert "SUV" not in summary[hours[0]][approaches[0]]

    rendered = format_summary(summary)
    assert "7월 14일 | 08:00~09:00" in rendered
    assert "7월 14일 | 09:00~10:00" in rendered
    assert rendered.count("합계") == 8
