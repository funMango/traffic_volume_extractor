from __future__ import annotations

from .codes import normalize_drct_cd


def sum_optional(current, value):
    if value is None:
        return current
    if current is None:
        return value
    return current + value


def aggregate_acsr_slots_from_drct(drct_slots: list[dict]) -> list[dict]:
    agg: dict[tuple, dict] = {}
    for slot in drct_slots:
        if normalize_drct_cd(slot.get("drct_cd")) == "00":
            continue
        key = (
            slot["date"],
            slot["hour"],
            slot["node_id"],
            slot["node_name"],
            slot["approach_id"],
            slot["approach_name"],
        )
        if key not in agg:
            agg[key] = {
                "date": slot["date"],
                "hour": slot["hour"],
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "approach_id": slot["approach_id"],
                "approach_name": slot["approach_name"],
                "traffic_volume": 0,
                "corrected_value": 0,
                "_has_raw": False,
                "_has_corr": False,
                "_has_a": False,
                "_has_b": False,
            }
        acc = agg[key]
        raw_val = slot.get("traffic_volume")
        if raw_val is not None:
            acc["traffic_volume"] += raw_val
            acc["_has_raw"] = True
        corr_val = slot.get("corrected_value")
        if corr_val is None:
            corr_val = raw_val
        if corr_val is not None:
            acc["corrected_value"] += corr_val
            acc["_has_corr"] = True
        anomaly_type = slot.get("anomaly_type")
        if anomaly_type == "A형":
            acc["_has_a"] = True
        elif anomaly_type == "B형":
            acc["_has_b"] = True
        elif anomaly_type == "A+B혼합":
            acc["_has_a"] = True
            acc["_has_b"] = True

    rows: list[dict] = []
    for item in agg.values():
        anomaly_type = None
        if item["_has_a"] and item["_has_b"]:
            anomaly_type = "A+B혼합"
        elif item["_has_a"]:
            anomaly_type = "A형"
        elif item["_has_b"]:
            anomaly_type = "B형"
        rows.append({
            "date": item["date"],
            "hour": item["hour"],
            "node_id": item["node_id"],
            "node_name": item["node_name"],
            "approach_id": item["approach_id"],
            "approach_name": item["approach_name"],
            "traffic_volume": item["traffic_volume"] if item["_has_raw"] else None,
            "anomaly_type": anomaly_type,
            "corrected_value": item["corrected_value"] if item["_has_corr"] else None,
            "correction_method": None,
            "confidence": None,
            "baseline": None,
        })
    rows.sort(key=lambda r: (r["date"], r["node_id"], r["hour"], r["approach_id"]))
    return rows


def aggregate_raw_slots_from_drct(drct_slots: list[dict]) -> list[dict]:
    agg: dict[tuple, dict] = {}
    for slot in drct_slots:
        if normalize_drct_cd(slot.get("drct_cd")) == "00":
            continue
        key = (
            slot["date"],
            slot["hour"],
            slot["node_id"],
            slot["node_name"],
            slot["approach_id"],
            slot["approach_name"],
        )
        if key not in agg:
            agg[key] = {
                "date": slot["date"],
                "hour": slot["hour"],
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "approach_id": slot["approach_id"],
                "approach_name": slot["approach_name"],
                "traffic_volume": None,
            }
        agg[key]["traffic_volume"] = sum_optional(
            agg[key]["traffic_volume"],
            slot.get("traffic_volume"),
        )
    rows = list(agg.values())
    rows.sort(key=lambda r: (r["date"], r["node_id"], r["hour"], r["approach_id"]))
    return rows


def aggregate_raw_node_slots(slots: list[dict]) -> list[dict]:
    agg: dict[tuple, dict] = {}
    for slot in slots:
        key = (slot["date"], slot["hour"], slot["node_id"], slot["node_name"])
        if key not in agg:
            agg[key] = {
                "date": slot["date"],
                "hour": slot["hour"],
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "traffic_volume": None,
            }
        agg[key]["traffic_volume"] = sum_optional(
            agg[key]["traffic_volume"],
            slot.get("traffic_volume"),
        )
    rows = list(agg.values())
    rows.sort(key=lambda r: (r["date"], r["node_id"], r["hour"]))
    return rows

