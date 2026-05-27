#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""교통량 이상탐지 FastAPI 서버

# 실행 방법
# 1. cd "01_Program"
# 2. uvicorn api_server:app --host 0.0.0.0 --port 8000 --reload
# API test: http://localhost:8000/docs

엔드포인트:
  POST /jobs              분석 작업 시작 → job_id 반환
  GET  /jobs/{job_id}     작업 상태·결과 조회
  POST /corrected-traffic ACSR 단위 보정 결과 즉시 조회
  POST /corrected-traffic-drct DRCT 단위 보정 + ACSR 집계 즉시 조회
  GET  /intersections     교차로 목록 조회
"""

import asyncio
import json
import sys
import uuid
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from datetime import date, datetime, timedelta
from pathlib import Path
from threading import Lock
from typing import Literal, Optional
import re

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, model_validator
from traffic_api.domain.aggregation import (
    aggregate_crsrd_slots_from_acsr as _domain_aggregate_crsrd_slots_from_acsr,
)

# 이상탐지 모듈 import (경로는 이상탐지.py의 __file__ 기준으로 동작)
sys.path.insert(0, str(Path(__file__).parent))
import 이상탐지 as ad

# ── 전역 상태 ─────────────────────────────────────────────────────────────────

executor = ThreadPoolExecutor(max_workers=4)

job_store: dict[str, dict] = {}
job_store_lock = Lock()

_intersections_cache: list[dict] | None = None
_holiday_dates: set | None = None
_drct_direction_presence_cache: dict | None = None
_drct_direction_presence_lock = Lock()

DRCT_DIRECTION_PRESENCE_PATH = (
    Path(__file__).parent / "traffic_api" / "cache" / "drct_direction_presence.json"
)
DRCT_DIRECTION_PRESENCE_BASIS = "last_1_year_1h_positive"


# ── 라이프사이클 ──────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    global _holiday_dates
    _holiday_dates = ad.load_holidays()
    _ensure_drct_direction_presence_cache()
    yield


# ── 앱 초기화 ─────────────────────────────────────────────────────────────────

app = FastAPI(
    title="교통량 이상탐지 API",
    version="1.0.0",
    description="교통량 이상값 탐지·보정 결과를 JSON으로 제공합니다.",
    lifespan=lifespan,
)


# ── Pydantic 모델 ─────────────────────────────────────────────────────────────

class JobRequest(BaseModel):
    node_ids: list[str | int]
    date_start: str       # "YYYY-MM-DD"
    date_end: str         # "YYYY-MM-DD"
    hours: Optional[list[int]] = None
    hours_preset: Optional[str] = None  # "all" | "peak"

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
        except ValueError as e:
            raise ValueError(f"날짜 형식이 올바르지 않습니다 (YYYY-MM-DD): {e}") from e
        if ds > de:
            raise ValueError("date_start가 date_end보다 늦습니다.")
        if self.hours is not None:
            invalid = [h for h in self.hours if not (0 <= h <= 23)]
            if invalid:
                raise ValueError(f"hours 값은 0–23 범위여야 합니다: {invalid}")
        if self.hours_preset is not None and self.hours_preset not in ("all", "peak"):
            raise ValueError("hours_preset은 'all' 또는 'peak'만 허용됩니다.")
        return self

    def resolve_hours(self) -> list[int]:
        if self.hours is not None:
            return sorted(set(self.hours))
        if self.hours_preset == "peak":
            return [7, 8, 12, 13, 17, 18]
        return list(range(24))  # "all" 또는 미지정


class RawTrafficVkndRequest(JobRequest):
    interval: Literal["15m", "1h"] = "1h"
    vknd_codes: Optional[list[str | int]] = None


class RawTrafficDrctRequest(JobRequest):
    interval: Literal["5m", "15m", "1h", "1d"] = "1h"
    approach_ids: Optional[list[str | int]] = None
    drct_codes: Optional[list[str | int]] = None


def _job_request_payload(req: JobRequest) -> dict:
    payload = req.model_dump()
    payload["hours"] = req.resolve_hours()
    payload.pop("hours_preset", None)
    return payload


def _req_get(req, key: str, default=None):
    if isinstance(req, dict):
        return req.get(key, default)
    return getattr(req, key, default)


def _req_hours(req) -> list[int]:
    if isinstance(req, dict):
        hours = req.get("hours")
        return sorted(set(hours)) if hours is not None else list(range(24))
    return req.resolve_hours()


# ── 헬퍼: 슬롯 직렬화 ────────────────────────────────────────────────────────

def _serialize_slot(r: dict, node_id: int, baselines: dict) -> dict:
    """이상탐지 결과 dict → API 응답 스키마로 변환"""
    cell = r["_cell"]
    bl   = baselines.get(cell, {})
    stats = bl.get("stats")
    return {
        "date":              r["날짜"].replace(".", "-"),   # "2026.03.15" → "2026-03-15"
        "hour":              r["_hour"],
        "node_id":           node_id,
        "node_name":         r["교차로"],
        "approach_id":       r["_acsr_id"],
        "approach_name":     r["방향"],
        "traffic_volume":    r["교통량"],
        "anomaly_type":      r["판정"],
        "corrected_value":   r.get("보정값"),
        "correction_method": r.get("보정방법"),
        "confidence":        r.get("신뢰도"),
        "baseline": {
            "median":          stats["median"]  if stats else None,
            "q1":              stats["q1"]      if stats else None,
            "q3":              stats["q3"]      if stats else None,
            "n_clean":         stats["n_clean"] if stats else None,
            "zero_rate":       bl.get("zero_rate"),
            "expansion_stage": bl.get("expansion_stage", 0),
        },
    }


def _serialize_drct_slot(r: dict, node_id: int, baselines: dict, drct_meta: dict) -> dict:
    """DRCT 단위 이상탐지 결과 dict → API 응답 스키마로 변환."""
    drct_key = r["_acsr_id"]  # (acsr_id, drct_cd)
    meta = drct_meta.get(drct_key, {})
    approach_id = meta.get("approach_id")
    drct_cd = meta.get("drct_cd")
    if approach_id is None and isinstance(drct_key, tuple) and len(drct_key) >= 1:
        approach_id = drct_key[0]
    if drct_cd is None and isinstance(drct_key, tuple) and len(drct_key) >= 2:
        drct_cd = drct_key[1]

    cell = r["_cell"]
    bl = baselines.get(cell, {})
    stats = bl.get("stats")

    return {
        "date":              r["날짜"].replace(".", "-"),
        "hour":              r["_hour"],
        "node_id":           node_id,
        "node_name":         r["교차로"],
        "approach_id":       approach_id,
        "approach_name":     meta.get("approach_name", str(approach_id)),
        "drct_cd":           drct_cd,
        "drct_name":         meta.get("drct_name", str(drct_cd)),
        "traffic_volume":    r["교통량"],
        "anomaly_type":      r["판정"],
        "corrected_value":   r.get("보정값"),
        "correction_method": r.get("보정방법"),
        "confidence":        r.get("신뢰도"),
        "baseline": {
            "median":          stats["median"] if stats else None,
            "q1":              stats["q1"] if stats else None,
            "q3":              stats["q3"] if stats else None,
            "n_clean":         stats["n_clean"] if stats else None,
            "zero_rate":       bl.get("zero_rate"),
            "expansion_stage": bl.get("expansion_stage", 0),
        },
    }


def _normalize_drct_cd(drct_cd) -> str:
    if drct_cd is None:
        return ""
    return str(drct_cd).strip().zfill(2)


def _normalize_acsr_id(acsr_id):
    if isinstance(acsr_id, int):
        return acsr_id
    if isinstance(acsr_id, str) and acsr_id.strip().isdigit():
        return int(acsr_id.strip())
    return acsr_id


def _new_drct_direction_presence_cache() -> dict:
    return {
        "basis": DRCT_DIRECTION_PRESENCE_BASIS,
        "generated_at": None,
        "nodes": {},
    }


def _load_drct_direction_presence_cache() -> dict:
    global _drct_direction_presence_cache
    with _drct_direction_presence_lock:
        if _drct_direction_presence_cache is not None:
            return _drct_direction_presence_cache
        if DRCT_DIRECTION_PRESENCE_PATH.exists():
            with DRCT_DIRECTION_PRESENCE_PATH.open("r", encoding="utf-8") as f:
                _drct_direction_presence_cache = json.load(f)
        else:
            _drct_direction_presence_cache = _new_drct_direction_presence_cache()
        _drct_direction_presence_cache.setdefault("basis", DRCT_DIRECTION_PRESENCE_BASIS)
        _drct_direction_presence_cache.setdefault("nodes", {})
        return _drct_direction_presence_cache


def _save_drct_direction_presence_cache(cache: dict) -> None:
    DRCT_DIRECTION_PRESENCE_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = DRCT_DIRECTION_PRESENCE_PATH.with_suffix(".json.tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2, sort_keys=True)
        f.write("\n")
    tmp_path.replace(DRCT_DIRECTION_PRESENCE_PATH)


def _set_drct_direction_presence(
    cache: dict,
    node_id,
    approach_id,
    drct_cd,
    exists: bool,
    checked_date: str | None = None,
) -> None:
    node_key = str(node_id)
    approach_key = str(approach_id)
    drct_key = _normalize_drct_cd(drct_cd)
    checked_date = checked_date or date.today().isoformat()
    node = cache.setdefault("nodes", {}).setdefault(node_key, {"approaches": {}})
    approach = node.setdefault("approaches", {}).setdefault(approach_key, {"directions": {}})
    approach.setdefault("directions", {})[drct_key] = {
        "exists": bool(exists),
        "last_checked_date": checked_date,
        "basis": DRCT_DIRECTION_PRESENCE_BASIS,
    }


def _get_drct_direction_presence(cache: dict, node_id, approach_id, drct_cd) -> dict | None:
    return (
        cache.get("nodes", {})
        .get(str(node_id), {})
        .get("approaches", {})
        .get(str(approach_id), {})
        .get("directions", {})
        .get(_normalize_drct_cd(drct_cd))
    )


def _build_drct_direction_presence_cache(conn) -> dict:
    cutoff = datetime.combine(date.today() - timedelta(days=365), datetime.min.time())
    checked_date = date.today().isoformat()
    cache = _new_drct_direction_presence_cache()
    cache["generated_at"] = datetime.now().isoformat(timespec="seconds")
    sql = """
        SELECT NODE_ID, ACSR_ID, TRIM(TO_CHAR(DRCT_CD)) AS DRCT_CD
        FROM S_CRSRD_DRCT_TRF_1HH
        WHERE TOT_DT >= :cutoff
          AND NVL(TRF_QNTY, 0) > 0
          AND TRIM(TO_CHAR(DRCT_CD)) <> '00'
        GROUP BY NODE_ID, ACSR_ID, TRIM(TO_CHAR(DRCT_CD))
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, cutoff=cutoff)
        for node_id, approach_id, drct_cd in cur.fetchall():
            _set_drct_direction_presence(
                cache,
                node_id,
                approach_id,
                drct_cd,
                True,
                checked_date,
            )
    return cache


def _ensure_drct_direction_presence_cache(conn=None) -> dict:
    cache = _load_drct_direction_presence_cache()
    if DRCT_DIRECTION_PRESENCE_PATH.exists():
        return cache

    owns_conn = conn is None
    if owns_conn:
        conn = ad.connect_db()
    try:
        cache = _build_drct_direction_presence_cache(conn)
        with _drct_direction_presence_lock:
            global _drct_direction_presence_cache
            _drct_direction_presence_cache = cache
            _save_drct_direction_presence_cache(cache)
        return cache
    finally:
        if owns_conn:
            conn.close()


def _refresh_drct_direction_presence(
    conn,
    node_id,
    approach_ids: list,
    drct_codes: list[str],
) -> dict:
    cache = _load_drct_direction_presence_cache()
    if not approach_ids or not drct_codes:
        return cache

    cutoff = datetime.combine(date.today() - timedelta(days=365), datetime.min.time())
    checked_date = date.today().isoformat()
    bind_params = {"nid": node_id, "cutoff": cutoff}
    acsr_names = []
    for idx, approach_id in enumerate(approach_ids):
        name = f"acsr{idx}"
        bind_params[name] = approach_id
        acsr_names.append(f":{name}")
    drct_names = []
    for idx, drct_cd in enumerate(drct_codes):
        name = f"drct{idx}"
        bind_params[name] = drct_cd
        drct_names.append(f":{name}")

    sql = f"""
        SELECT ACSR_ID, TRIM(TO_CHAR(DRCT_CD)) AS DRCT_CD
        FROM S_CRSRD_DRCT_TRF_1HH
        WHERE NODE_ID = :nid
          AND TOT_DT >= :cutoff
          AND NVL(TRF_QNTY, 0) > 0
          AND ACSR_ID IN ({','.join(acsr_names)})
          AND TRIM(TO_CHAR(DRCT_CD)) IN ({','.join(drct_names)})
          AND TRIM(TO_CHAR(DRCT_CD)) <> '00'
        GROUP BY ACSR_ID, TRIM(TO_CHAR(DRCT_CD))
    """
    positive = set()
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, **bind_params)
        for approach_id, drct_cd in cur.fetchall():
            positive.add((_normalize_acsr_id(approach_id), _normalize_drct_cd(drct_cd)))

    with _drct_direction_presence_lock:
        for approach_id in approach_ids:
            for drct_cd in drct_codes:
                _set_drct_direction_presence(
                    cache,
                    node_id,
                    approach_id,
                    drct_cd,
                    (_normalize_acsr_id(approach_id), _normalize_drct_cd(drct_cd)) in positive,
                    checked_date,
                )
        cache["generated_at"] = cache.get("generated_at") or datetime.now().isoformat(timespec="seconds")
        _save_drct_direction_presence_cache(cache)
    return cache


def _ensure_drct_direction_presence_entries(
    conn,
    node_id,
    approach_ids: list,
    drct_codes: list[str],
) -> dict:
    cache = _ensure_drct_direction_presence_cache(conn)
    today = date.today().isoformat()
    needs_refresh = False
    for approach_id in approach_ids:
        for drct_cd in drct_codes:
            entry = _get_drct_direction_presence(cache, node_id, approach_id, drct_cd)
            if entry is None or entry.get("last_checked_date") != today:
                needs_refresh = True
                break
        if needs_refresh:
            break
    if needs_refresh:
        cache = _refresh_drct_direction_presence(conn, node_id, approach_ids, drct_codes)
    return cache


def _aggregate_acsr_slots_from_drct(drct_slots: list[dict]) -> list[dict]:
    """DRCT 슬롯 목록을 ACSR+시각 단위로 집계(전/후 합계 + A/B/혼합 판정)."""
    agg: dict[tuple, dict] = {}
    for slot in drct_slots:
        if _normalize_drct_cd(slot.get("drct_cd")) == "00":
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
                "date":              slot["date"],
                "hour":              slot["hour"],
                "node_id":           slot["node_id"],
                "node_name":         slot["node_name"],
                "approach_id":       slot["approach_id"],
                "approach_name":     slot["approach_name"],
                "traffic_volume":    0,
                "corrected_value":   0,
                "_has_raw":          False,
                "_has_corr":         False,
                "_has_a":            False,
                "_has_b":            False,
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
    for _, item in agg.items():
        anomaly_type = None
        if item["_has_a"] and item["_has_b"]:
            anomaly_type = "A+B혼합"
        elif item["_has_a"]:
            anomaly_type = "A형"
        elif item["_has_b"]:
            anomaly_type = "B형"
        rows.append({
            "date":              item["date"],
            "hour":              item["hour"],
            "node_id":           item["node_id"],
            "node_name":         item["node_name"],
            "approach_id":       item["approach_id"],
            "approach_name":     item["approach_name"],
            "traffic_volume":    item["traffic_volume"] if item["_has_raw"] else None,
            "anomaly_type":      anomaly_type,
            "corrected_value":   item["corrected_value"] if item["_has_corr"] else None,
            "correction_method": None,
            "confidence":        None,
            "baseline":          None,
        })
    rows.sort(key=lambda r: (r["date"], r["node_id"], r["hour"], r["approach_id"]))
    return rows


# ── 헬퍼: job 직렬화 ─────────────────────────────────────────────────────────

def _aggregate_crsrd_slots_from_acsr(acsr_slots: list[dict]) -> list[dict]:
    return _domain_aggregate_crsrd_slots_from_acsr(acsr_slots)


def _sum_optional(current, value):
    if value is None:
        return current
    if current is None:
        return value
    return current + value


def _aggregate_raw_slots_from_drct(drct_slots: list[dict]) -> list[dict]:
    """Aggregate raw DRCT rows into approach slots with NULL-aware sums."""
    agg: dict[tuple, dict] = {}
    for slot in drct_slots:
        if _normalize_drct_cd(slot.get("drct_cd")) == "00":
            continue
        key = (
            slot.get("interval"),
            slot.get("timestamp"),
            slot["date"],
            slot["hour"],
            slot.get("minute"),
            slot["node_id"],
            slot["node_name"],
            slot["approach_id"],
            slot["approach_name"],
        )
        if key not in agg:
            agg[key] = {
                **({"interval": slot.get("interval")} if "interval" in slot else {}),
                **({"timestamp": slot.get("timestamp")} if "timestamp" in slot else {}),
                "date": slot["date"],
                "hour": slot["hour"],
                **({"minute": slot.get("minute")} if "minute" in slot else {}),
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "approach_id": slot["approach_id"],
                "approach_name": slot["approach_name"],
                "traffic_volume": None,
            }
        agg[key]["traffic_volume"] = _sum_optional(
            agg[key]["traffic_volume"],
            slot.get("traffic_volume"),
        )

    rows = list(agg.values())
    rows.sort(key=lambda r: (r.get("timestamp", r["date"]), r["node_id"], r["hour"], r["approach_id"]))
    return rows


def _aggregate_raw_node_slots(slots: list[dict]) -> list[dict]:
    """Aggregate approach slots into node slots with NULL-aware sums."""
    agg: dict[tuple, dict] = {}
    for slot in slots:
        key = (
            slot.get("interval"),
            slot.get("timestamp"),
            slot["date"],
            slot["hour"],
            slot.get("minute"),
            slot["node_id"],
            slot["node_name"],
        )
        if key not in agg:
            agg[key] = {
                **({"interval": slot.get("interval")} if "interval" in slot else {}),
                **({"timestamp": slot.get("timestamp")} if "timestamp" in slot else {}),
                "date": slot["date"],
                "hour": slot["hour"],
                **({"minute": slot.get("minute")} if "minute" in slot else {}),
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "traffic_volume": None,
            }
        agg[key]["traffic_volume"] = _sum_optional(
            agg[key]["traffic_volume"],
            slot.get("traffic_volume"),
        )

    rows = list(agg.values())
    rows.sort(key=lambda r: (r.get("timestamp", r["date"]), r["node_id"], r["hour"]))
    return rows


def _normalize_vknd_cd(vknd_cd) -> str:
    if vknd_cd is None:
        return ""
    return str(vknd_cd).strip()


def _aggregate_raw_slots_from_vknd(vknd_slots: list[dict]) -> list[dict]:
    """Aggregate raw VKND rows into approach+vehicle-kind slots with NULL-aware sums."""
    agg: dict[tuple, dict] = {}
    for slot in vknd_slots:
        key = (
            slot["interval"],
            slot["timestamp"],
            slot["date"],
            slot["hour"],
            slot["minute"],
            slot["node_id"],
            slot["node_name"],
            slot["approach_id"],
            slot["approach_name"],
            slot["vknd_cd"],
            slot["vknd_name"],
        )
        if key not in agg:
            agg[key] = {
                "interval": slot["interval"],
                "timestamp": slot["timestamp"],
                "date": slot["date"],
                "hour": slot["hour"],
                "minute": slot["minute"],
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "approach_id": slot["approach_id"],
                "approach_name": slot["approach_name"],
                "vknd_cd": slot["vknd_cd"],
                "vknd_name": slot["vknd_name"],
                "traffic_volume": None,
            }
        agg[key]["traffic_volume"] = _sum_optional(
            agg[key]["traffic_volume"],
            slot.get("traffic_volume"),
        )

    rows = list(agg.values())
    rows.sort(
        key=lambda r: (
            r["timestamp"],
            r["node_id"],
            r["approach_id"],
            r["vknd_cd"],
        )
    )
    return rows


def _aggregate_raw_node_slots_from_vknd(slots: list[dict]) -> list[dict]:
    """Aggregate approach+vehicle-kind slots into node+vehicle-kind slots."""
    agg: dict[tuple, dict] = {}
    for slot in slots:
        key = (
            slot["interval"],
            slot["timestamp"],
            slot["date"],
            slot["hour"],
            slot["minute"],
            slot["node_id"],
            slot["node_name"],
            slot["vknd_cd"],
            slot["vknd_name"],
        )
        if key not in agg:
            agg[key] = {
                "interval": slot["interval"],
                "timestamp": slot["timestamp"],
                "date": slot["date"],
                "hour": slot["hour"],
                "minute": slot["minute"],
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "vknd_cd": slot["vknd_cd"],
                "vknd_name": slot["vknd_name"],
                "traffic_volume": None,
            }
        agg[key]["traffic_volume"] = _sum_optional(
            agg[key]["traffic_volume"],
            slot.get("traffic_volume"),
        )

    rows = list(agg.values())
    rows.sort(key=lambda r: (r["timestamp"], r["node_id"], r["vknd_cd"]))
    return rows


def _aggregate_corrected_slots_from_vknd(vknd_slots: list[dict]) -> list[dict]:
    """Aggregate corrected VKND rows into approach+vehicle-kind slots."""
    agg: dict[tuple, dict] = {}
    for slot in vknd_slots:
        key = (
            slot["interval"],
            slot["timestamp"],
            slot["date"],
            slot["hour"],
            slot["minute"],
            slot["node_id"],
            slot["node_name"],
            slot["approach_id"],
            slot["approach_name"],
            slot["vknd_cd"],
            slot["vknd_name"],
        )
        if key not in agg:
            agg[key] = {
                "interval": slot["interval"],
                "timestamp": slot["timestamp"],
                "date": slot["date"],
                "hour": slot["hour"],
                "minute": slot["minute"],
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "approach_id": slot["approach_id"],
                "approach_name": slot["approach_name"],
                "vknd_cd": slot["vknd_cd"],
                "vknd_name": slot["vknd_name"],
                "traffic_volume": None,
                "corrected_value": None,
                "_has_a": False,
                "_has_b": False,
            }
        acc = agg[key]
        acc["traffic_volume"] = _sum_optional(acc["traffic_volume"], slot.get("traffic_volume"))
        corr_val = slot.get("corrected_value")
        if corr_val is None:
            corr_val = slot.get("traffic_volume")
        acc["corrected_value"] = _sum_optional(acc["corrected_value"], corr_val)
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
            "interval": item["interval"],
            "timestamp": item["timestamp"],
            "date": item["date"],
            "hour": item["hour"],
            "minute": item["minute"],
            "node_id": item["node_id"],
            "node_name": item["node_name"],
            "approach_id": item["approach_id"],
            "approach_name": item["approach_name"],
            "vknd_cd": item["vknd_cd"],
            "vknd_name": item["vknd_name"],
            "traffic_volume": item["traffic_volume"],
            "anomaly_type": anomaly_type,
            "corrected_value": item["corrected_value"],
            "correction_method": None,
            "confidence": None,
            "baseline": None,
        })
    rows.sort(key=lambda r: (r["timestamp"], r["node_id"], r["approach_id"], r["vknd_cd"]))
    return rows


def _aggregate_corrected_node_slots_from_vknd(slots: list[dict]) -> list[dict]:
    """Aggregate corrected approach+vehicle-kind slots into node+vehicle-kind slots."""
    agg: dict[tuple, dict] = {}
    for slot in slots:
        key = (
            slot["interval"],
            slot["timestamp"],
            slot["date"],
            slot["hour"],
            slot["minute"],
            slot["node_id"],
            slot["node_name"],
            slot["vknd_cd"],
            slot["vknd_name"],
        )
        if key not in agg:
            agg[key] = {
                "interval": slot["interval"],
                "timestamp": slot["timestamp"],
                "date": slot["date"],
                "hour": slot["hour"],
                "minute": slot["minute"],
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "vknd_cd": slot["vknd_cd"],
                "vknd_name": slot["vknd_name"],
                "traffic_volume": None,
                "corrected_value": None,
                "_has_a": False,
                "_has_b": False,
            }
        acc = agg[key]
        acc["traffic_volume"] = _sum_optional(acc["traffic_volume"], slot.get("traffic_volume"))
        corr_val = slot.get("corrected_value")
        if corr_val is None:
            corr_val = slot.get("traffic_volume")
        acc["corrected_value"] = _sum_optional(acc["corrected_value"], corr_val)
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
            "interval": item["interval"],
            "timestamp": item["timestamp"],
            "date": item["date"],
            "hour": item["hour"],
            "minute": item["minute"],
            "node_id": item["node_id"],
            "node_name": item["node_name"],
            "vknd_cd": item["vknd_cd"],
            "vknd_name": item["vknd_name"],
            "traffic_volume": item["traffic_volume"],
            "anomaly_type": anomaly_type,
            "corrected_value": item["corrected_value"],
            "correction_method": None,
            "confidence": None,
            "baseline": None,
        })
    rows.sort(key=lambda r: (r["timestamp"], r["node_id"], r["vknd_cd"]))
    return rows


def _serialize_job(job: dict) -> dict:
    def fmt(dt: datetime | None) -> str | None:
        return dt.isoformat() if dt else None

    return {
        "job_id":      job["job_id"],
        "status":      job["status"],
        "created_at":  fmt(job["created_at"]),
        "started_at":  fmt(job["started_at"]),
        "finished_at": fmt(job["finished_at"]),
        "progress":    job["progress"],
        "result":      job["result"],
        "error":       job["error"],
    }


# ── 백그라운드 작업 ───────────────────────────────────────────────────────────

def _run_job(job_id: str) -> None:
    """ThreadPoolExecutor 스레드에서 실행되는 분석 작업"""
    job = job_store[job_id]
    req = job["request"]

    with job_store_lock:
        job["status"]     = "running"
        job["started_at"] = datetime.now()

    conn = None
    try:
        if _holiday_dates is None:
            raise RuntimeError("서버 초기화가 완료되지 않았습니다. 잠시 후 다시 시도하세요.")

        conn = ad.connect_db()

        date_start = date.fromisoformat(req["date_start"])
        date_end   = date.fromisoformat(req["date_end"])
        hours      = req["hours"]

        baseline_years = ad.select_baseline_years(date_start.year)
        fallback_year  = ad.select_fallback_year(date_start.year)
        adj_periods    = ad.get_adjacent_month_periods(date_start, date_end)

        # node_id → 교차로명 매핑
        id_to_name = {nid: nm for nid, nm in ad.load_intersections(conn)}

        node_ids = req["node_ids"]
        job["progress"] = {"total": len(node_ids), "done": 0, "current_node": None}

        all_slots: list[dict] = []

        for idx, node_id in enumerate(node_ids):
            node_name = id_to_name.get(node_id, str(node_id))
            job["progress"]["current_node"] = node_name

            node_results, baselines, _ = ad.analyse_node(
                conn, node_id, node_name,
                date_start, date_end, hours,
                baseline_years, fallback_year, adj_periods,
                _holiday_dates,
            )

            slots = [_serialize_slot(r, node_id, baselines) for r in node_results]
            all_slots.extend(slots)
            job["progress"]["done"] = idx + 1

        job["progress"]["current_node"] = None

        count_a = sum(1 for s in all_slots if s["anomaly_type"] == "A형")
        count_b = sum(1 for s in all_slots if s["anomaly_type"] == "B형")
        count_n = sum(1 for s in all_slots if s["anomaly_type"] == "정상")

        job["result"] = {
            "summary": {
                "total_anomalies": len(all_slots),
                "count_A":         count_a,
                "count_B":         count_b,
                "count_normal_stage2": count_n,
                "baseline_years":  baseline_years,
                "fallback_year":   fallback_year,
                "adj_periods":     adj_periods,
            },
            "slots": all_slots,
        }
        job["status"] = "done"

    except Exception as e:
        job["status"] = "failed"
        job["error"]  = str(e)

    finally:
        if conn:
            conn.close()
        job["finished_at"] = datetime.now()


def _fetch_intersections() -> list[dict]:
    """동기 DB 조회 (executor에서 실행)"""
    conn = ad.connect_db()
    try:
        rows = ad.load_intersections(conn)
        return [{"node_id": nid, "name": nm} for nid, nm in rows]
    finally:
        conn.close()


def _get_id_to_name_map(conn) -> dict:
    """교차로 id→name 맵 반환 (가능하면 메모리 캐시 재사용)"""
    global _intersections_cache
    if _intersections_cache is not None:
        return {it["node_id"]: it["name"] for it in _intersections_cache}
    return {nid: nm for nid, nm in ad.load_intersections(conn)}


def _normalize_node_id(node_id: str | int):
    if isinstance(node_id, int):
        return node_id
    if isinstance(node_id, str) and node_id.isdigit():
        return int(node_id)
    return node_id


def _fetch_corrected_traffic(req: JobRequest) -> dict:
    """이상탐지+보정 결과를 동기적으로 계산 (executor에서 실행)

    정상 슬롯도 포함해 반환하므로 이상 없을 때도 원본 교통량 그래프를 그릴 수 있다.
    """
    if _holiday_dates is None:
        raise RuntimeError("서버 초기화가 완료되지 않았습니다. 잠시 후 다시 시도하세요.")

    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(_req_get(req, "date_start"))
        date_end   = date.fromisoformat(_req_get(req, "date_end"))
        hours      = _req_hours(req)

        baseline_years = ad.select_baseline_years(date_start.year)
        fallback_year  = ad.select_fallback_year(date_start.year)
        adj_periods    = ad.get_adjacent_month_periods(date_start, date_end)

        id_to_name = _get_id_to_name_map(conn)

        all_slots: list[dict] = []
        for node_id in _req_get(req, "node_ids", []):
            norm_node_id = _normalize_node_id(node_id)
            node_name = id_to_name.get(norm_node_id, str(node_id))
            node_results, baselines, target_data = ad.analyse_node(
                conn, norm_node_id, node_name,
                date_start, date_end, hours,
                baseline_years, fallback_year, adj_periods,
                _holiday_dates,
            )

            # approach name lookup — str 정규화로 Oracle 타입 불일치 방지
            acsr_names: dict = {str(aid): anm for aid, anm in ad.load_approaches(conn, norm_node_id)}
            for r in node_results:
                acsr_names.setdefault(str(r["_acsr_id"]), r["방향"])

            # 이상 슬롯 인덱스 (중복 방지)
            anomaly_index = {
                (r["_acsr_id"], r["_date"], r["_hour"]): r
                for r in node_results
            }

            # target_data 전체를 슬롯으로 직렬화 (이상 슬롯은 보정 정보 포함)
            for (acsr_id, d, h), vol in target_data.items():
                if (acsr_id, d, h) in anomaly_index:
                    all_slots.append(_serialize_slot(anomaly_index[(acsr_id, d, h)], node_id, baselines))
                else:
                    all_slots.append({
                        "date":              d.isoformat(),
                        "hour":              h,
                        "node_id":           norm_node_id,
                        "node_name":         node_name,
                        "approach_id":       acsr_id,
                        "approach_name":     acsr_names.get(str(acsr_id), str(acsr_id)),
                        "traffic_volume":    vol,
                        "anomaly_type":      None,
                        "corrected_value":   None,
                        "correction_method": None,
                        "confidence":        None,
                        "baseline":          None,
                    })

        return {"slots": all_slots}
    finally:
        conn.close()


def _fetch_corrected_traffic_drct(req: JobRequest) -> dict:
    """DRCT 단위 이상탐지+보정 결과와 ACSR 집계 결과를 함께 반환."""
    if _holiday_dates is None:
        raise RuntimeError("서버 초기화가 완료되지 않았습니다. 잠시 후 다시 시도하세요.")

    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(req.date_start)
        date_end = date.fromisoformat(req.date_end)
        hours = req.resolve_hours()

        baseline_years = ad.select_baseline_years(date_start.year)
        fallback_year = ad.select_fallback_year(date_start.year)
        adj_periods = ad.get_adjacent_month_periods(date_start, date_end)

        id_to_name = _get_id_to_name_map(conn)

        drct_slots: list[dict] = []
        for node_id in req.node_ids:
            norm_node_id = _normalize_node_id(node_id)
            node_name = id_to_name.get(norm_node_id, str(node_id))
            node_results, baselines, target_data, drct_meta = ad.analyse_node_drct(
                conn,
                norm_node_id,
                node_name,
                date_start,
                date_end,
                hours,
                baseline_years,
                fallback_year,
                adj_periods,
                _holiday_dates,
            )

            anomaly_index = {
                (r["_acsr_id"], r["_date"], r["_hour"]): r
                for r in node_results
            }
            seen_positions: set[tuple] = set()

            for (drct_key, d, h), vol in target_data.items():
                position = (drct_key, d, h)
                seen_positions.add(position)
                if position in anomaly_index:
                    drct_slots.append(_serialize_drct_slot(anomaly_index[position], norm_node_id, baselines, drct_meta))
                    continue

                meta = drct_meta.get(drct_key, {})
                approach_id = meta.get("approach_id")
                drct_cd = meta.get("drct_cd")
                if approach_id is None and isinstance(drct_key, tuple) and len(drct_key) >= 1:
                    approach_id = drct_key[0]
                if drct_cd is None and isinstance(drct_key, tuple) and len(drct_key) >= 2:
                    drct_cd = drct_key[1]
                drct_slots.append({
                    "date":              d.isoformat(),
                    "hour":              h,
                    "node_id":           norm_node_id,
                    "node_name":         node_name,
                    "approach_id":       approach_id,
                    "approach_name":     meta.get("approach_name", str(approach_id)),
                    "drct_cd":           drct_cd,
                    "drct_name":         meta.get("drct_name", str(drct_cd)),
                    "traffic_volume":    vol,
                    "anomaly_type":      None,
                    "corrected_value":   None,
                    "correction_method": None,
                    "confidence":        None,
                    "baseline":          None,
                })

            for r in node_results:
                position = (r["_acsr_id"], r["_date"], r["_hour"])
                if position in seen_positions:
                    continue
                drct_slots.append(_serialize_drct_slot(r, norm_node_id, baselines, drct_meta))

        filtered_drct_slots = [
            slot for slot in drct_slots
            if _normalize_drct_cd(slot.get("drct_cd")) != "00"
        ]
        slots = _aggregate_acsr_slots_from_drct(filtered_drct_slots)
        return {"slots": slots, "drct_slots": filtered_drct_slots}
    finally:
        conn.close()


def _fetch_corrected_traffic_acsr(req: JobRequest) -> dict:
    drct_result = _fetch_corrected_traffic_drct(req)
    return {"slots": drct_result["slots"]}


def _fetch_corrected_traffic_crsrd(req: JobRequest) -> dict:
    acsr_result = _fetch_corrected_traffic_acsr(req)
    slots = _aggregate_crsrd_slots_from_acsr(acsr_result["slots"])
    return {"slots": slots}


def _fetch_raw_traffic(req: JobRequest) -> dict:
    """Return raw DRCT rows plus approach and node aggregates."""
    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(req.date_start)
        date_end = date.fromisoformat(req.date_end)
        hours = req.resolve_hours()

        id_to_name = _get_id_to_name_map(conn)
        drct_code_map = ad.load_drct_code_map(conn)

        drct_slots: list[dict] = []
        for node_id in req.node_ids:
            norm_node_id = _normalize_node_id(node_id)
            node_name = id_to_name.get(norm_node_id, str(node_id))
            _, drct_meta = ad.load_drct_approaches(
                conn,
                norm_node_id,
                drct_code_map=drct_code_map,
            )
            target_data = ad.load_drct_target_data(
                conn,
                norm_node_id,
                date_start,
                date_end,
                hours,
            )

            for (drct_key, d, h), vol in target_data.items():
                approach_id = None
                drct_cd = None
                if isinstance(drct_key, tuple) and len(drct_key) >= 1:
                    approach_id = drct_key[0]
                if isinstance(drct_key, tuple) and len(drct_key) >= 2:
                    drct_cd = _normalize_drct_cd(drct_key[1])

                meta_key = (approach_id, drct_cd)
                meta = drct_meta.get(meta_key, {})
                approach_id = meta.get("approach_id", approach_id)
                drct_cd = _normalize_drct_cd(meta.get("drct_cd", drct_cd))
                if drct_cd == "00":
                    continue

                drct_slots.append({
                    "date": d.isoformat(),
                    "hour": h,
                    "node_id": norm_node_id,
                    "node_name": node_name,
                    "approach_id": approach_id,
                    "approach_name": meta.get("approach_name", str(approach_id)),
                    "drct_cd": drct_cd,
                    "drct_name": meta.get(
                        "drct_name",
                        drct_code_map.get(drct_cd, str(drct_cd)),
                    ),
                    "traffic_volume": vol,
                })

        drct_slots.sort(
            key=lambda r: (
                r["date"],
                r["node_id"],
                r["hour"],
                r["approach_id"],
                r["drct_cd"],
            )
        )
        slots = _aggregate_raw_slots_from_drct(drct_slots)
        node_slots = _aggregate_raw_node_slots(slots)
        return {
            "drct_slots": drct_slots,
            "slots": slots,
            "node_slots": node_slots,
        }
    finally:
        conn.close()


def _load_vknd_code_map(conn) -> dict:
    """{vknd_cd: vknd_name} from M_CD_INF where GRP_CD='VHCL_ATTR_CD'."""
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT CD, CD_NM
            FROM M_CD_INF
            WHERE GRP_CD = 'VHCL_ATTR_CD'
            ORDER BY CD
            """
        )
        return {_normalize_vknd_cd(cd): cd_nm for cd, cd_nm in cur.fetchall()}


def _fetch_vehicle_kinds() -> list[dict]:
    """Return vehicle-kind code options for VKND API consumers."""
    conn = ad.connect_db()
    try:
        code_map = _load_vknd_code_map(conn)
        return [
            {"code": code, "name": name}
            for code, name in sorted(code_map.items(), key=lambda item: item[0])
        ]
    finally:
        conn.close()


def _load_raw_vknd_target_rows(
    conn,
    node_id,
    date_start: date,
    date_end: date,
    hours: list[int],
    interval: str,
    vknd_codes: list[str] | None = None,
) -> list[tuple]:
    """Return raw vehicle-kind rows from the selected interval table."""
    if not hours:
        return []
    table_name = {
        "15m": "S_CRSRD_VKND_TRF_15MI",
        "1h": "S_CRSRD_VKND_TRF_1HH",
    }[interval]
    ds_dt = datetime(date_start.year, date_start.month, date_start.day)
    de_next = datetime(date_end.year, date_end.month, date_end.day) + timedelta(days=1)

    bind_params = {"nid": node_id, "ds": ds_dt, "de_next": de_next}
    hour_clause = ""
    if set(hours) != set(range(24)):
        hour_names = []
        for idx, hour in enumerate(sorted(set(hours))):
            name = f"h{idx}"
            bind_params[name] = hour
            hour_names.append(f":{name}")
        hour_clause = f"AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({','.join(hour_names)})"

    vknd_clause = ""
    if vknd_codes:
        vknd_names = []
        for idx, code in enumerate(vknd_codes):
            name = f"vknd{idx}"
            bind_params[name] = code
            vknd_names.append(f":{name}")
        vknd_clause = (
            "AND TRIM(TO_CHAR(VKND_CD)) "
            f"IN ({','.join(vknd_names)})"
        )

    sql = f"""
        SELECT ACSR_ID, TRIM(TO_CHAR(VKND_CD)) AS VKND_CD, TOT_DT, TRF_QNTY
        FROM {table_name}
        WHERE NODE_ID = :nid
          AND TOT_DT >= :ds
          AND TOT_DT < :de_next
          {hour_clause}
          {vknd_clause}
        ORDER BY TOT_DT, ACSR_ID, VKND_CD
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, **bind_params)
        return cur.fetchall()


def _load_raw_drct_target_rows(
    conn,
    node_id,
    date_start: date,
    date_end: date,
    hours: list[int],
    interval: str,
    approach_ids: list | None = None,
    drct_codes: list[str] | None = None,
) -> list[tuple]:
    """Return raw direction rows from the selected interval table."""
    if not hours and interval != "1d":
        return []
    table_name = {
        "5m": "S_CRSRD_DRCT_TRF_5MI",
        "15m": "S_CRSRD_DRCT_TRF_15MI",
        "1h": "S_CRSRD_DRCT_TRF_1HH",
        "1d": "S_CRSRD_DRCT_TRF_1DD",
    }[interval]
    ds_dt = datetime(date_start.year, date_start.month, date_start.day)
    de_next = datetime(date_end.year, date_end.month, date_end.day) + timedelta(days=1)

    bind_params = {"nid": node_id, "ds": ds_dt, "de_next": de_next}
    hour_clause = ""
    if interval != "1d" and set(hours) != set(range(24)):
        hour_names = []
        for idx, hour in enumerate(sorted(set(hours))):
            name = f"h{idx}"
            bind_params[name] = hour
            hour_names.append(f":{name}")
        hour_clause = f"AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({','.join(hour_names)})"

    acsr_clause = ""
    if approach_ids:
        acsr_names = []
        for idx, approach_id in enumerate(approach_ids):
            name = f"acsr{idx}"
            bind_params[name] = approach_id
            acsr_names.append(f":{name}")
        acsr_clause = f"AND ACSR_ID IN ({','.join(acsr_names)})"

    drct_clause = ""
    if drct_codes:
        drct_names = []
        for idx, drct_cd in enumerate(drct_codes):
            name = f"drct{idx}"
            bind_params[name] = drct_cd
            drct_names.append(f":{name}")
        drct_clause = f"AND TRIM(TO_CHAR(DRCT_CD)) IN ({','.join(drct_names)})"

    sql = f"""
        SELECT ACSR_ID, TRIM(TO_CHAR(DRCT_CD)) AS DRCT_CD, TOT_DT, TRF_QNTY
        FROM {table_name}
        WHERE NODE_ID = :nid
          AND TOT_DT >= :ds
          AND TOT_DT < :de_next
          AND TRIM(TO_CHAR(DRCT_CD)) <> '00'
          {hour_clause}
          {acsr_clause}
          {drct_clause}
        ORDER BY TOT_DT, ACSR_ID, DRCT_CD
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, **bind_params)
        return cur.fetchall()


def _timestamp_parts(tot_dt, interval: str) -> tuple[str, str, int, int]:
    if isinstance(tot_dt, datetime):
        return tot_dt.isoformat(), tot_dt.date().isoformat(), tot_dt.hour, tot_dt.minute
    timestamp = datetime.combine(tot_dt, datetime.min.time()).isoformat()
    return timestamp, tot_dt.isoformat(), 0, 0


def _iter_raw_drct_requested_times(
    date_start: date,
    date_end: date,
    hours: list[int],
    interval: str,
) -> list[datetime]:
    current = datetime(date_start.year, date_start.month, date_start.day)
    end_next = datetime(date_end.year, date_end.month, date_end.day) + timedelta(days=1)
    if interval == "1d":
        rows = []
        while current < end_next:
            rows.append(current)
            current += timedelta(days=1)
        return rows

    step = {
        "5m": timedelta(minutes=5),
        "15m": timedelta(minutes=15),
        "1h": timedelta(hours=1),
    }[interval]
    hour_set = set(hours)
    rows = []
    while current < end_next:
        if current.hour in hour_set:
            rows.append(current)
        current += step
    return rows


def _positive_actual_drct_presence(
    cache: dict,
    node_id,
    rows: list[tuple],
) -> bool:
    changed = False
    checked_date = date.today().isoformat()
    for approach_id, drct_cd, _tot_dt, trf_qnty in rows:
        if trf_qnty is not None and trf_qnty > 0:
            entry = _get_drct_direction_presence(cache, node_id, approach_id, drct_cd)
            if not entry or entry.get("exists") is not True:
                _set_drct_direction_presence(
                    cache,
                    node_id,
                    approach_id,
                    drct_cd,
                    True,
                    checked_date,
                )
                changed = True
    return changed


def _load_vknd_15m_rows(
    conn,
    node_id,
    date_start: date,
    date_end: date,
    hours: list[int],
    vknd_codes: list[str] | None = None,
    acsr_ids: list | None = None,
) -> list[tuple]:
    """Return 15-minute VKND rows for an arbitrary date range."""
    if not hours:
        return []
    ds_dt = datetime(date_start.year, date_start.month, date_start.day)
    de_next = datetime(date_end.year, date_end.month, date_end.day) + timedelta(days=1)
    bind_params = {"nid": node_id, "ds": ds_dt, "de_next": de_next}

    hour_clause = ""
    if set(hours) != set(range(24)):
        hour_names = []
        for idx, hour in enumerate(sorted(set(hours))):
            name = f"h{idx}"
            bind_params[name] = hour
            hour_names.append(f":{name}")
        hour_clause = f"AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({','.join(hour_names)})"

    vknd_clause = ""
    if vknd_codes:
        vknd_names = []
        for idx, code in enumerate(vknd_codes):
            name = f"vknd{idx}"
            bind_params[name] = code
            vknd_names.append(f":{name}")
        vknd_clause = f"AND TRIM(TO_CHAR(VKND_CD)) IN ({','.join(vknd_names)})"

    acsr_clause = ""
    if acsr_ids:
        acsr_names = []
        for idx, acsr_id in enumerate(acsr_ids):
            name = f"acsr{idx}"
            bind_params[name] = acsr_id
            acsr_names.append(f":{name}")
        acsr_clause = f"AND ACSR_ID IN ({','.join(acsr_names)})"

    sql = f"""
        SELECT ACSR_ID, TRIM(TO_CHAR(VKND_CD)) AS VKND_CD, TOT_DT, TRF_QNTY
        FROM S_CRSRD_VKND_TRF_15MI
        WHERE NODE_ID = :nid
          AND TOT_DT >= :ds
          AND TOT_DT < :de_next
          {hour_clause}
          {vknd_clause}
          {acsr_clause}
        ORDER BY TOT_DT, ACSR_ID, VKND_CD
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, **bind_params)
        return cur.fetchall()


def _load_vknd_15m_baseline_rows(
    conn,
    node_id,
    years: list[int],
    hours: list[int],
    months: list[int] | None,
    vknd_codes: list[str] | None,
    acsr_ids: list | None,
) -> list[tuple]:
    rows: list[tuple] = []
    month_filter = set(months) if months else None
    for year in years:
        year_rows = _load_vknd_15m_rows(
            conn,
            node_id,
            date(year, 1, 1),
            date(year, 12, 31),
            hours,
            vknd_codes,
            acsr_ids,
        )
        if month_filter is None:
            rows.extend(year_rows)
        else:
            rows.extend(
                row for row in year_rows
                if (row[2].month if isinstance(row[2], datetime) else row[2].month) in month_filter
            )
    return rows


def _load_vknd_15m_adjacent_rows(
    conn,
    node_id,
    adj_periods: list[tuple[int, int]],
    hours: list[int],
    vknd_codes: list[str] | None,
    acsr_ids: list | None,
) -> list[tuple]:
    rows: list[tuple] = []
    for year, month in adj_periods:
        month_start = date(year, month, 1)
        if month == 12:
            month_end = date(year, 12, 31)
        else:
            month_end = date(year, month + 1, 1) - timedelta(days=1)
        rows.extend(
            _load_vknd_15m_rows(
                conn,
                node_id,
                month_start,
                month_end,
                hours,
                vknd_codes,
                acsr_ids,
            )
        )
    return rows


def _rows_to_vknd_target_data(rows: list[tuple]) -> dict:
    target_data = {}
    for acsr_id, vknd_cd, tot_dt, trf_qnty in rows:
        d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
        h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
        m = tot_dt.minute if isinstance(tot_dt, datetime) else 0
        target_data[(acsr_id, _normalize_vknd_cd(vknd_cd), d, h, m)] = trf_qnty
    return target_data


def _build_vknd_15m_baselines(
    raw_rows: list[tuple],
    baseline_years: list[int],
    holiday_dates: set,
    approaches: list[tuple],
    vknd_codes: list[str],
    date_start: date,
    date_end: date,
    hours: list[int],
    adj_month_rows: list[tuple] | None = None,
    fallback_rows: list[tuple] | None = None,
    expand_stages: bool = True,
) -> tuple[dict, dict]:
    needed_cells = set()
    cur_d = date_start
    while cur_d <= date_end:
        day_type = ad.get_day_type(cur_d, holiday_dates)
        month = cur_d.month
        for hour in hours:
            for minute in (0, 15, 30, 45):
                for acsr_id, _ in approaches:
                    for vknd_cd in vknd_codes:
                        needed_cells.add((acsr_id, vknd_cd, day_type, hour, minute, month))
        cur_d += timedelta(days=1)

    cell_values = defaultdict(list)
    cell_actual_dates = defaultdict(set)
    cell_zero_counts: dict[tuple, int] = defaultdict(int)
    all_month_values: dict = defaultdict(list)
    all_month_actual_dates: dict = defaultdict(set)

    for acsr_id, vknd_cd, tot_dt, trf_qnty in raw_rows:
        d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
        h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
        m = tot_dt.minute if isinstance(tot_dt, datetime) else 0
        norm_vknd_cd = _normalize_vknd_cd(vknd_cd)
        month = d.month
        day_type = ad.get_day_type(d, holiday_dates)
        cell = (acsr_id, norm_vknd_cd, day_type, h, m, month)
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
        target_months = sorted({month for *_prefix, month in needed_cells})
        needed_keys = {(a, v, dt, h, m) for a, v, dt, h, m, _ in needed_cells}
        for acsr_id, vknd_cd, tot_dt, trf_qnty in adj_month_rows:
            if trf_qnty is None:
                continue
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            m = tot_dt.minute if isinstance(tot_dt, datetime) else 0
            norm_vknd_cd = _normalize_vknd_cd(vknd_cd)
            day_type = ad.get_day_type(d, holiday_dates)
            if (acsr_id, norm_vknd_cd, day_type, h, m) not in needed_keys:
                continue
            nearest_month = min(target_months, key=lambda target_m: abs(target_m - d.month))
            cell = (acsr_id, norm_vknd_cd, day_type, h, m, nearest_month)
            if cell in needed_cells:
                cell_values[cell].append(trf_qnty)
                cell_adj_counts[cell] += 1

    cell_fallback_counts: dict[tuple, int] = defaultdict(int)
    if fallback_rows:
        fallback_cells = {
            cell for cell in needed_cells
            if len(cell_values.get(cell, [])) < ad.MIN_CLEAN_SAMPLES
        }
        fallback_by_key: dict[tuple, list] = defaultdict(list)
        for acsr_id, vknd_cd, tot_dt, trf_qnty in fallback_rows:
            if trf_qnty is None:
                continue
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            m = tot_dt.minute if isinstance(tot_dt, datetime) else 0
            day_type = ad.get_day_type(d, holiday_dates)
            fallback_by_key[(acsr_id, _normalize_vknd_cd(vknd_cd), day_type, h, m)].append((d, trf_qnty))

        for cell in fallback_cells:
            acsr_id, vknd_cd, day_type, hour, minute, _month = cell
            already_seen = cell_actual_dates.get(cell, set())
            for d, val in fallback_by_key.get((acsr_id, vknd_cd, day_type, hour, minute), []):
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
            day_type = ad.get_day_type(cur_d, holiday_dates)
            month = cur_d.month
            for h in range(24):
                for m in (0, 15, 30, 45):
                    expected_counts[(day_type, h, m, month)] += 1
            cur_d += timedelta(days=1)

    baselines = {}
    for cell in needed_cells:
        acsr_id, vknd_cd, day_type, hour, minute, month = cell
        exp_count = expected_counts.get((day_type, hour, minute, month), 0)
        if exp_count == 0:
            continue
        act_count = len(cell_actual_dates.get(cell, set()))
        zero_rate = (exp_count - act_count) / exp_count
        baselines[cell] = {
            "zero_rate": zero_rate,
            "stage0_zr": zero_rate,
            "stage1_zr": None,
            "stage2_zr": None,
            "stats": ad.compute_baseline(cell_values.get(cell, [])),
            "n_adj": cell_adj_counts.get(cell, 0),
            "n_fallback": cell_fallback_counts.get(cell, 0),
            "expansion_stage": 0,
            "exp_count": exp_count,
            "act_count": act_count,
            "n_zero_baseline": cell_zero_counts.get(cell, 0),
            "n_total_baseline": len(cell_values.get(cell, [])),
        }

    if expand_stages:
        for cell, bl in baselines.items():
            if bl["zero_rate"] < ad.ZERO_RATE_EXPANSION_THRESHOLD:
                continue
            acsr_id, vknd_cd, day_type, hour, minute, month = cell
            prev_m = ((month - 2) % 12) + 1
            next_m = (month % 12) + 1
            stage1_months = [prev_m, month, next_m]
            exp_1 = sum(expected_counts.get((day_type, hour, minute, m), 0) for m in stage1_months)
            if exp_1 > 0:
                act_dates_1: set = set()
                vals_1 = []
                for m in stage1_months:
                    src_cell = (acsr_id, vknd_cd, day_type, hour, minute, m)
                    act_dates_1 |= all_month_actual_dates.get(src_cell, set())
                    vals_1.extend(all_month_values.get(src_cell, []))
                zr_1 = (exp_1 - len(act_dates_1)) / exp_1
                bl["zero_rate"] = zr_1
                bl["stage1_zr"] = zr_1
                bl["stats"] = ad.compute_baseline(vals_1)
                bl["expansion_stage"] = 1
                if zr_1 < ad.ZERO_RATE_EXPANSION_THRESHOLD and bl["stats"] is not None:
                    continue

            exp_2 = sum(expected_counts.get((day_type, hour, minute, m), 0) for m in range(1, 13))
            if exp_2 > 0:
                act_dates_2: set = set()
                vals_2 = []
                for m in range(1, 13):
                    src_cell = (acsr_id, vknd_cd, day_type, hour, minute, m)
                    act_dates_2 |= all_month_actual_dates.get(src_cell, set())
                    vals_2.extend(all_month_values.get(src_cell, []))
                zr_2 = (exp_2 - len(act_dates_2)) / exp_2
                bl["zero_rate"] = zr_2
                bl["stage2_zr"] = zr_2
                bl["stats"] = ad.compute_baseline(vals_2)
                bl["expansion_stage"] = 2

    fallback_applied = {c for c, v in cell_fallback_counts.items() if v > 0}
    return baselines, {
        "n_cells": len(fallback_applied),
        "n_values": sum(cell_fallback_counts.values()),
        "n_stage1": sum(1 for bl in baselines.values() if bl["expansion_stage"] >= 1),
        "n_stage2": sum(1 for bl in baselines.values() if bl["expansion_stage"] == 2),
    }


def _detect_vknd_15m_anomalies(
    node_name: str,
    approaches: list[tuple],
    vknd_codes: list[str],
    vknd_code_map: dict,
    baselines: dict,
    target_data: dict,
    date_start: date,
    date_end: date,
    hours: list[int],
    holiday_dates: set,
) -> list[dict]:
    results = []
    approach_names = dict(approaches)
    cur_d = date_start
    while cur_d <= date_end:
        day_type = ad.get_day_type(cur_d, holiday_dates)
        month = cur_d.month
        for hour in hours:
            for minute in (0, 15, 30, 45):
                for acsr_id, acsr_nm in approaches:
                    for vknd_cd in vknd_codes:
                        cell = (acsr_id, vknd_cd, day_type, hour, minute, month)
                        bl = baselines.get(cell)
                        if bl is None:
                            continue
                        key = (acsr_id, vknd_cd, cur_d, hour, minute)
                        has_record = key in target_data
                        trf_qnty = target_data.get(key)
                        anomaly_type = None
                        trf_val = 0

                        if not has_record:
                            if bl["zero_rate"] < ad.ZERO_RATE_THRESHOLD:
                                anomaly_type = "A형"
                            elif bl["zero_rate"] >= ad.ZERO_RATE_EXPANSION_THRESHOLD:
                                anomaly_type = "정상"
                            else:
                                anomaly_type = "B형"
                        elif bl["stats"] is not None:
                            median = bl["stats"]["median"]
                            q1 = bl["stats"]["q1"]
                            q3 = bl["stats"]["q3"]
                            iqr = q3 - q1
                            lower_iqr = q1 - ad.IQR_MULTIPLIER * iqr
                            ratio = (trf_qnty / median) if median > 0 else 1.0
                            if ratio < ad.RATIO_THRESHOLD or trf_qnty < lower_iqr:
                                anomaly_type = "B형"
                                trf_val = trf_qnty

                        if anomaly_type:
                            results.append({
                                "날짜": cur_d.strftime("%Y.%m.%d"),
                                "시간": f"{hour:02d}:{minute:02d}",
                                "교차로": node_name,
                                "방향": acsr_nm or approach_names.get(acsr_id, str(acsr_id)),
                                "차종": vknd_code_map.get(vknd_cd, vknd_cd),
                                "교통량": trf_val,
                                "판정": anomaly_type,
                                "_acsr_id": acsr_id,
                                "_vknd_cd": vknd_cd,
                                "_date": cur_d,
                                "_hour": hour,
                                "_minute": minute,
                                "_cell": cell,
                            })
        cur_d += timedelta(days=1)
    return results


def _apply_vknd_15m_corrections(results: list[dict], baselines: dict, target_data: dict) -> None:
    anomaly_set = {
        (r["_acsr_id"], r["_vknd_cd"], r["_date"], r["_hour"], r["_minute"])
        for r in results
    }
    for r in results:
        if r["판정"] == "정상":
            r["보정값"] = None
            r["보정방법"] = None
            r["신뢰도"] = None
            continue

        acsr_id = r["_acsr_id"]
        vknd_cd = r["_vknd_cd"]
        cur_d = r["_date"]
        hour = r["_hour"]
        minute = r["_minute"]
        bl = baselines.get(r["_cell"])
        stats = bl["stats"] if bl else None

        prev_val, d_before = None, None
        for dist in range(1, ad.MAX_INTERP_DAYS + 1):
            key = (acsr_id, vknd_cd, cur_d - timedelta(days=dist), hour, minute)
            if key in anomaly_set:
                continue
            if key in target_data:
                prev_val = target_data[key]
                d_before = dist
                break

        next_val, d_after = None, None
        for dist in range(1, ad.MAX_INTERP_DAYS + 1):
            key = (acsr_id, vknd_cd, cur_d + timedelta(days=dist), hour, minute)
            if key in anomaly_set:
                continue
            if key in target_data:
                next_val = target_data[key]
                d_after = dist
                break

        corrected, method, interp_dist = None, None, None
        if prev_val is not None and next_val is not None and d_before + d_after <= ad.MAX_INTERP_DAYS:
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
                (stats and stats["n_clean"] < ad.MIN_CONFIDENCE_SAMPLES)
                or (bl and bl.get("n_fallback", 0) > 0)
                or (interp_dist is not None and interp_dist > ad.MAX_INTERP_DAYS_OK)
            )
            r["신뢰도"] = "LOW" if is_low else "OK"
        else:
            r["보정값"] = None
            r["보정방법"] = None
            r["신뢰도"] = None


def _serialize_vknd_slot(
    r: dict,
    node_id,
    node_name: str,
    approach_names: dict,
    vknd_code_map: dict,
    baselines: dict,
    interval: str = "15m",
) -> dict:
    cell = r["_cell"]
    bl = baselines.get(cell, {})
    stats = bl.get("stats")
    d = r["_date"]
    h = r["_hour"]
    m = r["_minute"]
    timestamp = datetime(d.year, d.month, d.day, h, m).isoformat()
    vknd_cd = r["_vknd_cd"]
    return {
        "interval": interval,
        "timestamp": timestamp,
        "date": d.isoformat(),
        "hour": h,
        "minute": m,
        "node_id": node_id,
        "node_name": node_name,
        "approach_id": r["_acsr_id"],
        "approach_name": approach_names.get(r["_acsr_id"], r.get("방향", str(r["_acsr_id"]))),
        "vknd_cd": vknd_cd,
        "vknd_name": vknd_code_map.get(vknd_cd, vknd_cd),
        "traffic_volume": r["교통량"],
        "anomaly_type": r["판정"],
        "corrected_value": r.get("보정값"),
        "correction_method": r.get("보정방법"),
        "confidence": r.get("신뢰도"),
        "baseline": {
            "median": stats["median"] if stats else None,
            "q1": stats["q1"] if stats else None,
            "q3": stats["q3"] if stats else None,
            "n_clean": stats["n_clean"] if stats else None,
            "zero_rate": bl.get("zero_rate"),
            "expansion_stage": bl.get("expansion_stage", 0),
        },
    }


def _make_vknd_raw_slot(
    node_id,
    node_name: str,
    approach_id,
    approach_name: str,
    vknd_cd: str,
    vknd_name: str,
    d: date,
    hour: int,
    minute: int,
    volume,
    interval: str = "15m",
) -> dict:
    return {
        "interval": interval,
        "timestamp": datetime(d.year, d.month, d.day, hour, minute).isoformat(),
        "date": d.isoformat(),
        "hour": hour,
        "minute": minute,
        "node_id": node_id,
        "node_name": node_name,
        "approach_id": approach_id,
        "approach_name": approach_name,
        "vknd_cd": vknd_cd,
        "vknd_name": vknd_name,
        "traffic_volume": volume,
        "anomaly_type": None,
        "corrected_value": None,
        "correction_method": None,
        "confidence": None,
        "baseline": None,
    }


def _aggregate_vknd_15m_to_1h(slots_15m: list[dict]) -> list[dict]:
    agg: dict[tuple, dict] = {}
    for slot in slots_15m:
        key = (
            slot["date"],
            slot["hour"],
            slot["node_id"],
            slot["node_name"],
            slot["approach_id"],
            slot["approach_name"],
            slot["vknd_cd"],
            slot["vknd_name"],
        )
        if key not in agg:
            timestamp = datetime.fromisoformat(slot["timestamp"]).replace(minute=0).isoformat()
            agg[key] = {
                "interval": "1h",
                "timestamp": timestamp,
                "date": slot["date"],
                "hour": slot["hour"],
                "minute": None,
                "node_id": slot["node_id"],
                "node_name": slot["node_name"],
                "approach_id": slot["approach_id"],
                "approach_name": slot["approach_name"],
                "vknd_cd": slot["vknd_cd"],
                "vknd_name": slot["vknd_name"],
                "traffic_volume": None,
                "corrected_value": None,
                "_has_a": False,
                "_has_b": False,
                "_methods": set(),
                "_conf_low": False,
                "_has_baseline": False,
            }
        acc = agg[key]
        acc["traffic_volume"] = _sum_optional(acc["traffic_volume"], slot.get("traffic_volume"))
        corr_val = slot.get("corrected_value")
        if corr_val is None:
            corr_val = slot.get("traffic_volume")
        acc["corrected_value"] = _sum_optional(acc["corrected_value"], corr_val)
        anomaly_type = slot.get("anomaly_type")
        if anomaly_type == "A형":
            acc["_has_a"] = True
        elif anomaly_type == "B형":
            acc["_has_b"] = True
        elif anomaly_type == "A+B혼합":
            acc["_has_a"] = True
            acc["_has_b"] = True
        if slot.get("correction_method"):
            acc["_methods"].add(slot["correction_method"])
        if slot.get("confidence") == "LOW":
            acc["_conf_low"] = True
        if slot.get("baseline"):
            acc["_has_baseline"] = True

    rows: list[dict] = []
    for item in agg.values():
        anomaly_type = None
        if item["_has_a"] and item["_has_b"]:
            anomaly_type = "A+B혼합"
        elif item["_has_a"]:
            anomaly_type = "A형"
        elif item["_has_b"]:
            anomaly_type = "B형"
        methods = sorted(item["_methods"])
        rows.append({
            "interval": "1h",
            "timestamp": item["timestamp"],
            "date": item["date"],
            "hour": item["hour"],
            "minute": None,
            "node_id": item["node_id"],
            "node_name": item["node_name"],
            "approach_id": item["approach_id"],
            "approach_name": item["approach_name"],
            "vknd_cd": item["vknd_cd"],
            "vknd_name": item["vknd_name"],
            "traffic_volume": item["traffic_volume"],
            "anomaly_type": anomaly_type,
            "corrected_value": item["corrected_value"],
            "correction_method": "+".join(methods) if methods else None,
            "confidence": "LOW" if item["_conf_low"] else ("OK" if anomaly_type else None),
            "baseline": None,
        })
    rows.sort(key=lambda r: (r["timestamp"], r["node_id"], r["approach_id"], r["vknd_cd"]))
    return rows


def _analyse_node_vknd_15m(
    conn,
    node_id,
    node_name: str,
    date_start: date,
    date_end: date,
    hours: list[int],
    baseline_years: list[int],
    fallback_year: int,
    adj_periods: list[tuple[int, int]],
    holiday_dates: set,
    vknd_codes: list[str] | None,
    vknd_code_map: dict,
) -> tuple[list[dict], dict, dict, list[tuple], list[str]]:
    approaches = ad.load_approaches(conn, node_id)
    if not approaches:
        return [], {}, {}, [], []

    approach_ids = [acsr_id for acsr_id, _ in approaches]
    target_months: set[int] = set()
    cur_m = date_start.replace(day=1)
    while cur_m <= date_end:
        target_months.add(cur_m.month)
        cur_m = date(cur_m.year + 1, 1, 1) if cur_m.month == 12 else date(cur_m.year, cur_m.month + 1, 1)
    stage1_months: set[int] = set(target_months)
    for month in target_months:
        stage1_months.add(((month - 2) % 12) + 1)
        stage1_months.add((month % 12) + 1)

    raw_rows_core = _load_vknd_15m_baseline_rows(
        conn,
        node_id,
        baseline_years,
        hours,
        sorted(stage1_months),
        vknd_codes,
        approach_ids,
    )
    target_rows = _load_vknd_15m_rows(
        conn,
        node_id,
        date_start,
        date_end,
        hours,
        vknd_codes,
        approach_ids,
    )

    resolved_vknd_codes = (
        sorted({_normalize_vknd_cd(code) for code in vknd_codes if _normalize_vknd_cd(code)})
        if vknd_codes
        else sorted({
            _normalize_vknd_cd(row[1])
            for row in raw_rows_core + target_rows
            if _normalize_vknd_cd(row[1])
        } or set(vknd_code_map.keys()))
    )
    if not resolved_vknd_codes:
        return [], {}, _rows_to_vknd_target_data(target_rows), approaches, []

    adj_rows = _load_vknd_15m_adjacent_rows(
        conn,
        node_id,
        adj_periods,
        hours,
        resolved_vknd_codes,
        approach_ids,
    )
    baselines_pre, _ = _build_vknd_15m_baselines(
        raw_rows_core,
        baseline_years,
        holiday_dates,
        approaches,
        resolved_vknd_codes,
        date_start,
        date_end,
        hours,
        adj_month_rows=adj_rows,
        expand_stages=False,
    )

    stage2_keys = {
        (acsr_id, vknd_cd, day_type, hour, minute)
        for (acsr_id, vknd_cd, day_type, hour, minute, _), bl in baselines_pre.items()
        if bl["zero_rate"] >= ad.ZERO_RATE_EXPANSION_THRESHOLD
    }
    raw_rows_stage2: list[tuple] = []
    if stage2_keys:
        stage2_acsr_ids = sorted({acsr_id for acsr_id, *_ in stage2_keys})
        stage2_vknd_codes = sorted({vknd_cd for _, vknd_cd, *_ in stage2_keys})
        annual_rows = _load_vknd_15m_baseline_rows(
            conn,
            node_id,
            baseline_years,
            hours,
            None,
            stage2_vknd_codes,
            stage2_acsr_ids,
        )
        for acsr_id, vknd_cd, tot_dt, trf_qnty in annual_rows:
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            m = tot_dt.minute if isinstance(tot_dt, datetime) else 0
            day_type = ad.get_day_type(d, holiday_dates)
            if (acsr_id, _normalize_vknd_cd(vknd_cd), day_type, h, m) in stage2_keys:
                raw_rows_stage2.append((acsr_id, vknd_cd, tot_dt, trf_qnty))

    seen_positions: set[tuple] = set()
    raw_rows: list[tuple] = []
    for source_rows in (raw_rows_core, raw_rows_stage2):
        for acsr_id, vknd_cd, tot_dt, trf_qnty in source_rows:
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            m = tot_dt.minute if isinstance(tot_dt, datetime) else 0
            key = (acsr_id, _normalize_vknd_cd(vknd_cd), d, h, m)
            if key in seen_positions:
                continue
            seen_positions.add(key)
            raw_rows.append((acsr_id, vknd_cd, tot_dt, trf_qnty))

    fallback_keys = {
        (acsr_id, vknd_cd, day_type, hour, minute)
        for (acsr_id, vknd_cd, day_type, hour, minute, _), bl in baselines_pre.items()
        if bl.get("n_total_baseline", 0) < ad.MIN_CLEAN_SAMPLES
    }
    fallback_raw_rows: list[tuple] = []
    if fallback_keys:
        fallback_acsr_ids = sorted({acsr_id for acsr_id, *_ in fallback_keys})
        fallback_vknd_codes = sorted({vknd_cd for _, vknd_cd, *_ in fallback_keys})
        fallback_rows_all = _load_vknd_15m_baseline_rows(
            conn,
            node_id,
            [fallback_year],
            hours,
            None,
            fallback_vknd_codes,
            fallback_acsr_ids,
        )
        for acsr_id, vknd_cd, tot_dt, trf_qnty in fallback_rows_all:
            if trf_qnty is None:
                continue
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            m = tot_dt.minute if isinstance(tot_dt, datetime) else 0
            day_type = ad.get_day_type(d, holiday_dates)
            if (acsr_id, _normalize_vknd_cd(vknd_cd), day_type, h, m) in fallback_keys:
                fallback_raw_rows.append((acsr_id, vknd_cd, tot_dt, trf_qnty))

    baselines, _ = _build_vknd_15m_baselines(
        raw_rows,
        baseline_years,
        holiday_dates,
        approaches,
        resolved_vknd_codes,
        date_start,
        date_end,
        hours,
        adj_month_rows=adj_rows,
        fallback_rows=fallback_raw_rows,
    )
    target_data = _rows_to_vknd_target_data(target_rows)
    node_results = _detect_vknd_15m_anomalies(
        node_name,
        approaches,
        resolved_vknd_codes,
        vknd_code_map,
        baselines,
        target_data,
        date_start,
        date_end,
        hours,
        holiday_dates,
    )
    _apply_vknd_15m_corrections(node_results, baselines, target_data)
    return node_results, baselines, target_data, approaches, resolved_vknd_codes


def _fetch_raw_traffic_vknd(req: RawTrafficVkndRequest) -> dict:
    """Return raw vehicle-kind traffic plus approach and node aggregates."""
    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(req.date_start)
        date_end = date.fromisoformat(req.date_end)
        hours = req.resolve_hours()
        interval = req.interval
        vknd_codes = (
            sorted({_normalize_vknd_cd(code) for code in req.vknd_codes if _normalize_vknd_cd(code)})
            if req.vknd_codes
            else None
        )

        id_to_name = _get_id_to_name_map(conn)
        vknd_code_map = _load_vknd_code_map(conn)

        vknd_slots: list[dict] = []
        for node_id in req.node_ids:
            norm_node_id = _normalize_node_id(node_id)
            node_name = id_to_name.get(norm_node_id, str(node_id))
            approach_names = {
                approach_id: approach_name
                for approach_id, approach_name in ad.load_approaches(conn, norm_node_id)
            }
            rows = _load_raw_vknd_target_rows(
                conn,
                norm_node_id,
                date_start,
                date_end,
                hours,
                interval,
                vknd_codes,
            )

            for approach_id, vknd_cd, tot_dt, trf_qnty in rows:
                norm_vknd_cd = _normalize_vknd_cd(vknd_cd)
                if isinstance(tot_dt, datetime):
                    timestamp = tot_dt.isoformat()
                    slot_date = tot_dt.date().isoformat()
                    hour = tot_dt.hour
                    minute = tot_dt.minute
                else:
                    timestamp = datetime.combine(tot_dt, datetime.min.time()).isoformat()
                    slot_date = tot_dt.isoformat()
                    hour = 0
                    minute = 0
                vknd_slots.append({
                    "interval": interval,
                    "timestamp": timestamp,
                    "date": slot_date,
                    "hour": hour,
                    "minute": minute,
                    "node_id": norm_node_id,
                    "node_name": node_name,
                    "approach_id": approach_id,
                    "approach_name": approach_names.get(approach_id, str(approach_id)),
                    "vknd_cd": norm_vknd_cd,
                    "vknd_name": vknd_code_map.get(norm_vknd_cd, norm_vknd_cd),
                    "traffic_volume": trf_qnty,
                })

        vknd_slots.sort(
            key=lambda r: (
                r["timestamp"],
                r["node_id"],
                r["approach_id"],
                r["vknd_cd"],
            )
        )
        slots = _aggregate_raw_slots_from_vknd(vknd_slots)
        node_slots = _aggregate_raw_node_slots_from_vknd(slots)
        return {
            "vknd_slots": vknd_slots,
            "slots": slots,
            "node_slots": node_slots,
        }
    finally:
        conn.close()


def _fetch_raw_traffic_drct(req: RawTrafficDrctRequest) -> dict:
    """Return raw direction traffic, including cached missing-direction rows."""
    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(req.date_start)
        date_end = date.fromisoformat(req.date_end)
        hours = req.resolve_hours()
        interval = req.interval
        requested_approaches = (
            sorted({_normalize_acsr_id(v) for v in req.approach_ids})
            if req.approach_ids
            else None
        )
        requested_drcts = (
            sorted({
                _normalize_drct_cd(code)
                for code in req.drct_codes
                if _normalize_drct_cd(code) and _normalize_drct_cd(code) != "00"
            })
            if req.drct_codes
            else None
        )

        id_to_name = _get_id_to_name_map(conn)
        drct_code_map = ad.load_drct_code_map(conn)
        all_drct_codes = sorted(
            code for code in (_normalize_drct_cd(c) for c in drct_code_map.keys())
            if code and code != "00"
        )

        requested_times = _iter_raw_drct_requested_times(date_start, date_end, hours, interval)
        drct_slots: list[dict] = []

        for node_id in req.node_ids:
            norm_node_id = _normalize_node_id(node_id)
            node_name = id_to_name.get(norm_node_id, str(node_id))
            _, drct_meta = ad.load_drct_approaches(
                conn,
                norm_node_id,
                drct_code_map=drct_code_map,
            )
            approach_names = {
                approach_id: meta.get("approach_name", str(approach_id))
                for (approach_id, _drct_cd), meta in drct_meta.items()
            }
            if not approach_names:
                approach_names = {
                    approach_id: approach_name
                    for approach_id, approach_name in ad.load_approaches(conn, norm_node_id)
                }

            approach_ids = requested_approaches or sorted(approach_names.keys())
            rows = _load_raw_drct_target_rows(
                conn,
                norm_node_id,
                date_start,
                date_end,
                hours,
                interval,
                requested_approaches,
                requested_drcts,
            )
            actual_map = {}
            actual_drcts_by_approach: dict = defaultdict(set)
            for approach_id, drct_cd, tot_dt, trf_qnty in rows:
                norm_approach_id = _normalize_acsr_id(approach_id)
                norm_drct_cd = _normalize_drct_cd(drct_cd)
                if norm_drct_cd == "00":
                    continue
                timestamp, slot_date, hour, minute = _timestamp_parts(tot_dt, interval)
                actual_map[(timestamp, norm_approach_id, norm_drct_cd)] = (
                    slot_date,
                    hour,
                    minute,
                    trf_qnty,
                )
                actual_drcts_by_approach[norm_approach_id].add(norm_drct_cd)
                if norm_approach_id not in approach_names:
                    approach_names[norm_approach_id] = str(norm_approach_id)

            if not approach_ids:
                approach_ids = sorted(approach_names.keys())
            drct_codes_for_refresh = requested_drcts or sorted(
                set(all_drct_codes)
                | {drct_cd for codes in actual_drcts_by_approach.values() for drct_cd in codes}
            )
            cache = _ensure_drct_direction_presence_entries(
                conn,
                norm_node_id,
                approach_ids,
                drct_codes_for_refresh,
            )
            if _positive_actual_drct_presence(cache, norm_node_id, rows):
                with _drct_direction_presence_lock:
                    _save_drct_direction_presence_cache(cache)

            for ts in requested_times:
                timestamp = ts.isoformat()
                slot_date = ts.date().isoformat()
                hour = ts.hour
                minute = ts.minute
                for approach_id in approach_ids:
                    if requested_drcts:
                        candidate_drcts = requested_drcts
                    else:
                        cached_dirs = (
                            cache.get("nodes", {})
                            .get(str(norm_node_id), {})
                            .get("approaches", {})
                            .get(str(approach_id), {})
                            .get("directions", {})
                        )
                        candidate_drcts = sorted(
                            set(actual_drcts_by_approach.get(approach_id, set()))
                            | {
                                code for code, entry in cached_dirs.items()
                                if entry.get("exists") and _normalize_drct_cd(code) != "00"
                            }
                        )
                    for drct_cd in candidate_drcts:
                        actual = actual_map.get((timestamp, approach_id, drct_cd))
                        if actual:
                            row_date, row_hour, row_minute, volume = actual
                            traffic_volume = volume
                            is_collected = True
                        else:
                            row_date, row_hour, row_minute = slot_date, hour, minute
                            entry = _get_drct_direction_presence(cache, norm_node_id, approach_id, drct_cd)
                            traffic_volume = 0 if entry and entry.get("exists") else None
                            is_collected = False
                        meta = drct_meta.get((approach_id, drct_cd), {})
                        drct_slots.append({
                            "interval": interval,
                            "timestamp": timestamp,
                            "date": row_date,
                            "hour": row_hour,
                            "minute": row_minute,
                            "node_id": norm_node_id,
                            "node_name": node_name,
                            "approach_id": approach_id,
                            "approach_name": meta.get(
                                "approach_name",
                                approach_names.get(approach_id, str(approach_id)),
                            ),
                            "drct_cd": drct_cd,
                            "drct_name": meta.get(
                                "drct_name",
                                drct_code_map.get(drct_cd, drct_cd),
                            ),
                            "traffic_volume": traffic_volume,
                            "is_collected": is_collected,
                        })

        drct_slots.sort(
            key=lambda r: (
                r["timestamp"],
                r["node_id"],
                r["approach_id"],
                r["drct_cd"],
            )
        )
        slots = _aggregate_raw_slots_from_drct(drct_slots)
        node_slots = _aggregate_raw_node_slots(slots)
        return {
            "drct_slots": drct_slots,
            "slots": slots,
            "node_slots": node_slots,
        }
    finally:
        conn.close()


def _fetch_corrected_traffic_vknd(req: RawTrafficVkndRequest) -> dict:
    """Return VKND corrected traffic using 15-minute correction as the source of truth."""
    if _holiday_dates is None:
        raise RuntimeError("서버 초기화가 완료되지 않았습니다. 잠시 후 다시 시도하세요.")

    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(req.date_start)
        date_end = date.fromisoformat(req.date_end)
        hours = req.resolve_hours()
        interval = req.interval
        vknd_codes = (
            sorted({_normalize_vknd_cd(code) for code in req.vknd_codes if _normalize_vknd_cd(code)})
            if req.vknd_codes
            else None
        )

        baseline_years = ad.select_baseline_years(date_start.year)
        fallback_year = ad.select_fallback_year(date_start.year)
        adj_periods = ad.get_adjacent_month_periods(date_start, date_end)

        id_to_name = _get_id_to_name_map(conn)
        vknd_code_map = _load_vknd_code_map(conn)

        vknd_slots_15m: list[dict] = []
        for node_id in req.node_ids:
            norm_node_id = _normalize_node_id(node_id)
            node_name = id_to_name.get(norm_node_id, str(node_id))
            node_results, baselines, target_data, approaches, _resolved_vknd_codes = _analyse_node_vknd_15m(
                conn,
                norm_node_id,
                node_name,
                date_start,
                date_end,
                hours,
                baseline_years,
                fallback_year,
                adj_periods,
                _holiday_dates,
                vknd_codes,
                vknd_code_map,
            )
            approach_names = {acsr_id: acsr_nm for acsr_id, acsr_nm in approaches}
            anomaly_index = {
                (r["_acsr_id"], r["_vknd_cd"], r["_date"], r["_hour"], r["_minute"]): r
                for r in node_results
            }
            seen_positions: set[tuple] = set()

            for (acsr_id, vknd_cd, d, h, m), vol in target_data.items():
                position = (acsr_id, vknd_cd, d, h, m)
                seen_positions.add(position)
                if position in anomaly_index:
                    vknd_slots_15m.append(
                        _serialize_vknd_slot(
                            anomaly_index[position],
                            norm_node_id,
                            node_name,
                            approach_names,
                            vknd_code_map,
                            baselines,
                        )
                    )
                    continue
                vknd_slots_15m.append(
                    _make_vknd_raw_slot(
                        norm_node_id,
                        node_name,
                        acsr_id,
                        approach_names.get(acsr_id, str(acsr_id)),
                        vknd_cd,
                        vknd_code_map.get(vknd_cd, vknd_cd),
                        d,
                        h,
                        m,
                        vol,
                    )
                )

            for r in node_results:
                position = (r["_acsr_id"], r["_vknd_cd"], r["_date"], r["_hour"], r["_minute"])
                if position in seen_positions:
                    continue
                if r.get("판정") == "정상" and r.get("보정값") is None:
                    continue
                vknd_slots_15m.append(
                    _serialize_vknd_slot(
                        r,
                        norm_node_id,
                        node_name,
                        approach_names,
                        vknd_code_map,
                        baselines,
                    )
                )

        vknd_slots_15m.sort(
            key=lambda r: (
                r["timestamp"],
                r["node_id"],
                r["approach_id"],
                r["vknd_cd"],
            )
        )
        if interval == "1h":
            vknd_slots = _aggregate_vknd_15m_to_1h(vknd_slots_15m)
        else:
            vknd_slots = vknd_slots_15m
        slots = _aggregate_corrected_slots_from_vknd(vknd_slots)
        node_slots = _aggregate_corrected_node_slots_from_vknd(slots)
        return {
            "vknd_slots": vknd_slots,
            "slots": slots,
            "node_slots": node_slots,
        }
    finally:
        conn.close()


def _aggregate_daily_summary_rows(
    node_id: int,
    node_name: str,
    node_results: list,
    target_data: dict,
) -> list[dict]:
    """anomaly 판정 결과 + 타겟 슬롯으로 일별 집계 행 생성."""
    all_slot_index: dict[tuple[date, int], set[int]] = defaultdict(set)
    for acsr_id, d, h in target_data.keys():
        all_slot_index[(d, h)].add(acsr_id)

    anomaly_index: dict[tuple[date, int], dict[int, dict]] = defaultdict(dict)
    for r in node_results:
        if r.get("판정") == "A형":
            anomaly_index[(r["_date"], r["_hour"])][r["_acsr_id"]] = r

    node_counter: dict[date, int] = defaultdict(int)
    approach_counter: dict[tuple[date, int], list] = {}

    for (d, h), expected in all_slot_index.items():
        if not expected:
            continue
        anomaly_map = anomaly_index.get((d, h), {})
        if not anomaly_map:
            continue
        if set(anomaly_map.keys()) == expected:
            node_counter[d] += 1
        else:
            for acsr_id, r in anomaly_map.items():
                key = (d, acsr_id)
                if key not in approach_counter:
                    approach_counter[key] = [r.get("방향"), 0]
                approach_counter[key][1] += 1

    rows: list[dict] = []
    for d, cnt in node_counter.items():
        rows.append({
            "date": d.isoformat(),
            "node_id": node_id,
            "node_name": node_name,
            "approach_id": None,
            "approach_name": None,
            "missing_count": cnt,
        })
    for (d, acsr_id), (name, cnt) in approach_counter.items():
        rows.append({
            "date": d.isoformat(),
            "node_id": node_id,
            "node_name": node_name,
            "approach_id": acsr_id,
            "approach_name": name,
            "missing_count": cnt,
        })

    rows.sort(key=lambda r: (r["date"], r["node_id"], r["approach_id"] is not None, r["approach_id"] or -1))
    return rows


def _fetch_anomaly_daily_summary(req: JobRequest) -> dict:
    """일별집계 전용 응답을 동기 계산 (executor에서 실행)."""
    if _holiday_dates is None:
        raise RuntimeError("서버 초기화가 완료되지 않았습니다. 잠시 후 다시 시도하세요.")

    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(req.date_start)
        date_end = date.fromisoformat(req.date_end)
        hours = req.resolve_hours()

        baseline_years = ad.select_baseline_years(date_start.year)
        fallback_year = ad.select_fallback_year(date_start.year)
        adj_periods = ad.get_adjacent_month_periods(date_start, date_end)

        id_to_name = _get_id_to_name_map(conn)

        all_rows: list[dict] = []
        for node_id in req.node_ids:
            norm_node_id = _normalize_node_id(node_id)
            node_name = id_to_name.get(norm_node_id, str(node_id))
            node_results, _, target_data = ad.analyse_node(
                conn, norm_node_id, node_name,
                date_start, date_end, hours,
                baseline_years, fallback_year, adj_periods,
                _holiday_dates,
            )
            all_rows.extend(
                _aggregate_daily_summary_rows(norm_node_id, node_name, node_results, target_data)
            )

        all_rows.sort(key=lambda r: (r["date"], r["node_id"], r["approach_id"] is not None, r["approach_id"] or -1))
        return {"rows": all_rows}
    finally:
        conn.close()


# ── 엔드포인트 ────────────────────────────────────────────────────────────────

@app.post("/jobs", status_code=201)
async def create_job(req: JobRequest):
    """분석 작업을 시작하고 job_id를 반환합니다.

    처리는 백그라운드에서 비동기로 진행됩니다.
    GET /jobs/{job_id} 로 상태와 결과를 조회하세요.
    """
    job_id = str(uuid.uuid4())
    now    = datetime.now()

    job: dict = {
        "job_id":      job_id,
        "status":      "pending",
        "request": {
            "node_ids":   req.node_ids,
            "date_start": req.date_start,
            "date_end":   req.date_end,
            "hours":      req.resolve_hours(),
        },
        "created_at":  now,
        "started_at":  None,
        "finished_at": None,
        "progress":    None,
        "result":      None,
        "error":       None,
    }

    with job_store_lock:
        job_store[job_id] = job

    loop = asyncio.get_running_loop()
    loop.run_in_executor(executor, _run_job, job_id)

    return {
        "job_id":     job_id,
        "status":     "pending",
        "created_at": now.isoformat(),
    }


@app.get("/jobs/{job_id}")
async def get_job(job_id: str):
    """작업 상태와 결과를 반환합니다.

    status 값:
    - pending  : 대기 중
    - running  : 처리 중 (progress 필드 참고)
    - done     : 완료 (result 필드 참고)
    - failed   : 실패 (error 필드 참고)
    """
    job = job_store.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"job_id '{job_id}'를 찾을 수 없습니다.")
    return _serialize_job(job)


@app.post("/corrected-traffic")
async def corrected_traffic(req: JobRequest):
    """이상탐지+보정 결과를 동기적으로 반환합니다.

    /jobs 와 동일한 요청 형식이지만, 결과를 즉시 반환합니다 (job 대기 없음).
    응답: {"slots": [...]}
    """
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(executor, _fetch_corrected_traffic, _job_request_payload(req))
    return result


@app.post("/corrected-traffic-acsr")
async def corrected_traffic_acsr(req: JobRequest):
    """Return DRCT-based corrected traffic aggregated to ACSR slots."""
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(executor, _fetch_corrected_traffic_acsr, req)
    return result


@app.post("/corrected-traffic-crsrd")
async def corrected_traffic_crsrd(req: JobRequest):
    """Return DRCT-based corrected traffic aggregated to CRSRD slots."""
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(executor, _fetch_corrected_traffic_crsrd, req)
    return result


@app.post("/corrected-traffic-drct")
async def corrected_traffic_drct(req: JobRequest):
    """DRCT 단위 결과와 ACSR 집계 결과를 동기적으로 반환합니다.

    응답:
    - slots: ACSR 집계 결과
    - drct_slots: DRCT 단위 상세 결과
    """
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(executor, _fetch_corrected_traffic_drct, req)
    return result


@app.post("/corrected-traffic-vknd")
async def corrected_traffic_vknd(req: RawTrafficVkndRequest):
    """Return vehicle-kind corrected traffic from 15-minute correction results."""
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(executor, _fetch_corrected_traffic_vknd, req)
    return result


@app.get("/vehicle-kinds")
async def get_vehicle_kinds():
    """차종 코드 목록을 반환합니다.

    /corrected-traffic-vknd 의 vknd_codes 필드에 사용할 수 있습니다.
    """
    loop = asyncio.get_running_loop()
    items = await loop.run_in_executor(executor, _fetch_vehicle_kinds)
    return {"count": len(items), "items": items}


@app.post("/anomaly-daily-summary")
async def anomaly_daily_summary(req: JobRequest):
    """일별집계 전용 결과를 즉시 반환합니다.

    응답: {"rows": [{"date","node_id","node_name","approach_id","approach_name","missing_count"}, ...]}
    """
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(executor, _fetch_anomaly_daily_summary, req)
    return result


@app.get("/intersections")
async def get_intersections(refresh: bool = False):
    """교차로 목록을 반환합니다.

    POST /jobs 의 node_ids 필드에 사용할 node_id 값을 확인할 수 있습니다.
    서버 시작 후 첫 호출 시 DB를 조회하고 이후에는 메모리 캐시를 사용합니다.
    refresh=true 를 전달하면 DB를 재조회합니다.
    """
    global _intersections_cache

    if _intersections_cache is None or refresh:
        loop = asyncio.get_running_loop()
        items = await loop.run_in_executor(executor, _fetch_intersections)
        _intersections_cache = items

    return {"count": len(_intersections_cache), "items": _intersections_cache}


# ── 실행 ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
