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
  GET  /intersections     교차로 목록 조회
"""

import asyncio
import sys
import uuid
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from datetime import date, datetime
from pathlib import Path
from threading import Lock
from typing import Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, model_validator

# 이상탐지 모듈 import (경로는 이상탐지.py의 __file__ 기준으로 동작)
sys.path.insert(0, str(Path(__file__).parent))
import 이상탐지 as ad

# ── 전역 상태 ─────────────────────────────────────────────────────────────────

executor = ThreadPoolExecutor(max_workers=4)

job_store: dict[str, dict] = {}
job_store_lock = Lock()

_intersections_cache: list[dict] | None = None
_holiday_dates: set | None = None


# ── 라이프사이클 ──────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    global _holiday_dates
    _holiday_dates = ad.load_holidays()
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


# ── 헬퍼: job 직렬화 ─────────────────────────────────────────────────────────

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


def _fetch_corrected_traffic(req: JobRequest) -> dict:
    """이상탐지+보정 결과를 동기적으로 계산 (executor에서 실행)

    정상 슬롯도 포함해 반환하므로 이상 없을 때도 원본 교통량 그래프를 그릴 수 있다.
    """
    if _holiday_dates is None:
        raise RuntimeError("서버 초기화가 완료되지 않았습니다. 잠시 후 다시 시도하세요.")

    conn = ad.connect_db()
    try:
        date_start = date.fromisoformat(req.date_start)
        date_end   = date.fromisoformat(req.date_end)
        hours      = req.resolve_hours()

        baseline_years = ad.select_baseline_years(date_start.year)
        fallback_year  = ad.select_fallback_year(date_start.year)
        adj_periods    = ad.get_adjacent_month_periods(date_start, date_end)

        id_to_name = {nid: nm for nid, nm in ad.load_intersections(conn)}

        all_slots: list[dict] = []
        for node_id in req.node_ids:
            node_name = id_to_name.get(node_id, str(node_id))
            node_results, baselines, target_data = ad.analyse_node(
                conn, node_id, node_name,
                date_start, date_end, hours,
                baseline_years, fallback_year, adj_periods,
                _holiday_dates,
            )

            # approach name lookup — str 정규화로 Oracle 타입 불일치 방지
            acsr_names: dict = {str(aid): anm for aid, anm in ad.load_approaches(conn, node_id)}
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
                        "node_id":           node_id,
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

        id_to_name = {nid: nm for nid, nm in ad.load_intersections(conn)}

        all_rows: list[dict] = []
        for node_id in req.node_ids:
            node_name = id_to_name.get(node_id, str(node_id))
            node_results, _, target_data = ad.analyse_node(
                conn, node_id, node_name,
                date_start, date_end, hours,
                baseline_years, fallback_year, adj_periods,
                _holiday_dates,
            )
            all_rows.extend(
                _aggregate_daily_summary_rows(node_id, node_name, node_results, target_data)
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
    result = await loop.run_in_executor(executor, _fetch_corrected_traffic, req)
    return result


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
