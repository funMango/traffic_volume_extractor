from __future__ import annotations

import asyncio

from fastapi import APIRouter, HTTPException, Request

from .schemas import JobRequest, RawTrafficDrctRequest, RawTrafficVkndRequest


router = APIRouter()


def _services(request: Request) -> dict:
    return request.app.state.services


@router.post("/jobs", status_code=201)
async def create_job(req: JobRequest, request: Request):
    return _services(request)["jobs"].create(req)


@router.get("/jobs/{job_id}")
async def get_job(job_id: str, request: Request):
    job = _services(request)["jobs"].get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"job_id '{job_id}'를 찾을 수 없습니다.")
    return job


@router.post("/corrected-traffic")
async def corrected_traffic(req: JobRequest, request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["traffic"].corrected_traffic,
        req,
    )


@router.post("/corrected-traffic-drct")
async def corrected_traffic_drct(req: JobRequest, request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["traffic"].corrected_traffic_drct,
        req,
    )


@router.post("/raw-traffic")
async def raw_traffic(req: JobRequest, request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["traffic"].raw_traffic,
        req,
    )


@router.post("/raw-traffic-drct")
async def raw_traffic_drct(req: RawTrafficDrctRequest, request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["traffic"].raw_traffic_drct,
        req,
    )


@router.post("/raw-traffic-vknd")
async def raw_traffic_vknd(req: RawTrafficVkndRequest, request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["traffic"].raw_traffic_vknd,
        req,
    )


@router.post("/corrected-traffic-vknd")
async def corrected_traffic_vknd(req: RawTrafficVkndRequest, request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["traffic"].corrected_traffic_vknd,
        req,
    )


@router.get("/vehicle-kinds")
async def get_vehicle_kinds(request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["reference"].vehicle_kinds,
    )


@router.post("/anomaly-daily-summary")
async def anomaly_daily_summary(req: JobRequest, request: Request):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["traffic"].anomaly_daily_summary,
        req,
    )


@router.get("/intersections")
async def get_intersections(request: Request, refresh: bool = False):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        None,
        _services(request)["reference"].intersections,
        refresh,
    )
