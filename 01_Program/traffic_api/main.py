from __future__ import annotations

from contextlib import asynccontextmanager

from fastapi import FastAPI

from traffic_api.application.job_use_cases import JobUseCase
from traffic_api.application.traffic_use_cases import ReferenceDataUseCase, TrafficQueryUseCase
from traffic_api.infrastructure.legacy_analysis_gateway import (
    LegacyReferenceDataGateway,
    LegacyTrafficAnalysisGateway,
)
from traffic_api.infrastructure.memory_job_store import InMemoryJobStore
from traffic_api.infrastructure.thread_pool_runner import ThreadPoolBackgroundJobRunner
from traffic_api.presentation.routes import router


def create_services() -> dict:
    traffic_gateway = LegacyTrafficAnalysisGateway()
    reference_gateway = LegacyReferenceDataGateway()
    job_store = InMemoryJobStore()
    runner = ThreadPoolBackgroundJobRunner(max_workers=4)
    return {
        "traffic_gateway": traffic_gateway,
        "runner": runner,
        "jobs": JobUseCase(job_store, runner, traffic_gateway),
        "traffic": TrafficQueryUseCase(traffic_gateway),
        "reference": ReferenceDataUseCase(reference_gateway),
    }


@asynccontextmanager
async def lifespan(app: FastAPI):
    app.state.services["traffic_gateway"].initialize()
    try:
        yield
    finally:
        app.state.services["runner"].shutdown()


def create_app() -> FastAPI:
    app = FastAPI(
        title="교통량 이상탐지 API",
        version="1.0.0",
        description="교통량 이상값 탐지·보정 결과를 JSON으로 제공합니다.",
        lifespan=lifespan,
    )
    app.state.services = create_services()
    app.include_router(router)
    return app


app = create_app()


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)

