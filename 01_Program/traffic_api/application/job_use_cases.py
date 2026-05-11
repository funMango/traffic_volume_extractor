from __future__ import annotations

import uuid
from datetime import datetime

from .interfaces import BackgroundJobRunner, JobStore, TrafficAnalysisGateway


class JobUseCase:
    def __init__(
        self,
        store: JobStore,
        runner: BackgroundJobRunner,
        gateway: TrafficAnalysisGateway,
    ) -> None:
        self._store = store
        self._runner = runner
        self._gateway = gateway

    def create(self, request) -> dict:
        job_id = str(uuid.uuid4())
        now = datetime.now()
        job = {
            "job_id": job_id,
            "status": "pending",
            "request": {
                "node_ids": request.node_ids,
                "date_start": request.date_start,
                "date_end": request.date_end,
                "hours": request.resolve_hours(),
            },
            "created_at": now,
            "started_at": None,
            "finished_at": None,
            "progress": None,
            "result": None,
            "error": None,
        }
        self._store.create(job)
        self._runner.submit(self._run, job_id)
        return {
            "job_id": job_id,
            "status": "pending",
            "created_at": now.isoformat(),
        }

    def get(self, job_id: str) -> dict | None:
        job = self._store.get(job_id)
        if job is None:
            return None
        return serialize_job(job)

    def _run(self, job_id: str) -> None:
        job = self._store.get(job_id)
        if job is None:
            return
        self._gateway.run_job(job)


def serialize_job(job: dict) -> dict:
    def iso(value):
        return value.isoformat() if hasattr(value, "isoformat") else value

    return {
        "job_id": job["job_id"],
        "status": job["status"],
        "created_at": iso(job.get("created_at")),
        "started_at": iso(job.get("started_at")),
        "finished_at": iso(job.get("finished_at")),
        "progress": job.get("progress"),
        "result": job.get("result"),
        "error": job.get("error"),
    }

