from __future__ import annotations

from threading import Lock


class InMemoryJobStore:
    def __init__(self) -> None:
        self._items: dict[str, dict] = {}
        self._lock = Lock()

    def create(self, job: dict) -> None:
        with self._lock:
            self._items[job["job_id"]] = job

    def get(self, job_id: str) -> dict | None:
        with self._lock:
            return self._items.get(job_id)

    def update(self, job_id: str, changes: dict) -> None:
        with self._lock:
            if job_id in self._items:
                self._items[job_id].update(changes)

