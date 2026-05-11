from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from typing import Callable


class ThreadPoolBackgroundJobRunner:
    def __init__(self, max_workers: int = 4) -> None:
        self._executor = ThreadPoolExecutor(max_workers=max_workers)

    def submit(self, fn: Callable, *args) -> None:
        self._executor.submit(fn, *args)

    def shutdown(self) -> None:
        self._executor.shutdown(wait=False)

