from __future__ import annotations

import logging
import threading
from concurrent.futures import Future, ThreadPoolExecutor
from typing import Callable


logger = logging.getLogger(__name__)


class TaskManager:
    def __init__(self, max_workers: int = 4) -> None:
        self._executor = ThreadPoolExecutor(max_workers=max_workers)
        self._active: dict[str, Future] = {}
        self._lock = threading.Lock()

    def submit_task(self, task_id: str, fn: Callable, *args, **kwargs) -> None:
        future = self._executor.submit(fn, task_id, *args, **kwargs)
        with self._lock:
            self._active[task_id] = future
        future.add_done_callback(lambda f: self._on_task_done(task_id, f))

    def _on_task_done(self, task_id: str, future: Future) -> None:
        try:
            exc = future.exception()
            if exc is not None:
                logger.error("Task %s failed: %s", task_id, exc, exc_info=exc)
        except Exception as callback_exc:  # pragma: no cover
            logger.error("Task callback failed for %s: %s", task_id, callback_exc, exc_info=True)
        finally:
            with self._lock:
                self._active.pop(task_id, None)

    def is_active(self, task_id: str) -> bool:
        with self._lock:
            return task_id in self._active


# Global shared task manager instance
task_manager = TaskManager(max_workers=4)

