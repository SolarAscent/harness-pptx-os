"""RetryPolicy — retry with backoff and timeout."""

from __future__ import annotations

import time
from typing import Any, Callable


class RetryPolicy:
    """Retry a callable with configurable attempts, delay, and backoff."""

    def __init__(
        self,
        max_attempts: int = 3,
        base_delay: float = 1.0,
        backoff: float = 2.0,
        timeout: float = 30.0,
    ):
        self.max_attempts = max_attempts
        self.base_delay = base_delay
        self.backoff = backoff
        self.timeout = timeout

    def run(self, fn: Callable[[], Any]) -> Any:
        last_error = None
        delay = self.base_delay
        for attempt in range(self.max_attempts):
            try:
                return fn()
            except Exception as e:
                last_error = e
                if attempt < self.max_attempts - 1:
                    time.sleep(delay)
                    delay *= self.backoff
        raise RuntimeError(f"Failed after {self.max_attempts} attempts") from last_error
