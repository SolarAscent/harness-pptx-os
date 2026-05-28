"""Transaction — operation wrapper with rollback support."""

from __future__ import annotations

from typing import Any, Callable


class Transaction:
    """Wraps an operation with optional rollback.

    Usage::

        txn = Transaction()
        txn.add_step(lambda: do_thing(), lambda: undo_thing())
        txn.commit()
    """

    def __init__(self):
        self._steps: list[tuple[Callable[[], Any], Callable[[], None] | None]] = []

    def add_step(
        self,
        forward: Callable[[], Any],
        rollback: Callable[[], None] | None = None,
    ) -> None:
        self._steps.append((forward, rollback))

    def commit(self) -> list[Any]:
        results = []
        executed = 0
        try:
            for forward, _ in self._steps:
                results.append(forward())
                executed += 1
        except Exception:
            # Rollback in reverse
            for i in range(executed - 1, -1, -1):
                _, rollback = self._steps[i]
                if rollback:
                    try:
                        rollback()
                    except Exception:
                        pass
            raise
        return results
