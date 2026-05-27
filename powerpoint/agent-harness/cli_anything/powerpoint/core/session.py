from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class Session:
    active_presentation: str | None = None
    history: list[dict[str, Any]] = field(default_factory=list)

    def record(self, action: str, **details: Any) -> None:
        self.history.append({"action": action, **details})
