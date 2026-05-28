"""AgentMemory — record slide intent, source, transformation, and QA history."""

from __future__ import annotations

from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


class SlideMemory(BaseModel):
    slide_id: str
    intent: str = ""
    source_text: str = ""
    transformation: str = ""
    template: str = ""
    constraints: dict[str, Any] = Field(default_factory=dict)
    qa_history: list[dict[str, Any]] = Field(default_factory=list)
    created_at: str = Field(default_factory=lambda: datetime.now().isoformat())


class AgentMemory:
    """Records the rationale behind each slide for future agent context."""

    def __init__(self):
        self._records: dict[str, SlideMemory] = {}

    def record(
        self,
        slide_id: str,
        intent: str = "",
        source_text: str = "",
        transformation: str = "",
        template: str = "",
        constraints: dict[str, Any] | None = None,
    ) -> SlideMemory:
        mem = SlideMemory(
            slide_id=slide_id,
            intent=intent,
            source_text=source_text,
            transformation=transformation,
            template=template,
            constraints=constraints or {},
        )
        self._records[slide_id] = mem
        return mem

    def log_qa(self, slide_id: str, issue_id: str, result: str) -> None:
        if slide_id in self._records:
            self._records[slide_id].qa_history.append({
                "issue_id": issue_id,
                "result": result,
                "timestamp": datetime.now().isoformat(),
            })

    def get(self, slide_id: str) -> SlideMemory | None:
        return self._records.get(slide_id)

    def all(self) -> dict[str, SlideMemory]:
        return dict(self._records)
