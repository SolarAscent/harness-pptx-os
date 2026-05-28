"""DeckSpec — the top-level data contract for a complete presentation.

This is the central format that all agents produce and all renderers consume.
"""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field

from harness_pptx.models.slide_spec import SlideSpec
from harness_pptx.models.theme import Theme


class Tone(str, Enum):
    PROFESSIONAL = "professional"
    INSPIRING = "inspiring"
    ANALYTICAL = "analytical"
    EDUCATIONAL = "educational"
    PERSUASIVE = "persuasive"
    MINIMAL = "minimal"


class Language(str, Enum):
    EN = "en"
    ZH = "zh"
    MIXED = "mixed"


class DeckMeta(BaseModel):
    """Top-level deck metadata — audience, goal, tone, etc."""

    title: str = ""
    subtitle: str | None = None
    audience: str = "general"
    goal: str = "inform"
    tone: Tone = Tone.PROFESSIONAL
    language: Language = Language.EN
    duration: int | None = None
    author: str | None = None
    date: str | None = None
    tags: list[str] = Field(default_factory=list)
    extra: dict[str, Any] = Field(default_factory=dict)


class DeckSpec(BaseModel):
    """Complete, render-ready specification of a presentation."""

    deck: DeckMeta = Field(default_factory=DeckMeta)
    theme: Theme = Field(default_factory=Theme)
    slides: list[SlideSpec] = Field(default_factory=list)

    @property
    def slide_count(self) -> int:
        return len(self.slides)

    def slide_by_id(self, slide_id: str) -> SlideSpec | None:
        for s in self.slides:
            if s.id == slide_id:
                return s
        return None
