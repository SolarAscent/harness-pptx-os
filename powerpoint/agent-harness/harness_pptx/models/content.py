"""Content-understanding layer models.

These models represent the intermediate artifacts produced by the content
understanding pipeline: Brief → NarrativeStructure → Outline → SlideIntent.
"""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field

from harness_pptx.models.deck_spec import Language, Tone
from harness_pptx.models.slide_spec import SlideType


# ---- Brief ------------------------------------------------------------------

class Brief(BaseModel):
    """Parsed understanding of the user's raw request."""

    topic: str = ""
    audience: str = "general"
    goal: str = "inform"
    tone: Tone = Tone.PROFESSIONAL
    language: Language = Language.EN
    duration: int | None = None
    key_points: list[str] = Field(default_factory=list)
    constraints: dict[str, Any] = Field(default_factory=dict)
    source_text: str = ""


# ---- Outline ----------------------------------------------------------------

class OutlineItem(BaseModel):
    """One item in the slide outline — a provisional slide before typing."""

    seq: int
    title: str = ""
    key_message: str = ""
    section: str | None = None
    estimated_slide_type: SlideType = SlideType.CUSTOM


class Outline(BaseModel):
    """The complete slide outline — pages, sections, and narrative flow."""

    title: str = ""
    total_slides: int = 0
    sections: list[str] = Field(default_factory=list)
    items: list[OutlineItem] = Field(default_factory=list)


# ---- Slide Intent -----------------------------------------------------------

class SlideIntent(BaseModel):
    """Classified intent for a single slide — the bridge between outline and
    slide-type template selection."""

    slide_id: str
    seq: int
    slide_type: SlideType
    title: str = ""
    key_message: str = ""
    bullet_points: list[str] = Field(default_factory=list)
    visual_type: str | None = None
    data_refs: list[str] = Field(default_factory=list)
    notes: str = ""
    section: str | None = None
    narrative_role: str | None = None
    extra: dict[str, Any] = Field(default_factory=dict)
