"""SlideSpec — single slide within a DeckSpec."""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field

from harness_pptx.models.element import BaseElement
from harness_pptx.models.layout import LayoutSpec


class SlideType(str, Enum):
    COVER = "cover"
    AGENDA = "agenda"
    SECTION_DIVIDER = "section-divider"
    EXECUTIVE_SUMMARY = "executive-summary"
    PROBLEM = "problem"
    SOLUTION = "solution"
    TIMELINE = "timeline"
    PROCESS = "process"
    FRAMEWORK = "framework"
    COMPARISON = "comparison"
    BEFORE_AFTER = "before-after"
    DATA_INSIGHT = "data-insight"
    CHART = "chart"
    TABLE = "table"
    CASE_STUDY = "case-study"
    QUOTE = "quote"
    TEAM = "team"
    ROADMAP = "roadmap"
    ARCHITECTURE = "architecture"
    WORKFLOW = "workflow"
    RISK = "risk"
    RECOMMENDATION = "recommendation"
    CONCLUSION = "conclusion"
    THANK_YOU = "thank-you"
    APPENDIX = "appendix"
    CUSTOM = "custom"


class SlideLayer(BaseModel):
    """A named layer within a slide for z-ordering."""

    name: str
    z_index: int = 0
    elements: list[BaseElement] = Field(default_factory=list)


class SlideSpec(BaseModel):
    """Single slide specification.

    Every slide has a stable ``id``, a ``type`` that maps to a template,
    an ``intent`` that describes its narrative role, and structured content.
    """

    id: str
    type: SlideType = SlideType.CUSTOM
    title: str = ""
    intent: str = ""
    content: dict[str, Any] = Field(default_factory=dict)
    layout: LayoutSpec | None = None
    layers: list[SlideLayer] = Field(default_factory=list)
    elements: list[BaseElement] = Field(default_factory=list)
    notes: str = ""
    constraints: dict[str, Any] = Field(default_factory=dict)
    metadata: dict[str, Any] = Field(default_factory=dict)
