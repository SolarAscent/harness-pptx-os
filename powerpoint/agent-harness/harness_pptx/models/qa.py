"""QA and Repair data models."""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field, model_validator


class Severity(str, Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class QACategory(str, Enum):
    TEXT_OVERFLOW = "text_overflow"
    ELEMENT_OVERLAP = "element_overlap"
    ALIGNMENT = "alignment"
    MARGIN = "margin"
    FONT_SIZE = "font_size"
    CONTRAST = "contrast"
    IMAGE_STRETCH = "image_stretch"
    CHART_COMPLETENESS = "chart_completeness"
    SLIDE_DENSITY = "slide_density"
    COLOR_THEME = "color_theme"
    PUNCTUATION = "punctuation"
    PAGE_NUMBER = "page_number"
    STRUCTURE_CONSISTENCY = "structure_consistency"


class QAIssue(BaseModel):
    """A single issue found during QA."""

    id: str
    slide_id: str
    element_id: str | None = None
    category: QACategory
    severity: Severity = Severity.WARNING
    message: str = ""
    detail: dict[str, Any] = Field(default_factory=dict)


class QAReport(BaseModel):
    """Complete QA report for a deck.

    Count fields (total_issues, errors, warnings, infos) are auto-computed
    from the issues list on validation.
    """

    deck_id: str
    passed: bool = False
    total_issues: int = 0
    errors: int = 0
    warnings: int = 0
    infos: int = 0
    issues: list[QAIssue] = Field(default_factory=list)

    @model_validator(mode="after")
    def _compute_counts(self) -> "QAReport":
        self.total_issues = len(self.issues)
        self.errors = sum(1 for i in self.issues if i.severity == Severity.ERROR)
        self.warnings = sum(1 for i in self.issues if i.severity == Severity.WARNING)
        self.infos = sum(1 for i in self.issues if i.severity == Severity.INFO)
        self.passed = self.errors == 0
        return self


class RepairActionType(str, Enum):
    SHRINK_TEXT = "shrink_text"
    EXPAND_BOX = "expand_box"
    MOVE_ELEMENT = "move_element"
    SPLIT_SLIDE = "split_slide"
    SIMPLIFY_BULLETS = "simplify_bullets"
    PROMOTE_TO_APPENDIX = "promote_to_appendix"
    RESIZE_IMAGE = "resize_image"
    REBALANCE_COLUMNS = "rebalance_columns"
    INCREASE_CONTRAST = "increase_contrast"
    ALIGN_GROUP = "align_group"
    ADJUST_MARGIN = "adjust_margin"
    FIX_PUNCTUATION = "fix_punctuation"
    ADD_CHART_LABELS = "add_chart_labels"
    ADD_PAGE_NUMBER = "add_page_number"


class RepairAction(BaseModel):
    """A single repair action to fix a QA issue."""

    id: str
    issue_id: str
    action: RepairActionType
    slide_id: str
    element_id: str | None = None
    params: dict[str, Any] = Field(default_factory=dict)
    description: str = ""


class RepairPlan(BaseModel):
    """Ordered plan of repair actions."""

    deck_id: str
    actions: list[RepairAction] = Field(default_factory=list)
