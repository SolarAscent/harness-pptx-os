"""Element model — the atomic building blocks of every slide.

Every element carries a stable ``id`` (never PowerPoint shape index),
a semantic ``role``, an optional ``bbox`` (set by the layout engine),
and structured ``content``.
"""

from __future__ import annotations

from enum import Enum
from typing import Any, Literal

from pydantic import BaseModel, Field

from harness_pptx.models.layout import BBox


# ---- Element role -----------------------------------------------------------

class ElementRole(str, Enum):
    TITLE = "title"
    SUBTITLE = "subtitle"
    BODY = "body"
    CAPTION = "caption"
    FOOTNOTE = "footnote"
    CALLOUT = "callout"
    BADGE = "badge"
    LABEL = "label"
    KICKER = "kicker"
    PAGE_NUMBER = "page_number"
    LOGO = "logo"
    DECORATION = "decoration"


# ---- Style overrides --------------------------------------------------------

class StyleOverrides(BaseModel):
    """Per-element style tweaks referenced to theme tokens.

    A value of ``None`` means "inherit from theme / parent".
    """

    font_size: float | None = None
    font_color: str | None = None  # token name or #RRGGBB
    bold: bool | None = None
    italic: bool | None = None
    alignment: Literal["left", "center", "right", "justify"] | None = None
    fill_color: str | None = None
    line_color: str | None = None
    line_weight: float | None = None
    opacity: float | None = None
    z_index: int | None = None


# ---- Element types ----------------------------------------------------------

class ElementType(str, Enum):
    TEXT = "text"
    IMAGE = "image"
    SHAPE = "shape"
    CHART = "chart"
    TABLE = "table"
    DIAGRAM = "diagram"
    FORMULA = "formula"
    GROUP = "group"


# ---- Base element -----------------------------------------------------------

class BaseElement(BaseModel):
    """Every element on a slide derives from this."""

    id: str
    type: ElementType
    role: ElementRole = ElementRole.BODY
    bbox: BBox | None = None
    style: StyleOverrides = Field(default_factory=StyleOverrides)
    constraints: dict[str, Any] = Field(default_factory=dict)
    source: str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)


# ---- Concrete element types -------------------------------------------------

class TextElement(BaseElement):
    """Rich text element with automatic sizing support."""

    type: ElementType = ElementType.TEXT
    text: str = ""
    font_name: str | None = None
    auto_size: bool = True
    max_lines: int | None = None
    line_spacing: float | None = None
    bullet: bool = False
    bullet_style: str | None = None
    numbered: bool = False
    lang: Literal["en", "zh", "mixed"] | None = None


class ImageElement(BaseElement):
    """Image with fit, crop, mask, and source tracking."""

    type: ElementType = ElementType.IMAGE
    path: str = ""
    alt_text: str = ""
    fit: Literal["contain", "cover", "fill", "none"] = "contain"
    crop: BBox | None = None
    border: bool = False
    shadow: str | None = None
    caption: str | None = None
    source_url: str | None = None


class ShapeElement(BaseElement):
    """Native PowerPoint shape or grouped composite."""

    type: ElementType = ElementType.SHAPE
    shape_type: str = "rectangle"
    fill_color: str | None = None
    line_color: str | None = None
    line_weight: float | None = None
    text: str | None = None
    corner_radius: float | None = None


class ChartElement(BaseElement):
    """Chart — either native (editable) or rendered (matplotlib SVG/PNG)."""

    type: ElementType = ElementType.CHART
    chart_type: str = "bar"
    data: dict[str, Any] = Field(default_factory=dict)
    native: bool = True
    title: str | None = None
    x_label: str | None = None
    y_label: str | None = None
    legend: bool = True
    image_path: str | None = None


class TableElement(BaseElement):
    """Structured table with header styling and column widths."""

    type: ElementType = ElementType.TABLE
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    col_widths: list[float] | None = None
    header_style: StyleOverrides = Field(default_factory=StyleOverrides)
    zebra: bool = True
    auto_width: bool = True


class DiagramElement(BaseElement):
    """Diagram from nodes/edges — flowchart, architecture, tree, etc."""

    type: ElementType = ElementType.DIAGRAM
    diagram_type: str = "flowchart"
    nodes: list[dict[str, Any]] = Field(default_factory=list)
    edges: list[dict[str, Any]] = Field(default_factory=list)
    layout_algorithm: str = "dagre"
    direction: Literal["TB", "BT", "LR", "RL"] = "TB"


class FormulaElement(BaseElement):
    """LaTeX formula rendered as SVG/PNG."""

    type: ElementType = ElementType.FORMULA
    latex: str = ""
    display: bool = True
    image_path: str | None = None
    source_code: str | None = None


class CodeBlockElement(BaseElement):
    """Syntax-highlighted code block."""

    type: ElementType = ElementType.TEXT
    code: str = ""
    language: str = "python"
    line_numbers: bool = False
    image_path: str | None = None


class GroupElement(BaseElement):
    """Composite element grouping sub-elements for unified transforms."""

    type: ElementType = ElementType.GROUP
    children: list[BaseElement] = Field(default_factory=list)
