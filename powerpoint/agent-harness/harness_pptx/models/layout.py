"""Layout primitives: BBox, LayoutSpec, and layout child definitions.

All coordinates in PowerPoint points (72 dpi). Default slide canvas: 960×540.
"""

from __future__ import annotations

from enum import Enum
from typing import Any, Literal

from pydantic import BaseModel, Field


# ---- BBox -------------------------------------------------------------------

class BBox(BaseModel):
    """Absolute or relative bounding box in PowerPoint points."""

    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0

    @property
    def right(self) -> float:
        return self.x + self.w

    @property
    def bottom(self) -> float:
        return self.y + self.h

    @property
    def center_x(self) -> float:
        return self.x + self.w / 2

    @property
    def center_y(self) -> float:
        return self.y + self.h / 2

    def shifted(self, dx: float = 0, dy: float = 0) -> "BBox":
        return BBox(x=self.x + dx, y=self.y + dy, w=self.w, h=self.h)

    def inset(self, top: float = 0, right: float = 0, bottom: float = 0, left: float = 0) -> "BBox":
        return BBox(
            x=self.x + left,
            y=self.y + top,
            w=max(0, self.w - left - right),
            h=max(0, self.h - top - bottom),
        )


# ---- Page spec --------------------------------------------------------------

class PageSpec(BaseModel):
    """Slide canvas dimensions in points."""

    width: float = 960.0
    height: float = 540.0

    @property
    def safe_bbox(self) -> BBox:
        """Return bbox with standard 36pt margin."""
        return BBox(x=0, y=0, w=self.width, h=self.height).inset(36, 36, 36, 36)


# ---- Layout direction -------------------------------------------------------

class LayoutDirection(str, Enum):
    VSTACK = "vstack"
    HSTACK = "hstack"
    GRID = "grid"
    COLUMNS = "columns"
    SPLIT = "split"
    SIDEBAR = "sidebar"
    HERO = "hero"
    OVERLAY = "overlay"
    ABSOLUTE = "absolute"
    FIT = "fit"


class AlignH(str, Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class AlignV(str, Enum):
    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"


class FitMode(str, Enum):
    CONTAIN = "contain"
    COVER = "cover"
    FILL = "fill"
    NONE = "none"


# ---- Layout child -----------------------------------------------------------

class LayoutChild(BaseModel):
    """A single child in a layout container.

    May be a nested LayoutSpec or a reference to an element id.
    """

    id: str | None = None
    element_ref: str | None = None
    width: float | None = None
    height: float | None = None
    flex: float = 1.0
    min_width: float | None = None
    min_height: float | None = None
    max_width: float | None = None
    max_height: float | None = None
    align_h: AlignH = AlignH.LEFT
    align_v: AlignV = AlignV.TOP
    padding: float = 0.0
    margin: float = 0.0
    children: list["LayoutChild"] = Field(default_factory=list)


# ---- Layout spec ------------------------------------------------------------

class LayoutSpec(BaseModel):
    """Declarative layout for a slide or a component.

    Layouts compose: a split may contain a vstack in its left pane.
    The LayoutEngine resolves them to absolute BBox values.
    """

    type: LayoutDirection
    gap: float = 8.0
    ratio: float | None = None
    rows: int | None = None
    cols: int | None = None
    align_h: AlignH = AlignH.LEFT
    align_v: AlignV = AlignV.TOP
    children: list[LayoutChild] = Field(default_factory=list)
    extra: dict[str, Any] = Field(default_factory=dict)
