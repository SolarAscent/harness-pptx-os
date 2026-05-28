"""RendererInterface — abstract base for all PowerPoint rendering backends."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any


# ---- Handle types -----------------------------------------------------------

@dataclass
class PresentationHandle:
    """Opaque handle to an open presentation."""

    id: str
    path: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class SlideHandle:
    """Opaque handle to a slide in a presentation."""

    id: str
    index: int
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class ElementHandle:
    """Opaque handle to a rendered element."""

    id: str
    element_type: str
    shape_index: int | None = None
    metadata: dict[str, Any] = field(default_factory=dict)


# ---- Style structs ----------------------------------------------------------

@dataclass
class TextStyle:
    font_name: str = "Calibri"
    font_size: float = 14.0
    font_color: str = "#000000"
    bold: bool = False
    italic: bool = False
    alignment: str = "left"
    line_spacing: float = 1.2


@dataclass
class ShapeStyle:
    fill_color: str | None = None
    line_color: str | None = None
    line_weight: float = 0.5
    corner_radius: float = 0.0
    opacity: float = 1.0


@dataclass
class TableStyle:
    header_fill: str = "#003366"
    header_text: str = "#FFFFFF"
    body_fill: str = "#FFFFFF"
    alt_body_fill: str = "#F5F5F5"
    border_color: str = "#D0D0D0"
    font_size: float = 12.0


@dataclass
class ChartStyle:
    palette: list[str] = field(default_factory=list)
    show_legend: bool = True
    title_size: float = 14.0


# ---- Text measurement result ------------------------------------------------

@dataclass
class TextMetrics:
    width: float = 0.0
    height: float = 0.0
    lines: int = 1


# ---- Renderer interface -----------------------------------------------------

class RendererInterface(ABC):
    """Abstract interface for PowerPoint rendering backends.

    Implementations: AppleScript (macOS), VBA (Windows/macOS), pptx XML (cross-platform).
    """

    @abstractmethod
    def create_presentation(self) -> PresentationHandle: ...

    @abstractmethod
    def open_presentation(self, path: str) -> PresentationHandle: ...

    @abstractmethod
    def add_slide(self, pres: PresentationHandle, layout: str = "blank") -> SlideHandle: ...

    @abstractmethod
    def delete_slide(self, pres: PresentationHandle, slide: SlideHandle) -> None: ...

    @abstractmethod
    def add_text_box(
        self, slide: SlideHandle, x: float, y: float, w: float, h: float,
        text: str, style: TextStyle,
    ) -> ElementHandle: ...

    @abstractmethod
    def add_image(
        self, slide: SlideHandle, x: float, y: float, w: float, h: float,
        path: str, fit: str = "contain",
    ) -> ElementHandle: ...

    @abstractmethod
    def add_shape(
        self, slide: SlideHandle, shape_type: str,
        x: float, y: float, w: float, h: float, style: ShapeStyle,
    ) -> ElementHandle: ...

    @abstractmethod
    def add_line(
        self, slide: SlideHandle,
        x1: float, y1: float, x2: float, y2: float,
        color: str = "#000000", weight: float = 1.0,
    ) -> ElementHandle: ...

    @abstractmethod
    def add_table(
        self, slide: SlideHandle,
        rows: int, cols: int, data: list[list[str]],
        x: float, y: float, w: float, h: float, style: TableStyle,
    ) -> ElementHandle: ...

    @abstractmethod
    def add_chart(
        self, slide: SlideHandle, chart_type: str,
        data: dict[str, Any],
        x: float, y: float, w: float, h: float, style: ChartStyle,
    ) -> ElementHandle: ...

    @abstractmethod
    def add_group(
        self, slide: SlideHandle, elements: list[ElementHandle],
    ) -> ElementHandle: ...

    @abstractmethod
    def set_text(self, slide: SlideHandle, element: ElementHandle, text: str) -> None: ...

    @abstractmethod
    def delete_element(self, slide: SlideHandle, element: ElementHandle) -> None: ...

    @abstractmethod
    def set_z_order(self, slide: SlideHandle, element: ElementHandle, action: str) -> None: ...

    @abstractmethod
    def set_slide_background(self, slide: SlideHandle, color: str) -> None: ...

    @abstractmethod
    def slide_count(self, pres: PresentationHandle) -> int: ...

    @abstractmethod
    def list_shapes(self, slide: SlideHandle) -> list[dict[str, Any]]: ...

    @abstractmethod
    def save(self, pres: PresentationHandle, path: str) -> None: ...

    @abstractmethod
    def export_pdf(self, pres: PresentationHandle, path: str) -> None: ...

    @abstractmethod
    def export_png(self, pres: PresentationHandle, slide_index: int, path: str) -> None: ...

    def export_all_pngs(self, pres: PresentationHandle, output_dir: str) -> list[str]:
        """Export all slides as PNG images. Returns list of file paths."""
        return []

    @abstractmethod
    def close(self, pres: PresentationHandle) -> None: ...

    def measure_text(
        self, text: str, font_name: str, font_size: float, max_width: float
    ) -> TextMetrics:
        """Estimate rendered text dimensions. Backends may override."""
        char_w = font_size * 0.55
        text_w = len(text) * char_w
        lines = max(1, int(text_w / max_width) + 1) if text_w > max_width else 1
        return TextMetrics(
            width=min(text_w, max_width),
            height=lines * font_size * 1.2,
            lines=lines,
        )

    @property
    @abstractmethod
    def platform(self) -> str: ...

    @property
    @abstractmethod
    def backend_name(self) -> str: ...
