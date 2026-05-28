"""PPTX XML backend skeleton — cross-platform OOXML manipulation.

Uses python-pptx for direct XML-level slide construction.
Currently a skeleton; full implementation is deferred.
"""

from __future__ import annotations

from typing import Any

from harness_pptx.backends.interface import (
    ChartStyle,
    ElementHandle,
    PresentationHandle,
    RendererInterface,
    ShapeStyle,
    SlideHandle,
    TableStyle,
    TextMetrics,
    TextStyle,
)


class PPTXXmlBackend(RendererInterface):
    """Cross-platform backend using python-pptx for OOXML manipulation.

    Currently a minimal skeleton. Full implementation would use
    python-pptx's slide, shape, and XML APIs.
    """

    backend_name = "pptx-xml"
    platform = "cross"

    def __init__(self):
        self._prs = None

    def create_presentation(self) -> PresentationHandle:
        try:
            from pptx import Presentation
        except ImportError:
            raise RuntimeError("python-pptx is required for PPTXXmlBackend")
        self._prs = Presentation()
        self._prs.slide_width = 960 * 12700  # EMU
        self._prs.slide_height = 540 * 12700
        return PresentationHandle(id="pptx-1")

    def open_presentation(self, path: str) -> PresentationHandle:
        try:
            from pptx import Presentation
        except ImportError:
            raise RuntimeError("python-pptx is required for PPTXXmlBackend")
        self._prs = Presentation(path)
        return PresentationHandle(id="pptx-1", path=path)

    def add_slide(self, pres: PresentationHandle, layout: str = "blank") -> SlideHandle:
        if self._prs is None:
            raise RuntimeError("No presentation open")
        slide_layout = self._prs.slide_layouts[6]  # blank layout
        self._prs.slides.add_slide(slide_layout)
        idx = len(self._prs.slides) - 1
        return SlideHandle(id=f"slide-{idx}", index=idx)

    def delete_slide(self, pres: PresentationHandle, slide: SlideHandle) -> None:
        raise NotImplementedError("PPTXXmlBackend.delete_slide")

    def add_text_box(
        self, slide: SlideHandle,
        x: float, y: float, w: float, h: float,
        text: str, style: TextStyle,
    ) -> ElementHandle:
        raise NotImplementedError("PPTXXmlBackend.add_text_box")

    def add_image(
        self, slide: SlideHandle,
        x: float, y: float, w: float, h: float,
        path: str, fit: str = "contain",
    ) -> ElementHandle:
        raise NotImplementedError("PPTXXmlBackend.add_image")

    def add_shape(
        self, slide: SlideHandle, shape_type: str,
        x: float, y: float, w: float, h: float, style: ShapeStyle,
    ) -> ElementHandle:
        raise NotImplementedError("PPTXXmlBackend.add_shape")

    def add_line(
        self, slide: SlideHandle,
        x1: float, y1: float, x2: float, y2: float,
        color: str = "#000000", weight: float = 1.0,
    ) -> ElementHandle:
        raise NotImplementedError("PPTXXmlBackend.add_line")

    def add_table(
        self, slide: SlideHandle,
        rows: int, cols: int, data: list[list[str]],
        x: float, y: float, w: float, h: float, style: TableStyle,
    ) -> ElementHandle:
        raise NotImplementedError("PPTXXmlBackend.add_table")

    def add_chart(
        self, slide: SlideHandle, chart_type: str,
        data: dict[str, Any],
        x: float, y: float, w: float, h: float, style: ChartStyle,
    ) -> ElementHandle:
        raise NotImplementedError("PPTXXmlBackend.add_chart")

    def add_group(
        self, slide: SlideHandle, elements: list[ElementHandle],
    ) -> ElementHandle:
        raise NotImplementedError("PPTXXmlBackend.add_group")

    def set_text(self, slide: SlideHandle, element: ElementHandle, text: str) -> None:
        raise NotImplementedError("PPTXXmlBackend.set_text")

    def delete_element(self, slide: SlideHandle, element: ElementHandle) -> None:
        raise NotImplementedError("PPTXXmlBackend.delete_element")

    def set_z_order(self, slide: SlideHandle, element: ElementHandle, action: str) -> None:
        raise NotImplementedError("PPTXXmlBackend.set_z_order")

    def set_slide_background(self, slide: SlideHandle, color: str) -> None:
        raise NotImplementedError("PPTXXmlBackend.set_slide_background")

    def slide_count(self, pres: PresentationHandle) -> int:
        if self._prs is None:
            return 0
        return len(self._prs.slides)

    def list_shapes(self, slide: SlideHandle) -> list[dict[str, Any]]:
        return []

    def save(self, pres: PresentationHandle, path: str) -> None:
        if self._prs is None:
            raise RuntimeError("No presentation open")
        self._prs.save(path)

    def export_pdf(self, pres: PresentationHandle, path: str) -> None:
        raise NotImplementedError("PPTXXmlBackend.export_pdf")

    def export_png(self, pres: PresentationHandle, slide_index: int, path: str) -> None:
        raise NotImplementedError("PPTXXmlBackend.export_png")

    def close(self, pres: PresentationHandle) -> None:
        self._prs = None
