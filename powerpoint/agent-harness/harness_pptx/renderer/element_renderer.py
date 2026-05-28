"""Element renderer — dispatches elements by type to the backend."""

from __future__ import annotations

from harness_pptx.backends.interface import (
    ChartStyle,
    ElementHandle,
    RendererInterface,
    ShapeStyle,
    SlideHandle,
    TableStyle,
    TextStyle,
)
from harness_pptx.models.element import (
    BaseElement,
    ChartElement,
    DiagramElement,
    ElementType,
    FormulaElement,
    ImageElement,
    ShapeElement,
    TableElement,
    TextElement,
)
from harness_pptx.models.theme import Theme


class ElementRenderer:
    """Renders individual elements onto slides via a backend."""

    def __init__(self, backend: RendererInterface, theme: Theme | None = None):
        self._backend = backend
        self._theme = theme

    def _resolve_color(self, token_or_hex: str | None, fallback: str = "#000000") -> str:
        """Resolve a semantic color token to hex using the theme."""
        if token_or_hex is None:
            return fallback
        if token_or_hex.startswith("#"):
            return token_or_hex
        if self._theme:
            try:
                return self._theme.color(token_or_hex)
            except AttributeError:
                pass
        return fallback

    def render(self, slide: SlideHandle, element: BaseElement) -> ElementHandle:
        handler = self._dispatch.get(element.type)
        if handler is None:
            raise ValueError(f"Unsupported element type: {element.type}")
        return handler(slide, element)

    # ---- Text -----------------------------------------------------------

    def _render_text(self, slide: SlideHandle, el: TextElement) -> ElementHandle:
        bbox = el.bbox
        if bbox is None:
            raise ValueError(f"Text element {el.id} has no bbox")
        style = TextStyle(
            font_name=el.font_name or "Calibri",
            font_size=el.style.font_size or 14,
            font_color=self._resolve_color(el.style.font_color, "#000000"),
            bold=el.style.bold or False,
            italic=el.style.italic or False,
            alignment=el.style.alignment or "left",
            line_spacing=el.line_spacing or 1.2,
        )
        return self._backend.add_text_box(
            slide, bbox.x, bbox.y, bbox.w, bbox.h, el.text, style,
        )

    # ---- Image ----------------------------------------------------------

    def _render_image(self, slide: SlideHandle, el: ImageElement) -> ElementHandle:
        bbox = el.bbox
        if bbox is None:
            raise ValueError(f"Image element {el.id} has no bbox")
        return self._backend.add_image(
            slide, bbox.x, bbox.y, bbox.w, bbox.h, el.path, el.fit,
        )

    # ---- Shape ----------------------------------------------------------

    def _render_shape(self, slide: SlideHandle, el: ShapeElement) -> ElementHandle:
        bbox = el.bbox
        if bbox is None:
            raise ValueError(f"Shape element {el.id} has no bbox")
        style = ShapeStyle(
            fill_color=self._resolve_color(el.fill_color, "#FFFFFF"),
            line_color=self._resolve_color(el.line_color, "#000000"),
            line_weight=el.line_weight or 0.5,
            corner_radius=el.corner_radius or 0,
        )
        if el.shape_type == "line" and bbox.h <= 2:
            return self._backend.add_line(
                slide, bbox.x, bbox.y, bbox.x + bbox.w, bbox.y + bbox.h,
                color=self._resolve_color(el.line_color, "#000000"),
                weight=el.line_weight or 0.5,
            )
        return self._backend.add_shape(
            slide, el.shape_type, bbox.x, bbox.y, bbox.w, bbox.h, style,
        )

    # ---- Chart ----------------------------------------------------------

    def _render_chart(self, slide: SlideHandle, el: ChartElement) -> ElementHandle:
        bbox = el.bbox
        if bbox is None:
            raise ValueError(f"Chart element {el.id} has no bbox")
        style = ChartStyle(
            show_legend=el.legend,
        )
        return self._backend.add_chart(
            slide, el.chart_type, el.data,
            bbox.x, bbox.y, bbox.w, bbox.h, style,
        )

    # ---- Table ----------------------------------------------------------

    def _render_table(self, slide: SlideHandle, el: TableElement) -> ElementHandle:
        bbox = el.bbox
        if bbox is None:
            raise ValueError(f"Table element {el.id} has no bbox")
        data = [el.headers] + el.rows if el.headers else el.rows
        rows = len(data)
        cols = max(len(el.headers), max((len(r) for r in el.rows), default=0))
        style = TableStyle()
        return self._backend.add_table(
            slide, rows, cols, data,
            bbox.x, bbox.y, bbox.w, bbox.h, style,
        )

    # ---- Diagram / Formula (fallback) -----------------------------------

    def _render_diagram(self, slide: SlideHandle, el: DiagramElement) -> ElementHandle:
        bbox = el.bbox
        if bbox is None:
            raise ValueError(f"Diagram element {el.id} has no bbox")
        # Fallback: render as text placeholder + outer box
        label = f"[Diagram: {el.diagram_type} — {len(el.nodes)} nodes]"
        style = TextStyle(font_size=14, font_color="#767676")
        self._backend.add_text_box(
            slide, bbox.x, bbox.y, bbox.w, bbox.h, label, style,
        )
        return ElementHandle(id=el.id, element_type="diagram")

    def _render_formula(self, slide: SlideHandle, el: FormulaElement) -> ElementHandle:
        bbox = el.bbox
        if bbox is None:
            raise ValueError(f"Formula element {el.id} has no bbox")
        label = f"${el.latex}$" if el.display else f"\\({el.latex}\\)"
        style = TextStyle(font_name="Consolas", font_size=14, font_color="#000000")
        return self._backend.add_text_box(
            slide, bbox.x, bbox.y, bbox.w, bbox.h, label, style,
        )

    # ---- Dispatch table -------------------------------------------------

    @property
    def _dispatch(self):
        return {
            ElementType.TEXT: self._render_text,
            ElementType.IMAGE: self._render_image,
            ElementType.SHAPE: self._render_shape,
            ElementType.CHART: self._render_chart,
            ElementType.TABLE: self._render_table,
            ElementType.DIAGRAM: self._render_diagram,
            ElementType.FORMULA: self._render_formula,
        }
