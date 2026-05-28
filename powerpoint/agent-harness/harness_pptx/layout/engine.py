"""Layout engine — resolves LayoutSpec declarations to absolute BBox values."""

from __future__ import annotations

from harness_pptx.models.layout import (
    AlignH,
    AlignV,
    BBox,
    LayoutChild,
    LayoutDirection,
    LayoutSpec,
)


class LayoutEngine:
    """Resolves a LayoutSpec tree into absolute BBox values for each child.

    The engine works on a given parent container BBox and recursively
    resolves nested layouts. It uses layout primitives defined in
    ``primitives.py``.
    """

    def __init__(self, default_canvas: BBox | None = None):
        self.default_canvas = default_canvas or BBox(x=0, y=0, w=960, h=540)

    def resolve(
        self,
        spec: LayoutSpec,
        container: BBox | None = None,
    ) -> list[tuple[LayoutChild, BBox]]:
        """Resolve a LayoutSpec into (child, bbox) pairs.

        Args:
            spec: The layout specification to resolve.
            container: The parent container. Defaults to the engine's default canvas.

        Returns:
            A list of (LayoutChild, BBox) tuples in layout order.
        """
        container = container or self.default_canvas
        return self._dispatch(spec, container)

    def _dispatch(
        self, spec: LayoutSpec, container: BBox
    ) -> list[tuple[LayoutChild, BBox]]:
        from harness_pptx.layout.primitives import (
            columns,
            grid,
            hero,
            hstack,
            overlay,
            sidebar,
            split,
            vstack,
        )

        direction = spec.type
        children = spec.children
        gap = spec.gap

        if direction == LayoutDirection.VSTACK:
            return vstack(container, children, gap, spec.align_h)
        elif direction == LayoutDirection.HSTACK:
            return hstack(container, children, gap, spec.align_v)
        elif direction == LayoutDirection.GRID:
            return grid(
                container,
                children,
                rows=spec.rows or 1,
                cols=spec.cols or 1,
                gap=gap,
            )
        elif direction == LayoutDirection.COLUMNS:
            return columns(container, children, gap)
        elif direction == LayoutDirection.SPLIT:
            return split(container, children, spec.ratio or 0.5, gap)
        elif direction == LayoutDirection.SIDEBAR:
            return sidebar(container, children, spec.ratio or 0.25, gap)
        elif direction == LayoutDirection.HERO:
            return hero(container, children, spec.ratio or 0.6)
        elif direction == LayoutDirection.OVERLAY:
            return overlay(container, children)
        elif direction == LayoutDirection.ABSOLUTE:
            return [(c, container) for c in children]
        elif direction == LayoutDirection.FIT:
            return [(c, container) for c in children]
        else:
            return [(c, container) for c in children]
