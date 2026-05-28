"""Layout package — engine, primitives, alignment, and text measurement."""

from harness_pptx.layout.engine import LayoutEngine
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
from harness_pptx.layout.alignment import (
    align_center_h,
    align_center_v,
    align_left,
    align_right,
    distribute_horizontal,
    fit_to_content,
    safe_area,
    stack_below,
    stack_right,
)
from harness_pptx.layout.text_measure import TextMetrics, auto_size_text, measure_text

__all__ = [
    "LayoutEngine",
    "vstack",
    "hstack",
    "grid",
    "columns",
    "split",
    "sidebar",
    "hero",
    "overlay",
    "safe_area",
    "align_left",
    "align_right",
    "align_center_h",
    "align_center_v",
    "distribute_horizontal",
    "fit_to_content",
    "stack_below",
    "stack_right",
    "TextMetrics",
    "measure_text",
    "auto_size_text",
]
