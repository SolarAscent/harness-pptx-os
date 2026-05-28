"""Text measurement — estimate PowerPoint text rendering dimensions.

Initially uses a simple character-based heuristic. When a PowerPoint backend
is available, it can be swapped for actual on-canvas measurement.
"""

from __future__ import annotations


class TextMetrics:
    """Measured / estimated dimensions of a text block."""

    def __init__(self, width: float, height: float, lines: int = 1):
        self.width = width
        self.height = height
        self.lines = lines


def measure_text(
    text: str,
    font_size: float = 14,
    max_width: float = 800,
    font_name: str = "Calibri",
    line_spacing: float = 1.2,
) -> TextMetrics:
    """Estimate rendered text dimensions.

    Uses a simple heuristic: average character width ≈ 0.55 × font_size.
    In a full pipeline, this would delegate to AppleScript/VBA for real
    on-canvas measurement via temporary PowerPoint text boxes.

    Args:
        text: The text to measure.
        font_size: Point size.
        max_width: Maximum available width in points.
        font_name: Font family (unused in simple heuristic).
        line_spacing: Line height multiplier.

    Returns:
        TextMetrics with estimated width, height, and line count.
    """
    if not text:
        return TextMetrics(0, 0, 0)

    # Heuristic: average character width ≈ 0.55 × font_size for Latin
    # For CJK text, characters are roughly 1.0 × font_size wide
    char_width = _char_width_estimate(font_size)

    # Split into paragraphs
    paragraphs = text.split("\n")
    total_lines = 0
    max_line_w = 0.0

    for para in paragraphs:
        para_w = len(para) * char_width
        if para_w <= max_width:
            lines = 1 if para else 1
            max_line_w = max(max_line_w, para_w)
        else:
            lines = max(1, int(para_w / max_width) + 1)
            max_line_w = max_width
        total_lines += lines

    total_lines = max(total_lines, 1)
    line_h = font_size * line_spacing
    total_h = total_lines * line_h

    return TextMetrics(
        width=min(max_line_w, max_width),
        height=total_h,
        lines=total_lines,
    )


def _char_width_estimate(font_size: float) -> float:
    """Average character width estimate in points."""
    return font_size * 0.55


def auto_size_text(
    text: str,
    target_bbox_w: float,
    target_bbox_h: float,
    min_font_size: float = 8,
    max_font_size: float = 72,
    font_name: str = "Calibri",
) -> float:
    """Find the largest font size that fits text in the given bbox.

    Uses binary search with the text measurement heuristic.
    """
    lo, hi = min_font_size, max_font_size
    best = lo

    for _ in range(12):  # binary search iterations
        mid = (lo + hi) / 2
        metrics = measure_text(text, mid, target_bbox_w, font_name)
        if metrics.width <= target_bbox_w and metrics.height <= target_bbox_h:
            best = mid
            lo = mid
        else:
            hi = mid

    return best
