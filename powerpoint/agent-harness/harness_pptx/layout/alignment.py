"""Alignment helpers for post-layout adjustment."""

from __future__ import annotations

from harness_pptx.models.layout import BBox


def safe_area(
    page_width: float = 960.0,
    page_height: float = 540.0,
    margin: float = 36.0,
) -> BBox:
    """Standard safe area with uniform margin."""
    return BBox(x=margin, y=margin, w=page_width - 2 * margin, h=page_height - 2 * margin)


def align_left(bbox: BBox, to: BBox, offset: float = 0) -> BBox:
    return BBox(x=to.x + offset, y=bbox.y, w=bbox.w, h=bbox.h)


def align_right(bbox: BBox, to: BBox, offset: float = 0) -> BBox:
    return BBox(x=to.x + to.w - bbox.w - offset, y=bbox.y, w=bbox.w, h=bbox.h)


def align_center_h(bbox: BBox, to: BBox) -> BBox:
    return BBox(x=to.x + (to.w - bbox.w) / 2, y=bbox.y, w=bbox.w, h=bbox.h)


def align_center_v(bbox: BBox, to: BBox) -> BBox:
    return BBox(x=bbox.x, y=to.y + (to.h - bbox.h) / 2, w=bbox.w, h=bbox.h)


def distribute_horizontal(bboxes: list[BBox], container: BBox, gap: float = 0) -> list[BBox]:
    """Evenly distribute bboxes horizontally within container."""
    n = len(bboxes)
    if n == 0:
        return []
    total_w = sum(b.w for b in bboxes)
    total_gap = gap * (n - 1)
    x = container.x + (container.w - total_w - total_gap) / 2

    result = []
    for b in bboxes:
        result.append(BBox(x=x, y=b.y, w=b.w, h=b.h))
        x += b.w + gap
    return result


def fit_to_content(bbox: BBox, padding: float = 0) -> BBox:
    """Inset bbox by padding (used to shrink container to content)."""
    return bbox.inset(-padding, -padding, -padding, -padding)


def stack_below(top_bbox: BBox, h: float, w: float | None = None, gap: float = 8) -> BBox:
    """Create a BBox positioned below another."""
    return BBox(
        x=top_bbox.x,
        y=top_bbox.y + top_bbox.h + gap,
        w=w or top_bbox.w,
        h=h,
    )


def stack_right(left_bbox: BBox, w: float, h: float | None = None, gap: float = 8) -> BBox:
    """Create a BBox positioned to the right of another."""
    return BBox(
        x=left_bbox.x + left_bbox.w + gap,
        y=left_bbox.y,
        w=w,
        h=h or left_bbox.h,
    )
