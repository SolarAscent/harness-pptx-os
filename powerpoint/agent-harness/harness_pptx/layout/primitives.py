"""Layout primitives — pure functions that compute BBox allocations."""

from __future__ import annotations

from harness_pptx.models.layout import AlignH, AlignV, BBox, LayoutChild


# ---- Vertical stack ---------------------------------------------------------

def vstack(
    container: BBox,
    children: list[LayoutChild],
    gap: float = 0,
    align: AlignH = AlignH.LEFT,
) -> list[tuple[LayoutChild, BBox]]:
    """Stack children vertically within container."""
    result: list[tuple[LayoutChild, BBox]] = []
    total_flex = sum(c.flex for c in children)
    available_h = container.h - gap * (len(children) - 1)
    total_fixed_h = sum(
        c.height for c in children if c.height is not None and c.flex == 0
    )
    flex_h = max(0, available_h - total_fixed_h)

    y = container.y
    for child in children:
        if child.height is not None and child.flex == 0:
            h = child.height
        elif total_flex > 0:
            h = (flex_h * child.flex) / total_flex
        else:
            h = available_h / max(len(children), 1)

        w = child.width if child.width is not None else container.w

        x = _resolve_align_h(container, w, child.margin, align)

        h = _clamp_dim(h, child.min_height, child.max_height)
        w = _clamp_dim(w, child.min_width, child.max_width)

        result.append((child, BBox(x=x, y=y, w=w, h=h)))
        y += h + gap

    return result


# ---- Horizontal stack -------------------------------------------------------

def hstack(
    container: BBox,
    children: list[LayoutChild],
    gap: float = 0,
    align: AlignV = AlignV.TOP,
) -> list[tuple[LayoutChild, BBox]]:
    """Stack children horizontally within container."""
    result: list[tuple[LayoutChild, BBox]] = []
    total_flex = sum(c.flex for c in children)
    available_w = container.w - gap * (len(children) - 1)
    total_fixed_w = sum(
        c.width for c in children if c.width is not None and c.flex == 0
    )
    flex_w = max(0, available_w - total_fixed_w)

    x = container.x
    for child in children:
        if child.width is not None and child.flex == 0:
            w = child.width
        elif total_flex > 0:
            w = (flex_w * child.flex) / total_flex
        else:
            w = available_w / max(len(children), 1)

        h = child.height if child.height is not None else container.h

        y = _resolve_align_v(container, h, child.margin, align)

        w = _clamp_dim(w, child.min_width, child.max_width)
        h = _clamp_dim(h, child.min_height, child.max_height)

        result.append((child, BBox(x=x, y=y, w=w, h=h)))
        x += w + gap

    return result


# ---- Grid -------------------------------------------------------------------

def grid(
    container: BBox,
    children: list[LayoutChild],
    rows: int,
    cols: int,
    gap: float = 0,
) -> list[tuple[LayoutChild, BBox]]:
    """Arrange children in a grid."""
    result: list[tuple[LayoutChild, BBox]] = []
    cell_w = (container.w - gap * (cols - 1)) / cols
    cell_h = (container.h - gap * (rows - 1)) / rows if rows > 0 else 0

    for i, child in enumerate(children):
        if i >= rows * cols:
            break
        r = i // cols
        c = i % cols
        x = container.x + c * (cell_w + gap)
        y = container.y + r * (cell_h + gap)
        result.append((child, BBox(x=x, y=y, w=cell_w, h=cell_h)))

    return result


# ---- Columns ----------------------------------------------------------------

def columns(
    container: BBox,
    children: list[LayoutChild],
    gap: float = 0,
) -> list[tuple[LayoutChild, BBox]]:
    """Evenly split container into N columns."""
    n = max(len(children), 1)
    col_w = (container.w - gap * (n - 1)) / n
    result: list[tuple[LayoutChild, BBox]] = []

    for i, child in enumerate(children):
        x = container.x + i * (col_w + gap)
        result.append((child, BBox(x=x, y=container.y, w=col_w, h=container.h)))

    return result


# ---- Split ------------------------------------------------------------------

def split(
    container: BBox,
    children: list[LayoutChild],
    ratio: float = 0.5,
    gap: float = 0,
) -> list[tuple[LayoutChild, BBox]]:
    """Split container into two panes at given ratio."""
    result: list[tuple[LayoutChild, BBox]] = []
    if len(children) < 2:
        return columns(container, children, gap)

    left_w = (container.w - gap) * ratio
    right_w = container.w - gap - left_w

    result.append((
        children[0],
        BBox(x=container.x, y=container.y, w=left_w, h=container.h),
    ))
    result.append((
        children[1],
        BBox(x=container.x + left_w + gap, y=container.y, w=right_w, h=container.h),
    ))

    return result


# ---- Sidebar ----------------------------------------------------------------

def sidebar(
    container: BBox,
    children: list[LayoutChild],
    ratio: float = 0.25,
    gap: float = 0,
) -> list[tuple[LayoutChild, BBox]]:
    """Narrow sidebar + wide content area."""
    return split(container, children, ratio, gap)


# ---- Hero -------------------------------------------------------------------

def hero(
    container: BBox,
    children: list[LayoutChild],
    ratio: float = 0.6,
) -> list[tuple[LayoutChild, BBox]]:
    """Hero section on top, content below."""
    if not children:
        return []

    hero_h = container.h * ratio
    content_h = container.h - hero_h

    result: list[tuple[LayoutChild, BBox]] = [
        (children[0], BBox(x=container.x, y=container.y, w=container.w, h=hero_h)),
    ]

    if len(children) > 1:
        result.append((
            children[1],
            BBox(x=container.x, y=container.y + hero_h, w=container.w, h=content_h),
        ))

    return result


# ---- Overlay ----------------------------------------------------------------

def overlay(
    container: BBox,
    children: list[LayoutChild],
) -> list[tuple[LayoutChild, BBox]]:
    """All children get the full container (for absolute positioning)."""
    return [(c, BBox(x=container.x, y=container.y, w=container.w, h=container.h))
            for c in children]


# ---- Helpers ----------------------------------------------------------------

def _resolve_align_h(
    container: BBox, width: float, margin: float, align: AlignH
) -> float:
    if align == AlignH.CENTER:
        return container.x + (container.w - width) / 2
    elif align == AlignH.RIGHT:
        return container.x + container.w - width - margin
    return container.x + margin  # LEFT


def _resolve_align_v(
    container: BBox, height: float, margin: float, align: AlignV
) -> float:
    if align == AlignV.MIDDLE:
        return container.y + (container.h - height) / 2
    elif align == AlignV.BOTTOM:
        return container.y + container.h - height - margin
    return container.y + margin  # TOP


def _clamp_dim(
    value: float, min_val: float | None, max_val: float | None
) -> float:
    if min_val is not None:
        value = max(value, min_val)
    if max_val is not None:
        value = min(value, max_val)
    return value
