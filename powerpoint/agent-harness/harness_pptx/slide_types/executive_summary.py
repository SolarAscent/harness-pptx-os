"""Executive summary slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ExecutiveSummarySlide(SlideType):
    name = "executive-summary"
    required_fields = ["title", "key_points"]
    optional_fields = ["subtitle", "bottom_line"]
    visual_rules = {"max_points": 5}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "summary-1")
        title = content.get("title", "Executive Summary")
        key_points = content["key_points"]
        bottom_line = content.get("bottom_line", "")
        seq = content.get("seq", 0)

        elements: list = []
        element_ids: list[str] = []

        from harness_pptx.models.element import ShapeElement

        # Bottom accent bar
        bar_id = f"{slide_id}-accent-bar"
        elements.append(ShapeElement(
            id=bar_id, role=ElementRole.DECORATION,
            shape_type="rectangle",
            bbox=BBox(x=0, y=512, w=960, h=28),
            fill_color="primary", line_color="primary",
        ))
        element_ids.append(bar_id)

        tid = f"{slide_id}-title"
        elements.append(TextElement(
            id=tid, role=ElementRole.TITLE,
            text=title,
            bbox=BBox(x=80, y=40, w=800, h=44),
            style={"font_size": 26, "bold": True},
        ))
        element_ids.append(tid)

        y = 104
        for i, point in enumerate(key_points[:5]):
            pid = f"{slide_id}-point-{i}"
            elements.append(TextElement(
                id=pid, role=ElementRole.BODY,
                text=point,
                bbox=BBox(x=100, y=y, w=760, h=32),
                style={"font_size": 16},
                bullet=True,
            ))
            element_ids.append(pid)
            y += 48

        if bottom_line:
            bid = f"{slide_id}-bottom"
            y = max(y + 24, 400)
            line_id = f"{slide_id}-sep"
            elements.append(ShapeElement(
                id=line_id, role=ElementRole.DECORATION,
                shape_type="rectangle",
                bbox=BBox(x=80, y=y, w=800, h=2),
                fill_color="border", line_color="border",
            ))
            element_ids.append(line_id)
            y += 24
            elements.append(TextElement(
                id=bid, role=ElementRole.CALLOUT,
                text=bottom_line,
                bbox=BBox(x=80, y=y, w=800, h=32),
                style={"font_size": 16, "bold": True, "font_color": "primary"},
            ))
            element_ids.append(bid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="executive-summary",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
