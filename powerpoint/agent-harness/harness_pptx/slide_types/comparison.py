"""Comparison slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ComparisonSlide(SlideType):
    name = "comparison"
    required_fields = ["title", "left", "right"]
    optional_fields = ["left_label", "right_label", "criteria"]
    visual_rules = {"max_per_column": 6}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "comparison-1")
        title = content.get("title", "Comparison")
        left_items = content["left"]
        right_items = content["right"]
        left_label = content.get("left_label", "Option A")
        right_label = content.get("right_label", "Option B")
        seq = content.get("seq", 0)

        elements: list = []
        element_ids: list[str] = []

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

        col_w = 380
        gap = 40
        left_x = 80
        right_x = 80 + col_w + gap

        # Column headers
        for x, label in [(left_x, left_label), (right_x, right_label)]:
            hdr_id = f"{slide_id}-hdr-{label}"
            elements.append(ShapeElement(
                id=hdr_id, role=ElementRole.DECORATION,
                shape_type="rounded_rectangle",
                bbox=BBox(x=x, y=100, w=col_w, h=36),
                fill_color="primary", line_color="primary",
                corner_radius=4,
            ))
            element_ids.append(hdr_id)

            htxt_id = f"{slide_id}-htxt-{label}"
            elements.append(TextElement(
                id=htxt_id, role=ElementRole.SUBTITLE,
                text=label,
                bbox=BBox(x=x + 8, y=104, w=col_w - 16, h=28),
                style={"font_size": 16, "bold": True, "font_color": "background", "alignment": "center"},
            ))
            element_ids.append(htxt_id)

        # Items
        for col_x, items, prefix in [(left_x, left_items, "L"), (right_x, right_items, "R")]:
            y = 152
            for i, item in enumerate(items[:6]):
                iid = f"{slide_id}-{prefix}-{i}"
                elements.append(TextElement(
                    id=iid, role=ElementRole.BODY,
                    text=str(item),
                    bbox=BBox(x=col_x + 8, y=y, w=col_w - 16, h=28),
                    style={"font_size": 14},
                    bullet=True,
                ))
                element_ids.append(iid)
                y += 34

        # Divider line
        div_id = f"{slide_id}-divider"
        mid_x = 480
        elements.append(ShapeElement(
            id=div_id, role=ElementRole.DECORATION,
            shape_type="line",
            bbox=BBox(x=mid_x, y=100, w=0, h=400),
            line_color="border", line_weight=1,
        ))
        element_ids.append(div_id)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="comparison",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
