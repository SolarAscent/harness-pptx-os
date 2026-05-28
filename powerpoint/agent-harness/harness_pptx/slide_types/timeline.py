"""Timeline slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class TimelineSlide(SlideType):
    name = "timeline"
    required_fields = ["title", "milestones"]
    optional_fields = ["orientation"]
    visual_rules = {"max_milestones": 8}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "timeline-1")
        title = content.get("title", "Timeline")
        milestones = content["milestones"]
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

        # Horizontal timeline line
        line_id = f"{slide_id}-line"
        elements.append(ShapeElement(
            id=line_id, role=ElementRole.DECORATION,
            shape_type="line",
            bbox=BBox(x=80, y=280, w=800, h=2),
            line_color="border", line_weight=2,
        ))
        element_ids.append(line_id)

        n = min(len(milestones), 8)
        spacing = 800 // max(n, 1)
        for i, m in enumerate(milestones[:8]):
            cx = 80 + spacing * i + spacing // 2

            # Dot
            dot_id = f"{slide_id}-dot-{i}"
            elements.append(ShapeElement(
                id=dot_id, role=ElementRole.DECORATION,
                shape_type="circle",
                bbox=BBox(x=cx - 6, y=274, w=12, h=12),
                fill_color="accent", line_color="accent",
            ))
            element_ids.append(dot_id)

            # Label above
            lbl_id = f"{slide_id}-label-{i}"
            elements.append(TextElement(
                id=lbl_id, role=ElementRole.BODY,
                text=m.get("date", m.get("label", "")),
                bbox=BBox(x=cx - 60, y=220, w=120, h=28),
                style={"font_size": 13, "font_color": "muted", "alignment": "center"},
            ))
            element_ids.append(lbl_id)

            # Description below
            desc_id = f"{slide_id}-desc-{i}"
            elements.append(TextElement(
                id=desc_id, role=ElementRole.BODY,
                text=m.get("event", m.get("description", "")),
                bbox=BBox(x=cx - 60, y=300, w=120, h=60),
                style={"font_size": 13, "alignment": "center"},
            ))
            element_ids.append(desc_id)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="timeline",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
