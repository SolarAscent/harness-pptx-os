"""Conclusion slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ConclusionSlide(SlideType):
    name = "conclusion"
    required_fields = ["title", "key_takeaways"]
    optional_fields = ["call_to_action", "subtitle"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "conclusion-1")
        title = content.get("title", "Conclusion")
        takeaways = content["key_takeaways"]
        call_to_action = content.get("call_to_action", "")
        subtitle = content.get("subtitle", "")
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

        if subtitle:
            sid = f"{slide_id}-sub"
            elements.append(TextElement(
                id=sid, role=ElementRole.SUBTITLE,
                text=subtitle,
                bbox=BBox(x=80, y=88, w=800, h=24),
                style={"font_size": 14, "font_color": "muted"},
            ))
            element_ids.append(sid)

        y = 130
        for i, takeaway in enumerate(takeaways[:6]):
            # Accent bar on left
            ab_id = f"{slide_id}-accent-{i}"
            elements.append(ShapeElement(
                id=ab_id, role=ElementRole.DECORATION,
                shape_type="rectangle",
                bbox=BBox(x=80, y=y + 4, w=4, h=36),
                fill_color="primary", line_color="primary",
            ))
            element_ids.append(ab_id)

            tid2 = f"{slide_id}-ta-{i}"
            elements.append(TextElement(
                id=tid2, role=ElementRole.BODY,
                text=takeaway,
                bbox=BBox(x=100, y=y, w=780, h=44),
                style={"font_size": 16},
            ))
            element_ids.append(tid2)
            y += 52

        if call_to_action:
            y = max(y + 24, 420)
            bg_id = f"{slide_id}-cta-bg"
            elements.append(ShapeElement(
                id=bg_id, role=ElementRole.DECORATION,
                shape_type="rounded_rectangle",
                bbox=BBox(x=200, y=y, w=560, h=48),
                fill_color="accent", line_color="accent",
                corner_radius=8,
            ))
            element_ids.append(bg_id)

            cta_id = f"{slide_id}-cta"
            elements.append(TextElement(
                id=cta_id, role=ElementRole.CALLOUT,
                text=call_to_action,
                bbox=BBox(x=216, y=y + 8, w=528, h=32),
                style={"font_size": 18, "bold": True, "font_color": "background", "alignment": "center"},
            ))
            element_ids.append(cta_id)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="conclusion",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
