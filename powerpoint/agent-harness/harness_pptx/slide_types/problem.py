"""Problem slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ProblemSlide(SlideType):
    name = "problem"
    required_fields = ["title", "problem_statement"]
    optional_fields = ["pain_points", "impact", "context"]
    visual_rules = {"emphasis": True}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "problem-1")
        title = content.get("title", "The Problem")
        statement = content["problem_statement"]
        pain_points = content.get("pain_points", [])
        impact = content.get("impact", "")
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

        # Problem statement callout
        bg_id = f"{slide_id}-callout-bg"
        elements.append(ShapeElement(
            id=bg_id, role=ElementRole.DECORATION,
            shape_type="rounded_rectangle",
            bbox=BBox(x=80, y=104, w=800, h=80),
            fill_color="surface", line_color="border",
            corner_radius=6,
        ))
        element_ids.append(bg_id)

        sid = f"{slide_id}-statement"
        elements.append(TextElement(
            id=sid, role=ElementRole.CALLOUT,
            text=statement,
            bbox=BBox(x=100, y=114, w=760, h=60),
            style={"font_size": 18, "bold": True, "font_color": "error"},
        ))
        element_ids.append(sid)

        y = 210
        for i, p in enumerate(pain_points[:4]):
            pid = f"{slide_id}-pain-{i}"
            elements.append(TextElement(
                id=pid, role=ElementRole.BODY,
                text=p,
                bbox=BBox(x=100, y=y, w=760, h=28),
                style={"font_size": 15},
                bullet=True,
            ))
            element_ids.append(pid)
            y += 40

        if impact:
            iid = f"{slide_id}-impact"
            elements.append(TextElement(
                id=iid, role=ElementRole.CALLOUT,
                text=f"Impact: {impact}",
                bbox=BBox(x=80, y=y + 12, w=800, h=32),
                style={"font_size": 15, "font_color": "muted"},
            ))
            element_ids.append(iid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="problem",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
