"""Workflow slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class WorkflowSlide(SlideType):
    name = "workflow"
    required_fields = ["title", "steps"]
    optional_fields = ["direction", "roles"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "workflow-1")
        title = content.get("title", "Workflow")
        steps = content["steps"]
        seq = content.get("seq", 0)

        elements: list = []
        element_ids: list[str] = []

        tid = f"{slide_id}-title"
        elements.append(TextElement(
            id=tid, role=ElementRole.TITLE,
            text=title,
            bbox=BBox(x=60, y=36, w=840, h=44),
            style={"font_size": 26, "bold": True},
        ))
        element_ids.append(tid)

        n = min(len(steps), 5)
        box_w = 800 // n - 20
        y = 140

        for i, step in enumerate(steps[:5]):
            cx = 80 + i * (box_w + 20)
            step_name = step.get("name", step) if isinstance(step, dict) else str(step)
            step_desc = step.get("description", "") if isinstance(step, dict) else ""

            # Step box
            sbid = f"{slide_id}-step-bg-{i}"
            elements.append(ShapeElement(
                id=sbid, role=ElementRole.DECORATION,
                shape_type="rounded_rectangle",
                bbox=BBox(x=cx, y=y, w=box_w, h=44),
                fill_color="primary", line_color="primary",
                corner_radius=6,
            ))
            element_ids.append(sbid)

            # Step number
            snid = f"{slide_id}-step-n-{i}"
            elements.append(TextElement(
                id=snid, role=ElementRole.LABEL,
                text=str(i + 1),
                bbox=BBox(x=cx + 4, y=y + 8, w=24, h=28),
                style={"font_size": 18, "bold": True, "font_color": "background"},
            ))
            element_ids.append(snid)

            # Step name
            sid = f"{slide_id}-step-{i}"
            elements.append(TextElement(
                id=sid, role=ElementRole.BODY,
                text=step_name,
                bbox=BBox(x=cx + 32, y=y + 10, w=box_w - 36, h=24),
                style={"font_size": 13, "bold": True, "font_color": "background"},
            ))
            element_ids.append(sid)

            # Description below
            if step_desc:
                sdid = f"{slide_id}-desc-{i}"
                elements.append(TextElement(
                    id=sdid, role=ElementRole.BODY,
                    text=step_desc,
                    bbox=BBox(x=cx, y=y + 56, w=box_w, h=60),
                    style={"font_size": 11, "alignment": "center"},
                ))
                element_ids.append(sdid)

            # Arrow connector
            if i < n - 1:
                aid = f"{slide_id}-arrow-{i}"
                elements.append(TextElement(
                    id=aid, role=ElementRole.DECORATION,
                    text=" → ",
                    bbox=BBox(x=cx + box_w, y=y + 8, w=20, h=28),
                    style={"font_size": 18, "font_color": "border", "alignment": "center"},
                ))
                element_ids.append(aid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="workflow",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
