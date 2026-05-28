"""Process slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ProcessSlide(SlideType):
    name = "process"
    required_fields = ["title", "steps"]
    optional_fields = ["orientation"]
    visual_rules = {"max_steps": 6, "numbered": True}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "process-1")
        title = content.get("title", "Process")
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

        n = min(len(steps), 6)
        box_w = 800 // max(n, 1) - 16
        x_start = 80

        for i, step in enumerate(steps[:6]):
            cx = x_start + i * (box_w + 16)
            y = 120

            # Step number
            num_id = f"{slide_id}-num-{i}"
            elements.append(ShapeElement(
                id=num_id, role=ElementRole.DECORATION,
                shape_type="circle",
                bbox=BBox(x=cx + (box_w - 36) // 2, y=y, w=36, h=36),
                fill_color="accent", line_color="accent",
                text=str(i + 1),
            ))
            element_ids.append(num_id)

            # Step title
            st_id = f"{slide_id}-step-t-{i}"
            elements.append(TextElement(
                id=st_id, role=ElementRole.SUBTITLE,
                text=step.get("title", f"Step {i+1}"),
                bbox=BBox(x=cx, y=y + 46, w=box_w, h=28),
                style={"font_size": 14, "bold": True, "alignment": "center"},
            ))
            element_ids.append(st_id)

            # Step description
            sd_id = f"{slide_id}-step-d-{i}"
            elements.append(TextElement(
                id=sd_id, role=ElementRole.BODY,
                text=step.get("description", step if isinstance(step, str) else ""),
                bbox=BBox(x=cx, y=y + 80, w=box_w, h=80),
                style={"font_size": 11, "alignment": "center"},
            ))
            element_ids.append(sd_id)

            # Arrow between steps (except last)
            if i < n - 1:
                arr_id = f"{slide_id}-arrow-{i}"
                elements.append(ShapeElement(
                    id=arr_id, role=ElementRole.DECORATION,
                    shape_type="arrow",
                    bbox=BBox(x=cx + box_w + 2, y=y + 12, w=10, h=12),
                    fill_color="border", line_color="border",
                ))
                element_ids.append(arr_id)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="process",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
