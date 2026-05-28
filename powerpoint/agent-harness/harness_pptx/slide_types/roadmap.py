"""Roadmap slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class RoadmapSlide(SlideType):
    name = "roadmap"
    required_fields = ["title", "phases"]
    optional_fields = ["timeline_label"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "roadmap-1")
        title = content.get("title", "Roadmap")
        phases = content["phases"]
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

        n = min(len(phases), 5)
        box_w = 800 // n - 12
        y0 = 130

        # Timeline bar
        bar_id = f"{slide_id}-bar"
        elements.append(ShapeElement(
            id=bar_id, role=ElementRole.DECORATION,
            shape_type="rectangle",
            bbox=BBox(x=80, y=120, w=800, h=6),
            fill_color="accent", line_color="accent",
            corner_radius=3,
        ))
        element_ids.append(bar_id)

        for i, phase in enumerate(phases[:5]):
            cx = 80 + i * (box_w + 12)

            # Phase box
            pb_id = f"{slide_id}-phase-bg-{i}"
            elements.append(ShapeElement(
                id=pb_id, role=ElementRole.DECORATION,
                shape_type="rounded_rectangle",
                bbox=BBox(x=cx, y=y0, w=box_w, h=60),
                fill_color="surface", line_color="border",
                corner_radius=6,
            ))
            element_ids.append(pb_id)

            # Phase name
            phase_name = phase.get("name", phase) if isinstance(phase, dict) else str(phase)
            pn_id = f"{slide_id}-phase-{i}"
            elements.append(TextElement(
                id=pn_id, role=ElementRole.SUBTITLE,
                text=phase_name,
                bbox=BBox(x=cx + 8, y=y0 + 10, w=box_w - 16, h=40),
                style={"font_size": 13, "bold": True, "alignment": "center"},
            ))
            element_ids.append(pn_id)

            # Timeline label
            if isinstance(phase, dict) and "date" in phase:
                dt_id = f"{slide_id}-date-{i}"
                elements.append(TextElement(
                    id=dt_id, role=ElementRole.CAPTION,
                    text=phase["date"],
                    bbox=BBox(x=cx, y=100, w=box_w, h=18),
                    style={"font_size": 10, "font_color": "muted", "alignment": "center"},
                ))
                element_ids.append(dt_id)

            # Items below
            if isinstance(phase, dict) and "items" in phase:
                y2 = y0 + 80
                for j, item in enumerate(phase["items"][:4]):
                    it_id = f"{slide_id}-item-{i}-{j}"
                    elements.append(TextElement(
                        id=it_id, role=ElementRole.BODY,
                        text=item,
                        bbox=BBox(x=cx + 4, y=y2, w=box_w - 8, h=20),
                        style={"font_size": 10},
                        bullet=True,
                    ))
                    element_ids.append(it_id)
                    y2 += 22

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="roadmap",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
