"""Agenda slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class AgendaSlide(SlideType):
    name = "agenda"
    required_fields = ["title", "items"]
    optional_fields = ["numbered"]
    visual_rules = {"max_items": 8, "numbered": False}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "agenda-1")
        title = content.get("title", "Agenda")
        items = content["items"]
        numbered = content.get("numbered", False)
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
            bbox=BBox(x=80, y=40, w=800, h=48),
            style={"font_size": 28, "bold": True},
        ))
        element_ids.append(tid)

        # Items in a clean list
        y = 110
        for i, item in enumerate(items[:8]):
            iid = f"{slide_id}-item-{i}"
            prefix = f"{i + 1}.  " if numbered else "    "
            elements.append(TextElement(
                id=iid, role=ElementRole.BODY,
                text=f"{prefix}{item}",
                bbox=BBox(x=100, y=y, w=760, h=36),
                style={"font_size": 18},
            ))
            element_ids.append(iid)
            y += 44

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="agenda",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
