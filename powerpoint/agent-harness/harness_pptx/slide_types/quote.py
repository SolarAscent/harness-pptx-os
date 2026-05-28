"""Quote slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class QuoteSlide(SlideType):
    name = "quote"
    required_fields = ["quote_text"]
    optional_fields = ["attribution", "role", "context"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "quote-1")
        quote_text = content["quote_text"]
        attribution = content.get("attribution", "")
        role = content.get("role", "")
        context = content.get("context", "")
        seq = content.get("seq", 0)

        elements: list = []
        element_ids: list[str] = []

        # Large quote mark
        qm_id = f"{slide_id}-qm"
        elements.append(TextElement(
            id=qm_id, role=ElementRole.DECORATION,
            text='"',
            bbox=BBox(x=120, y=80, w=60, h=80),
            style={"font_size": 72, "font_color": "accent"},
        ))
        element_ids.append(qm_id)

        # Quote text
        qid = f"{slide_id}-quote"
        quote_display = f'"{quote_text}"'
        elements.append(TextElement(
            id=qid, role=ElementRole.BODY,
            text=quote_display,
            bbox=BBox(x=160, y=120, w=680, h=160),
            style={"font_size": 24, "italic": True, "alignment": "left"},
        ))
        element_ids.append(qid)

        # Accent line
        line_id = f"{slide_id}-line"
        elements.append(ShapeElement(
            id=line_id, role=ElementRole.DECORATION,
            shape_type="rectangle",
            bbox=BBox(x=160, y=300, w=80, h=3),
            fill_color="accent", line_color="accent",
        ))
        element_ids.append(line_id)

        # Attribution
        y = 320
        if attribution:
            aid = f"{slide_id}-attr"
            elements.append(TextElement(
                id=aid, role=ElementRole.BODY,
                text=attribution,
                bbox=BBox(x=160, y=y, w=640, h=28),
                style={"font_size": 16, "bold": True},
            ))
            element_ids.append(aid)
            y += 28

        if role:
            rid = f"{slide_id}-role"
            elements.append(TextElement(
                id=rid, role=ElementRole.CAPTION,
                text=role,
                bbox=BBox(x=160, y=y, w=640, h=24),
                style={"font_size": 13, "font_color": "muted"},
            ))
            element_ids.append(rid)
            y += 24

        if context:
            cid = f"{slide_id}-context"
            elements.append(TextElement(
                id=cid, role=ElementRole.FOOTNOTE,
                text=context,
                bbox=BBox(x=80, y=480, w=800, h=24),
                style={"font_size": 11, "font_color": "muted", "alignment": "center"},
            ))
            element_ids.append(cid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="quote",
            title=attribution,
            layers=[layer],
            element_ids=element_ids,
        )
