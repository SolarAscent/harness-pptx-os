"""Appendix slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class AppendixSlide(SlideType):
    name = "appendix"
    required_fields = ["title", "content"]
    optional_fields = ["type", "reference", "section_label"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "appendix-1")
        title = content.get("title", "Appendix")
        appendix_content = content["content"]
        appendix_type = content.get("type", "text")
        reference = content.get("reference", "")
        section_label = content.get("section_label", "Appendix")
        seq = content.get("seq", 0)

        elements: list = []
        element_ids: list[str] = []

        # Section label
        lid = f"{slide_id}-label"
        elements.append(TextElement(
            id=lid, role=ElementRole.KICKER,
            text=section_label,
            bbox=BBox(x=60, y=36, w=200, h=24),
            style={"font_size": 12, "font_color": "muted", "bold": True},
        ))
        element_ids.append(lid)

        tid = f"{slide_id}-title"
        elements.append(TextElement(
            id=tid, role=ElementRole.TITLE,
            text=title,
            bbox=BBox(x=60, y=64, w=840, h=36),
            style={"font_size": 22, "bold": True},
        ))
        element_ids.append(tid)

        y = 120
        if isinstance(appendix_content, list):
            for i, item in enumerate(appendix_content):
                iid = f"{slide_id}-item-{i}"
                elements.append(TextElement(
                    id=iid, role=ElementRole.BODY,
                    text=item,
                    bbox=BBox(x=80, y=y, w=800, h=28),
                    style={"font_size": 13},
                    bullet=appendix_type == "bullets",
                ))
                element_ids.append(iid)
                y += 36
        else:
            cid = f"{slide_id}-content"
            elements.append(TextElement(
                id=cid, role=ElementRole.BODY,
                text=str(appendix_content),
                bbox=BBox(x=80, y=y, w=800, h=360),
                style={"font_size": 13},
            ))
            element_ids.append(cid)

        if reference:
            rid = f"{slide_id}-ref"
            elements.append(TextElement(
                id=rid, role=ElementRole.FOOTNOTE,
                text=f"Reference: {reference}",
                bbox=BBox(x=80, y=500, w=800, h=20),
                style={"font_size": 9, "font_color": "muted"},
            ))
            element_ids.append(rid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="appendix",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
