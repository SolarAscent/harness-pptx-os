"""Section divider slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class SectionDividerSlide(SlideType):
    name = "section-divider"
    required_fields = ["title"]
    optional_fields = ["section_number", "subtitle", "background_color"]
    visual_rules = {"centered": True, "full_bleed": True}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "section-1")
        title = content["title"]
        section_num = content.get("section_number", "")
        subtitle = content.get("subtitle", "")
        bg_color = content.get("background_color", "primary")
        seq = content.get("seq", 0)

        elements: list = []
        element_ids: list[str] = []

        # Full-slide background band
        bg_id = f"{slide_id}-bg"
        elements.append(ShapeElement(
            id=bg_id, role=ElementRole.DECORATION,
            shape_type="rectangle",
            bbox=BBox(x=0, y=0, w=960, h=540),
            fill_color=bg_color, line_color=bg_color,
        ))
        element_ids.append(bg_id)

        y_center = 220
        if section_num:
            nid = f"{slide_id}-num"
            elements.append(TextElement(
                id=nid, role=ElementRole.KICKER,
                text=str(section_num),
                bbox=BBox(x=100, y=y_center, w=760, h=40),
                style={"font_size": 16, "font_color": "background", "alignment": "center"},
            ))
            element_ids.append(nid)
            y_center += 48

        tid = f"{slide_id}-title"
        elements.append(TextElement(
            id=tid, role=ElementRole.TITLE,
            text=title,
            bbox=BBox(x=80, y=y_center, w=800, h=80),
            style={"font_size": 34, "bold": True, "font_color": "background", "alignment": "center"},
        ))
        element_ids.append(tid)

        if subtitle:
            sid = f"{slide_id}-subtitle"
            elements.append(TextElement(
                id=sid, role=ElementRole.SUBTITLE,
                text=subtitle,
                bbox=BBox(x=180, y=y_center + 90, w=600, h=36),
                style={"font_size": 16, "font_color": "background", "alignment": "center"},
            ))
            element_ids.append(sid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="section-divider",
            title=title, background_color=bg_color,
            layers=[layer],
            element_ids=element_ids,
        )
