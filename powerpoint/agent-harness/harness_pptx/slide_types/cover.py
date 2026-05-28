"""Cover slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox, LayoutChild, LayoutDirection, LayoutSpec
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class CoverSlide(SlideType):
    name = "cover"
    required_fields = ["title"]
    optional_fields = ["subtitle", "author", "date", "background_image"]
    visual_rules = {"centered": True, "max_title_lines": 3}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "cover-1")
        title = content["title"]
        subtitle = content.get("subtitle", "")
        author = content.get("author", "")
        date = content.get("date", "")
        seq = content.get("seq", 0)

        elements: list = []
        element_ids: list[str] = []

        from harness_pptx.models.element import ShapeElement

        # Bottom accent bar — anchors the slide visually
        bar_id = f"{slide_id}-accent-bar"
        elements.append(ShapeElement(
            id=bar_id, role=ElementRole.DECORATION,
            shape_type="rectangle",
            bbox=BBox(x=0, y=512, w=960, h=28),
            fill_color="primary", line_color="primary",
        ))
        element_ids.append(bar_id)

        # Vertically centered content block: center ~270, start at ~200
        # Title
        tid = f"{slide_id}-title"
        elements.append(TextElement(
            id=tid, role=ElementRole.TITLE,
            text=title,
            bbox=BBox(x=120, y=200, w=720, h=80),
            style={"font_size": 36, "bold": True, "alignment": "center"},
        ))
        element_ids.append(tid)

        y_offset = 285
        if subtitle:
            sid = f"{slide_id}-subtitle"
            elements.append(TextElement(
                id=sid, role=ElementRole.SUBTITLE,
                text=subtitle,
                bbox=BBox(x=180, y=y_offset, w=600, h=40),
                style={"font_size": 20, "font_color": "muted", "alignment": "center"},
            ))
            element_ids.append(sid)
            y_offset += 50

        if author:
            aid = f"{slide_id}-author"
            elements.append(TextElement(
                id=aid, role=ElementRole.BODY,
                text=author,
                bbox=BBox(x=240, y=y_offset, w=480, h=28),
                style={"font_size": 14, "font_color": "muted", "alignment": "center"},
            ))
            element_ids.append(aid)
            y_offset += 30

        if date:
            did = f"{slide_id}-date"
            elements.append(TextElement(
                id=did, role=ElementRole.BODY,
                text=date,
                bbox=BBox(x=240, y=y_offset, w=480, h=28),
                style={"font_size": 14, "font_color": "muted", "alignment": "center"},
            ))
            element_ids.append(did)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="cover",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
