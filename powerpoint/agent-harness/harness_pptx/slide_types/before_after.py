"""Before/after slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class BeforeAfterSlide(SlideType):
    name = "before-after"
    required_fields = ["title", "before_points", "after_points"]
    optional_fields = ["before_label", "after_label", "before_image", "after_image"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "before-after-1")
        title = content.get("title", "Before & After")
        before = content["before_points"]
        after = content["after_points"]
        before_label = content.get("before_label", "Before")
        after_label = content.get("after_label", "After")
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

        col_w = 380
        for col_x, items, label, prefix, color in [
            (80, before, before_label, "before", "muted"),
            (500, after, after_label, "after", "success"),
        ]:
            hdr_id = f"{slide_id}-hdr-{prefix}"
            elements.append(ShapeElement(
                id=hdr_id, role=ElementRole.DECORATION,
                shape_type="rounded_rectangle",
                bbox=BBox(x=col_x, y=96, w=col_w, h=36),
                fill_color="surface", line_color="border",
                corner_radius=4,
            ))
            element_ids.append(hdr_id)

            htxt_id = f"{slide_id}-htxt-{prefix}"
            elements.append(TextElement(
                id=htxt_id, role=ElementRole.SUBTITLE,
                text=label,
                bbox=BBox(x=col_x + 8, y=100, w=col_w - 16, h=28),
                style={"font_size": 16, "bold": True, "font_color": color, "alignment": "center"},
            ))
            element_ids.append(htxt_id)

            y = 148
            for i, item in enumerate(items[:6]):
                iid = f"{slide_id}-{prefix}-{i}"
                elements.append(TextElement(
                    id=iid, role=ElementRole.BODY,
                    text=str(item),
                    bbox=BBox(x=col_x + 8, y=y, w=col_w - 16, h=28),
                    style={"font_size": 13},
                    bullet=True,
                ))
                element_ids.append(iid)
                y += 32

        # Arrow between
        arr_id = f"{slide_id}-arrow"
        elements.append(TextElement(
            id=arr_id, role=ElementRole.DECORATION,
            text=" → ",
            bbox=BBox(x=450, y=260, w=60, h=36),
            style={"font_size": 28, "font_color": "accent", "alignment": "center"},
        ))
        element_ids.append(arr_id)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="before-after",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
