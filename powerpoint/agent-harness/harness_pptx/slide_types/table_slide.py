"""Table slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TableElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class TableSlideType(SlideType):
    name = "table"
    required_fields = ["title", "headers", "rows"]
    optional_fields = ["caption", "source", "col_widths"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "table-1")
        title = content.get("title", "Data Table")
        headers = content["headers"]
        rows = content["rows"]
        caption = content.get("caption", "")
        source = content.get("source", "")
        col_widths = content.get("col_widths")
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

        tbl_id = f"{slide_id}-table"
        elements.append(TableElement(
            id=tbl_id, role=ElementRole.BODY,
            headers=headers,
            rows=rows,
            col_widths=col_widths,
            bbox=BBox(x=60, y=96, w=840, h=400),
            auto_width=True,
            zebra=True,
        ))
        element_ids.append(tbl_id)

        if caption:
            cap_id = f"{slide_id}-caption"
            elements.append(TextElement(
                id=cap_id, role=ElementRole.CAPTION,
                text=caption,
                bbox=BBox(x=60, y=500, w=840, h=20),
                style={"font_size": 11, "font_color": "muted"},
            ))
            element_ids.append(cap_id)

        if source:
            sid = f"{slide_id}-source"
            elements.append(TextElement(
                id=sid, role=ElementRole.FOOTNOTE,
                text=f"Source: {source}",
                bbox=BBox(x=60, y=518, w=840, h=16),
                style={"font_size": 9, "font_color": "muted"},
            ))
            element_ids.append(sid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="table",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
