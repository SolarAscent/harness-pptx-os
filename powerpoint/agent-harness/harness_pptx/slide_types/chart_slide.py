"""Chart slide template."""

from typing import Any

from harness_pptx.models.element import ChartElement, ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ChartSlideType(SlideType):
    name = "chart"
    required_fields = ["title", "chart_data"]
    optional_fields = ["chart_type", "x_label", "y_label", "caption", "source"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "chart-1")
        title = content.get("title", "Data")
        chart_data = content["chart_data"]
        chart_type = content.get("chart_type", "bar")
        x_label = content.get("x_label", "")
        y_label = content.get("y_label", "")
        caption = content.get("caption", "")
        source = content.get("source", "")
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

        cid = f"{slide_id}-chart"
        elements.append(ChartElement(
            id=cid, role=ElementRole.BODY,
            chart_type=chart_type,
            data=chart_data if isinstance(chart_data, dict) else {"series": chart_data},
            bbox=BBox(x=80, y=100, w=800, h=380),
            title=None,
            x_label=x_label,
            y_label=y_label,
            native=False,
        ))
        element_ids.append(cid)

        if caption:
            cap_id = f"{slide_id}-caption"
            elements.append(TextElement(
                id=cap_id, role=ElementRole.CAPTION,
                text=caption,
                bbox=BBox(x=80, y=480, w=800, h=24),
                style={"font_size": 11, "font_color": "muted"},
            ))
            element_ids.append(cap_id)

        if source:
            sid = f"{slide_id}-source"
            elements.append(TextElement(
                id=sid, role=ElementRole.FOOTNOTE,
                text=f"Source: {source}",
                bbox=BBox(x=80, y=506, w=800, h=18),
                style={"font_size": 9, "font_color": "muted"},
            ))
            element_ids.append(sid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="chart",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
