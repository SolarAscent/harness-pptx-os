"""Data insight slide template — chart + callout."""

from typing import Any

from harness_pptx.models.element import ChartElement, ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class DataInsightSlide(SlideType):
    name = "data-insight"
    required_fields = ["title", "insight"]
    optional_fields = ["chart_data", "chart_type", "supporting_points", "source"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "data-insight-1")
        title = content.get("title", "Key Insight")
        insight = content["insight"]
        chart_data = content.get("chart_data")
        chart_type = content.get("chart_type", "bar")
        points = content.get("supporting_points", [])
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

        # Insight callout — prominent
        bg_id = f"{slide_id}-insight-bg"
        elements.append(ShapeElement(
            id=bg_id, role=ElementRole.DECORATION,
            shape_type="rounded_rectangle",
            bbox=BBox(x=80, y=96, w=800, h=64),
            fill_color="surface", line_color="accent",
            corner_radius=6,
        ))
        element_ids.append(bg_id)

        iid = f"{slide_id}-insight"
        elements.append(TextElement(
            id=iid, role=ElementRole.CALLOUT,
            text=insight,
            bbox=BBox(x=100, y=104, w=760, h=48),
            style={"font_size": 18, "bold": True, "font_color": "primary"},
        ))
        element_ids.append(iid)

        # Chart (if data provided)
        y = 180
        if chart_data:
            cid = f"{slide_id}-chart"
            elements.append(ChartElement(
                id=cid, role=ElementRole.BODY,
                chart_type=chart_type,
                data=chart_data if isinstance(chart_data, dict) else {"series": chart_data},
                bbox=BBox(x=100, y=y, w=400, h=280),
                native=False,
            ))
            element_ids.append(cid)

        # Supporting points (right side or below)
        px = 540 if chart_data else 80
        py = y
        pw = 340 if chart_data else 800
        for i, point in enumerate(points[:4]):
            pid = f"{slide_id}-point-{i}"
            elements.append(TextElement(
                id=pid, role=ElementRole.BODY,
                text=point,
                bbox=BBox(x=px, y=py, w=pw, h=28),
                style={"font_size": 13},
                bullet=True,
            ))
            element_ids.append(pid)
            py += 36

        if source:
            sid = f"{slide_id}-source"
            elements.append(TextElement(
                id=sid, role=ElementRole.FOOTNOTE,
                text=f"Source: {source}",
                bbox=BBox(x=80, y=500, w=800, h=20),
                style={"font_size": 9, "font_color": "muted"},
            ))
            element_ids.append(sid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="data-insight",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
