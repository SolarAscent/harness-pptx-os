"""Solution slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class SolutionSlide(SlideType):
    name = "solution"
    required_fields = ["title", "solution_summary"]
    optional_fields = ["key_features", "how_it_works", "benefits"]
    visual_rules = {"positive_tone": True}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "solution-1")
        title = content.get("title", "Our Solution")
        summary = content["solution_summary"]
        features = content.get("key_features", [])
        how_it_works = content.get("how_it_works", "")
        benefits = content.get("benefits", [])
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
            bbox=BBox(x=80, y=40, w=800, h=44),
            style={"font_size": 26, "bold": True},
        ))
        element_ids.append(tid)

        bg_id = f"{slide_id}-callout-bg"
        elements.append(ShapeElement(
            id=bg_id, role=ElementRole.DECORATION,
            shape_type="rounded_rectangle",
            bbox=BBox(x=80, y=104, w=800, h=72),
            fill_color="surface", line_color="accent",
            corner_radius=6,
        ))
        element_ids.append(bg_id)

        sid = f"{slide_id}-summary"
        elements.append(TextElement(
            id=sid, role=ElementRole.CALLOUT,
            text=summary,
            bbox=BBox(x=100, y=114, w=760, h=52),
            style={"font_size": 18, "bold": True, "font_color": "primary"},
        ))
        element_ids.append(sid)

        y = 200
        if features:
            for i, f in enumerate(features[:5]):
                fid = f"{slide_id}-feat-{i}"
                elements.append(TextElement(
                    id=fid, role=ElementRole.BODY,
                    text=f,
                    bbox=BBox(x=100, y=y, w=760, h=28),
                    style={"font_size": 15},
                    bullet=True,
                ))
                element_ids.append(fid)
                y += 38

        if how_it_works:
            hid = f"{slide_id}-how"
            y = max(y + 12, 400)
            elements.append(TextElement(
                id=hid, role=ElementRole.BODY,
                text=f"How it works: {how_it_works}",
                bbox=BBox(x=80, y=y, w=800, h=32),
                style={"font_size": 14, "font_color": "muted"},
            ))
            element_ids.append(hid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="solution",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
