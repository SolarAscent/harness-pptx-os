"""Recommendation slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class RecommendationSlide(SlideType):
    name = "recommendation"
    required_fields = ["title", "recommendations"]
    optional_fields = ["rationale", "next_steps", "priority"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "recommendation-1")
        title = content.get("title", "Recommendations")
        recommendations = content["recommendations"]
        rationale = content.get("rationale", "")
        next_steps = content.get("next_steps", [])
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

        if rationale:
            rid = f"{slide_id}-rationale"
            elements.append(TextElement(
                id=rid, role=ElementRole.BODY,
                text=rationale,
                bbox=BBox(x=60, y=86, w=840, h=24),
                style={"font_size": 13, "font_color": "muted"},
            ))
            element_ids.append(rid)

        y = 120
        for i, rec in enumerate(recommendations[:4]):
            rec_name = rec.get("name", rec) if isinstance(rec, dict) else str(rec)
            rec_detail = rec.get("detail", "") if isinstance(rec, dict) else ""
            priority = rec.get("priority", "medium") if isinstance(rec, dict) else "medium"

            p_colors = {"high": "primary", "medium": "accent", "low": "muted"}

            # Number badge
            nbg_id = f"{slide_id}-num-bg-{i}"
            elements.append(ShapeElement(
                id=nbg_id, role=ElementRole.DECORATION,
                shape_type="rounded_rectangle",
                bbox=BBox(x=80, y=y, w=40, h=40),
                fill_color=p_colors.get(priority, "accent"),
                line_color=p_colors.get(priority, "accent"),
                corner_radius=6,
            ))
            element_ids.append(nbg_id)

            nid = f"{slide_id}-num-{i}"
            elements.append(TextElement(
                id=nid, role=ElementRole.LABEL,
                text=str(i + 1),
                bbox=BBox(x=80, y=y + 6, w=40, h=28),
                style={"font_size": 18, "bold": True, "font_color": "background", "alignment": "center"},
            ))
            element_ids.append(nid)

            rnid = f"{slide_id}-rec-{i}"
            elements.append(TextElement(
                id=rnid, role=ElementRole.BODY,
                text=rec_name,
                bbox=BBox(x=136, y=y + 4, w=740, h=28),
                style={"font_size": 16, "bold": True},
            ))
            element_ids.append(rnid)

            if rec_detail:
                rdid = f"{slide_id}-rec-d-{i}"
                elements.append(TextElement(
                    id=rdid, role=ElementRole.BODY,
                    text=rec_detail,
                    bbox=BBox(x=136, y=y + 35, w=740, h=28),
                    style={"font_size": 12, "font_color": "muted"},
                ))
                element_ids.append(rdid)

            y += 82

        if next_steps:
            y = max(y + 8, 410)
            nsl_id = f"{slide_id}-ns-label"
            elements.append(TextElement(
                id=nsl_id, role=ElementRole.LABEL,
                text="Next Steps:",
                bbox=BBox(x=80, y=y, w=140, h=24),
                style={"font_size": 13, "bold": True, "font_color": "primary"},
            ))
            element_ids.append(nsl_id)

            for j, step in enumerate(next_steps[:3]):
                nsid = f"{slide_id}-ns-{j}"
                elements.append(TextElement(
                    id=nsid, role=ElementRole.BODY,
                    text=step,
                    bbox=BBox(x=230, y=y, w=650, h=24),
                    style={"font_size": 12},
                ))
                element_ids.append(nsid)
                y += 24

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="recommendation",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
