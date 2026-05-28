"""Risk slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class RiskSlide(SlideType):
    name = "risk"
    required_fields = ["title", "risks"]
    optional_fields = ["mitigation", "risk_matrix"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "risk-1")
        title = content.get("title", "Risks & Mitigation")
        risks = content["risks"]
        mitigation = content.get("mitigation", "")
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

        y = 96
        for i, risk in enumerate(risks[:6]):
            risk_name = risk.get("name", risk) if isinstance(risk, dict) else str(risk)
            level = risk.get("level", "medium") if isinstance(risk, dict) else "medium"
            desc = risk.get("description", "") if isinstance(risk, dict) else ""

            level_colors = {"high": "error", "medium": "warning", "low": "success"}

            # Indicator dot
            dot_id = f"{slide_id}-dot-{i}"
            elements.append(ShapeElement(
                id=dot_id, role=ElementRole.DECORATION,
                shape_type="circle",
                bbox=BBox(x=80, y=y + 6, w=12, h=12),
                fill_color=level_colors.get(level, "warning"),
                line_color=level_colors.get(level, "warning"),
            ))
            element_ids.append(dot_id)

            rid = f"{slide_id}-risk-{i}"
            elements.append(TextElement(
                id=rid, role=ElementRole.BODY,
                text=risk_name,
                bbox=BBox(x=104, y=y, w=760, h=24),
                style={"font_size": 15, "bold": True},
            ))
            element_ids.append(rid)

            if desc:
                rdid = f"{slide_id}-risk-desc-{i}"
                elements.append(TextElement(
                    id=rdid, role=ElementRole.BODY,
                    text=desc,
                    bbox=BBox(x=104, y=y + 24, w=760, h=20),
                    style={"font_size": 11, "font_color": "muted"},
                ))
                element_ids.append(rdid)

            y += 60

        if mitigation:
            y = max(y + 16, 430)
            mid = f"{slide_id}-mitigation"
            elements.append(TextElement(
                id=mid, role=ElementRole.CALLOUT,
                text=f"Mitigation: {mitigation}",
                bbox=BBox(x=80, y=y, w=800, h=28),
                style={"font_size": 14, "font_color": "primary"},
            ))
            element_ids.append(mid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="risk",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
