"""Case study slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class CaseStudySlide(SlideType):
    name = "case-study"
    required_fields = ["title"]
    optional_fields = ["company", "challenge", "solution", "results", "quote"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "case-study-1")
        title = content.get("title", "Case Study")
        company = content.get("company", "")
        challenge = content.get("challenge", "")
        solution = content.get("solution", "")
        results = content.get("results", [])
        quote = content.get("quote", "")
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
        if company:
            cid = f"{slide_id}-company"
            elements.append(TextElement(
                id=cid, role=ElementRole.SUBTITLE,
                text=company,
                bbox=BBox(x=60, y=y, w=840, h=28),
                style={"font_size": 16, "bold": True, "font_color": "accent"},
            ))
            element_ids.append(cid)
            y += 36

        # Two-column layout
        for label, text, prefix in [
            ("Challenge", challenge, "ch"),
            ("Solution", solution, "sol"),
        ]:
            if text:
                lbl_id = f"{slide_id}-{prefix}-label"
                elements.append(TextElement(
                    id=lbl_id, role=ElementRole.LABEL,
                    text=label,
                    bbox=BBox(x=60, y=y, w=120, h=24),
                    style={"font_size": 13, "bold": True, "font_color": "primary"},
                ))
                element_ids.append(lbl_id)

                txt_id = f"{slide_id}-{prefix}-text"
                elements.append(TextElement(
                    id=txt_id, role=ElementRole.BODY,
                    text=text,
                    bbox=BBox(x=60, y=y + 26, w=840, h=48),
                    style={"font_size": 13},
                ))
                element_ids.append(txt_id)
                y += 84

        if quote:
            qbg_id = f"{slide_id}-quote-bg"
            elements.append(ShapeElement(
                id=qbg_id, role=ElementRole.DECORATION,
                shape_type="rounded_rectangle",
                bbox=BBox(x=80, y=y + 8, w=800, h=52),
                fill_color="surface", line_color="border",
                corner_radius=4,
            ))
            element_ids.append(qbg_id)

            qid = f"{slide_id}-quote"
            elements.append(TextElement(
                id=qid, role=ElementRole.BODY,
                text=f'"{quote}"',
                bbox=BBox(x=100, y=y + 16, w=760, h=36),
                style={"font_size": 14, "italic": True},
            ))
            element_ids.append(qid)
            y += 72

        if results:
            rid = f"{slide_id}-results-label"
            elements.append(TextElement(
                id=rid, role=ElementRole.LABEL,
                text="Results",
                bbox=BBox(x=60, y=y, w=120, h=24),
                style={"font_size": 13, "bold": True, "font_color": "success"},
            ))
            element_ids.append(rid)
            y += 28

            for i, r in enumerate(results[:4]):
                riid = f"{slide_id}-result-{i}"
                elements.append(TextElement(
                    id=riid, role=ElementRole.BODY,
                    text=r,
                    bbox=BBox(x=80, y=y, w=820, h=24),
                    style={"font_size": 13},
                    bullet=True,
                ))
                element_ids.append(riid)
                y += 28

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="case-study",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
