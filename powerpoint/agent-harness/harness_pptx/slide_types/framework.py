"""Framework slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class FrameworkSlide(SlideType):
    name = "framework"
    required_fields = ["title", "framework_name"]
    optional_fields = ["components", "description", "diagram_ref"]
    visual_rules = {"max_components": 8}

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "framework-1")
        title = content.get("title", "Framework")
        fw_name = content["framework_name"]
        components = content.get("components", [])
        description = content.get("description", "")
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

        # Framework name badge
        bid = f"{slide_id}-name"
        elements.append(ShapeElement(
            id=bid, role=ElementRole.DECORATION,
            shape_type="rounded_rectangle",
            bbox=BBox(x=60, y=96, w=840, h=44),
            fill_color="primary", line_color="primary",
            corner_radius=4,
            text=fw_name,
        ))
        element_ids.append(bid)

        fw_title_id = f"{slide_id}-fw-title"
        elements.append(TextElement(
            id=fw_title_id, role=ElementRole.SUBTITLE,
            text=fw_name,
            bbox=BBox(x=80, y=100, w=800, h=36),
            style={"font_size": 18, "bold": True, "font_color": "background"},
        ))
        element_ids.append(fw_title_id)

        if description:
            did = f"{slide_id}-desc"
            elements.append(TextElement(
                id=did, role=ElementRole.BODY,
                text=description,
                bbox=BBox(x=60, y=152, w=840, h=36),
                style={"font_size": 14, "font_color": "muted"},
            ))
            element_ids.append(did)

        # Components as a grid of boxes
        n = len(components)
        if n > 0:
            cols = min(n, 4)
            rows = (n + cols - 1) // cols
            box_w = 800 // cols - 16
            box_h = min((380 - 200) // max(rows, 1), 80)
            y0 = 200

            for i, comp in enumerate(components):
                r, c = divmod(i, cols)
                cx = 80 + c * (box_w + 16)
                cy = y0 + r * (box_h + 12)

                cbg_id = f"{slide_id}-comp-bg-{i}"
                elements.append(ShapeElement(
                    id=cbg_id, role=ElementRole.DECORATION,
                    shape_type="rounded_rectangle",
                    bbox=BBox(x=cx, y=cy, w=box_w, h=box_h),
                    fill_color="surface", line_color="border",
                    corner_radius=4,
                ))
                element_ids.append(cbg_id)

                ct_id = f"{slide_id}-comp-{i}"
                elements.append(TextElement(
                    id=ct_id, role=ElementRole.BODY,
                    text=comp.get("title", comp.get("name", str(comp))) if isinstance(comp, dict) else str(comp),
                    bbox=BBox(x=cx + 8, y=cy + 4, w=box_w - 16, h=box_h - 8),
                    style={"font_size": 12, "alignment": "center"},
                ))
                element_ids.append(ct_id)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="framework",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
