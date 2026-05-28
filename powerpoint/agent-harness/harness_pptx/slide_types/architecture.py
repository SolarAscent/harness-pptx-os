"""Architecture diagram slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, ShapeElement, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ArchitectureSlide(SlideType):
    name = "architecture"
    required_fields = ["title"]
    optional_fields = ["layers", "components", "description", "diagram_ref"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "architecture-1")
        title = content.get("title", "Architecture")
        layers = content.get("layers", [])
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

        if description:
            did = f"{slide_id}-desc"
            elements.append(TextElement(
                id=did, role=ElementRole.BODY,
                text=description,
                bbox=BBox(x=60, y=86, w=840, h=24),
                style={"font_size": 13, "font_color": "muted"},
            ))
            element_ids.append(did)

        # Stacked layers
        if layers:
            n = len(layers)
            layer_h = min((420) // max(n, 1), 80)
            y0 = 110
            for i, layer in enumerate(layers):
                ly = y0 + i * (layer_h + 8)
                layer_name = layer.get("name", f"Layer {i+1}") if isinstance(layer, dict) else str(layer)
                layer_color = ["primary", "accent", "surface", "muted"][i % 4] if isinstance(layer, dict) else "surface"

                lbg_id = f"{slide_id}-layer-bg-{i}"
                elements.append(ShapeElement(
                    id=lbg_id, role=ElementRole.DECORATION,
                    shape_type="rounded_rectangle",
                    bbox=BBox(x=100, y=ly, w=760, h=layer_h),
                    fill_color=layer_color if layer_color != "surface" else "surface",
                    line_color="border",
                    corner_radius=4,
                ))
                element_ids.append(lbg_id)

                ln_id = f"{slide_id}-layer-{i}"
                elements.append(TextElement(
                    id=ln_id, role=ElementRole.LABEL,
                    text=layer_name,
                    bbox=BBox(x=116, y=ly + layer_h // 2 - 14, w=728, h=28),
                    style={"font_size": 14, "bold": True, "font_color": "background" if layer_color != "surface" else "text"},
                ))
                element_ids.append(ln_id)

                if isinstance(layer, dict) and "sub_components" in layer:
                    y_sub = ly + layer_h // 2 + 8
                    # Simple inline sub-component labels
                    subs = ",  ".join(layer["sub_components"][:5])
                    sc_id = f"{slide_id}-sub-{i}"
                    elements.append(TextElement(
                        id=sc_id, role=ElementRole.FOOTNOTE,
                        text=subs,
                        bbox=BBox(x=116, y=y_sub, w=728, h=20),
                        style={"font_size": 10, "font_color": "muted"},
                    ))
                    element_ids.append(sc_id)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="architecture",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
