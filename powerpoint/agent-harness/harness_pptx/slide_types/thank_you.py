"""Thank you slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class ThankYouSlide(SlideType):
    name = "thank-you"
    required_fields = []
    optional_fields = ["message", "contact", "email", "website", "social", "subtitle"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "thank-you-1")
        message = content.get("message", "Thank You")
        contact = content.get("contact", "")
        email = content.get("email", "")
        website = content.get("website", "")
        subtitle = content.get("subtitle", "")
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

        # Vertically centered: canvas is 540, optical center ~250
        # Main thank you message — large and bold
        mid = f"{slide_id}-msg"
        elements.append(TextElement(
            id=mid, role=ElementRole.TITLE,
            text=message,
            bbox=BBox(x=80, y=190, w=800, h=80),
            style={"font_size": 44, "bold": True, "alignment": "center"},
        ))
        element_ids.append(mid)

        y = 285
        if subtitle:
            sid = f"{slide_id}-sub"
            elements.append(TextElement(
                id=sid, role=ElementRole.SUBTITLE,
                text=subtitle,
                bbox=BBox(x=80, y=y, w=800, h=36),
                style={"font_size": 18, "font_color": "muted", "alignment": "center"},
            ))
            element_ids.append(sid)
            y += 46

        contact_lines = []
        if contact:
            contact_lines.append(contact)
        if email:
            contact_lines.append(email)
        if website:
            contact_lines.append(website)

        for i, line in enumerate(contact_lines):
            cid = f"{slide_id}-contact-{i}"
            elements.append(TextElement(
                id=cid, role=ElementRole.BODY,
                text=line,
                bbox=BBox(x=80, y=y, w=800, h=28),
                style={"font_size": 16, "font_color": "muted", "alignment": "center"},
            ))
            element_ids.append(cid)
            y += 32

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="thank-you",
            title=message,
            layers=[layer],
            element_ids=element_ids,
        )
