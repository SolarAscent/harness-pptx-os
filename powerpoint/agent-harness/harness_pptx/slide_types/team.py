"""Team slide template."""

from typing import Any

from harness_pptx.models.element import ElementRole, TextElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.scene_graph import LayerNode, SlideNode
from harness_pptx.models.theme import Theme
from harness_pptx.slide_types.base import SlideType, register_element


class TeamSlide(SlideType):
    name = "team"
    required_fields = ["title", "members"]
    optional_fields = ["layout"]

    def build(self, content: dict[str, Any], theme: Theme, registry: dict[str, BaseElement] | None = None) -> SlideNode:
        slide_id = content.get("slide_id", "team-1")
        title = content.get("title", "Our Team")
        members = content["members"]
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

        cols = min(len(members), 4)
        col_w = 800 // cols - 16
        for i, member in enumerate(members[:8]):
            r = i // cols
            c = i % cols
            cx = 80 + c * (col_w + 16)
            cy = 110 + r * 180

            # Placeholder / name
            name = member.get("name", member) if isinstance(member, dict) else str(member)
            nid = f"{slide_id}-name-{i}"
            elements.append(TextElement(
                id=nid, role=ElementRole.SUBTITLE,
                text=name,
                bbox=BBox(x=cx, y=cy, w=col_w, h=28),
                style={"font_size": 15, "bold": True, "alignment": "center"},
            ))
            element_ids.append(nid)

            if isinstance(member, dict):
                title_txt = member.get("title", member.get("role", ""))
                if title_txt:
                    rid = f"{slide_id}-role-{i}"
                    elements.append(TextElement(
                        id=rid, role=ElementRole.CAPTION,
                        text=title_txt,
                        bbox=BBox(x=cx, y=cy + 30, w=col_w, h=24),
                        style={"font_size": 12, "font_color": "muted", "alignment": "center"},
                    ))
                    element_ids.append(rid)

                bio = member.get("bio", "")
                if bio:
                    bid = f"{slide_id}-bio-{i}"
                    elements.append(TextElement(
                        id=bid, role=ElementRole.BODY,
                        text=bio,
                        bbox=BBox(x=cx, y=cy + 60, w=col_w, h=60),
                        style={"font_size": 10, "alignment": "center"},
                    ))
                    element_ids.append(bid)

        for el in elements:
            register_element(registry, el)

        
        layer = LayerNode(name="main", z_index=0, element_ids=element_ids)
        return SlideNode(
            id=slide_id, index=seq, slide_type="team",
            title=title,
            layers=[layer],
            element_ids=element_ids,
        )
