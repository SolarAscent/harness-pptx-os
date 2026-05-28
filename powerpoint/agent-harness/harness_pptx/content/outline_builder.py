"""OutlineBuilder — refine and finalize outline structure."""

from __future__ import annotations

from harness_pptx.models.content import Outline


class OutlineBuilder:
    """Refine and validate an Outline."""

    def build(self, outline: Outline, max_slides: int = 50) -> Outline:
        if outline.total_slides > max_slides:
            outline.items = outline.items[:max_slides]
            outline.total_slides = len(outline.items)

        # Ensure sequential numbering
        for i, item in enumerate(outline.items):
            item.seq = i

        return outline
