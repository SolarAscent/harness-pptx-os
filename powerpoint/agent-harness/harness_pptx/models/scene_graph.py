"""SceneGraph — the resolved, render-ready tree of a complete deck.

Unlike DeckSpec (which describes *intent*), SceneGraph holds fully-resolved
BBox values for every element and maintains a flat element registry keyed by
stable id.
"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field

from harness_pptx.models.element import BaseElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.theme import Theme


class GraphEdge(BaseModel):
    """A connector edge between two elements (used for diagrams)."""

    id: str
    source_id: str
    target_id: str
    label: str | None = None
    style: dict[str, Any] = Field(default_factory=dict)


class LayerNode(BaseModel):
    """A z-ordered layer within a slide node."""

    name: str
    z_index: int = 0
    element_ids: list[str] = Field(default_factory=list)


class SlideNode(BaseModel):
    """A single slide with fully-resolved layout and element references."""

    id: str
    index: int
    slide_type: str = "custom"
    title: str = ""
    notes: str = ""
    background_color: str | None = None
    canvas: BBox = Field(default_factory=lambda: BBox(x=0, y=0, w=960, h=540))
    layers: list[LayerNode] = Field(default_factory=list)
    element_ids: list[str] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)

    def all_element_ids(self) -> list[str]:
        """Return unique element ids from both layers and flat list."""
        ids = list(self.element_ids)
        for layer in self.layers:
            ids.extend(layer.element_ids)
        return list(dict.fromkeys(ids))


class SceneGraph(BaseModel):
    """The full render-ready scene graph for a deck.

    Elements are stored in a flat registry (``element_registry``) keyed by
    stable id. Slides reference elements by id. This avoids PowerPoint shape
    index dependencies.
    """

    deck_id: str
    slide_count: int = 0
    canvas: BBox = Field(default_factory=lambda: BBox(x=0, y=0, w=960, h=540))
    theme: Theme = Field(default_factory=Theme)
    slides: dict[str, SlideNode] = Field(default_factory=dict)
    element_registry: dict[str, BaseElement] = Field(default_factory=dict)
    edges: list[GraphEdge] = Field(default_factory=list)

    def get_element(self, element_id: str) -> BaseElement | None:
        return self.element_registry.get(element_id)

    def register_element(self, element: BaseElement) -> None:
        self.element_registry[element.id] = element

    def slide_order(self) -> list[SlideNode]:
        """Return slides sorted by index."""
        return sorted(self.slides.values(), key=lambda s: s.index)
