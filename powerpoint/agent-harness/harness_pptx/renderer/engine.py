"""DeckRenderer — SceneGraph to PowerPoint rendering pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from harness_pptx.backends.interface import ElementHandle, RendererInterface, SlideHandle
from harness_pptx.models.scene_graph import SceneGraph, SlideNode
from harness_pptx.renderer.element_renderer import ElementRenderer


@dataclass
class RenderResult:
    """Result of rendering a scene graph to PowerPoint."""

    deck_id: str
    success: bool = True
    slides_rendered: int = 0
    elements_rendered: int = 0
    failed_slides: list[str] = field(default_factory=list)
    failed_elements: list[str] = field(default_factory=list)
    element_map: dict[str, ElementHandle] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)


class DeckRenderer:
    """Renders a SceneGraph to PowerPoint using a backend."""

    def __init__(self, backend: RendererInterface):
        self._backend = backend
        self._element_renderer: ElementRenderer | None = None

    def render(self, scene_graph: SceneGraph, output_path: str) -> RenderResult:
        """Render complete scene graph to a .pptx file."""
        result = RenderResult(deck_id=scene_graph.deck_id)

        # Create element renderer with theme for color resolution
        theme = scene_graph.theme if hasattr(scene_graph, 'theme') else None
        self._element_renderer = ElementRenderer(self._backend, theme)

        pres = self._backend.create_presentation()

        for slide_node in scene_graph.slide_order():
            try:
                self._render_slide(pres, slide_node, scene_graph, result)
                result.slides_rendered += 1
            except Exception as e:
                result.failed_slides.append(slide_node.id)
                result.errors.append(f"Slide {slide_node.id}: {e}")

        self._backend.save(pres, output_path)
        self._backend.close(pres)
        result.success = len(result.failed_slides) == 0
        return result

    def _render_slide(
        self,
        pres: Any,
        slide_node: SlideNode,
        sg: SceneGraph,
        result: RenderResult,
    ) -> None:
        slide = self._backend.add_slide(pres)

        if slide_node.background_color:
            resolved_bg = slide_node.background_color
            if not resolved_bg.startswith("#") and hasattr(sg, 'theme') and sg.theme:
                try:
                    resolved_bg = sg.theme.color(resolved_bg)
                except (AttributeError, KeyError):
                    pass
            self._backend.set_slide_background(slide, resolved_bg)

        for element_id in slide_node.all_element_ids():
            element = sg.get_element(element_id)
            if element is None:
                result.failed_elements.append(element_id)
                result.errors.append(f"Element not found: {element_id}")
                continue

            try:
                handle = self._element_renderer.render(slide, element)
                result.element_map[element_id] = handle
                result.elements_rendered += 1
            except Exception as e:
                result.failed_elements.append(element_id)
                result.errors.append(f"Element {element_id}: {e}")
