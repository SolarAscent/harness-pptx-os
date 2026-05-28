"""Renderer package — SceneGraph to PowerPoint rendering pipeline."""

from harness_pptx.renderer.engine import DeckRenderer, RenderResult
from harness_pptx.renderer.element_renderer import ElementRenderer

__all__ = ["DeckRenderer", "RenderResult", "ElementRenderer"]
