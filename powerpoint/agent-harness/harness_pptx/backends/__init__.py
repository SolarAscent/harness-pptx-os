"""Backends package — PowerPoint rendering abstraction layer."""

from harness_pptx.backends.interface import (
    ChartStyle,
    ElementHandle,
    PresentationHandle,
    RendererInterface,
    ShapeStyle,
    SlideHandle,
    TableStyle,
    TextMetrics,
    TextStyle,
)
from harness_pptx.backends.registry import (
    BackendRegistry,
    get_backend,
    get_backend_registry,
)

__all__ = [
    "RendererInterface",
    "PresentationHandle",
    "SlideHandle",
    "ElementHandle",
    "TextStyle",
    "ShapeStyle",
    "TableStyle",
    "ChartStyle",
    "TextMetrics",
    "BackendRegistry",
    "get_backend",
    "get_backend_registry",
]
