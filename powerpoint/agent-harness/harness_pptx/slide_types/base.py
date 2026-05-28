"""Slide type base protocol and registry."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

from harness_pptx.models.element import BaseElement
from harness_pptx.models.scene_graph import SlideNode
from harness_pptx.models.theme import Theme


def register_element(registry: dict[str, BaseElement] | None, element: BaseElement) -> None:
    """Register an element in the scene-graph registry if provided."""
    if registry is not None:
        registry[element.id] = element


class SlideType(ABC):
    """A slide type template.

    Each concrete slide type accepts structured content, a theme, and an
    optional element registry; it returns a fully-formed SlideNode. Elements
    are registered into ``registry`` so the SceneGraph can find them later.
    """

    name: str
    required_fields: list[str] = []
    optional_fields: list[str] = []
    visual_rules: dict[str, Any] = {}
    qa_rules: list[dict[str, Any]] = []

    @abstractmethod
    def build(
        self,
        content: dict[str, Any],
        theme: Theme,
        registry: dict[str, BaseElement] | None = None,
    ) -> SlideNode: ...


class SlideTypeRegistry:
    """Registry of all available slide types."""

    def __init__(self):
        self._types: dict[str, type[SlideType]] = {}

    def register(self, slide_type_cls: type[SlideType]) -> None:
        inst = slide_type_cls()
        self._types[inst.name] = slide_type_cls

    def get(self, name: str) -> SlideType:
        cls = self._types.get(name)
        if cls is None:
            available = sorted(self._types.keys())
            raise KeyError(f"Slide type '{name}' not found. Available: {available}")
        return cls()

    def list(self) -> list[str]:
        return sorted(self._types.keys())

    def has(self, name: str) -> bool:
        return name in self._types


# Singleton
_registry: SlideTypeRegistry | None = None


def get_slide_type_registry() -> SlideTypeRegistry:
    global _registry
    if _registry is None:
        _registry = SlideTypeRegistry()
        _register_all(_registry)
    return _registry


def _register_all(reg: SlideTypeRegistry) -> None:
    from harness_pptx.slide_types.cover import CoverSlide
    from harness_pptx.slide_types.agenda import AgendaSlide
    from harness_pptx.slide_types.section_divider import SectionDividerSlide
    from harness_pptx.slide_types.executive_summary import ExecutiveSummarySlide
    from harness_pptx.slide_types.problem import ProblemSlide
    from harness_pptx.slide_types.solution import SolutionSlide
    from harness_pptx.slide_types.timeline import TimelineSlide
    from harness_pptx.slide_types.process import ProcessSlide
    from harness_pptx.slide_types.framework import FrameworkSlide
    from harness_pptx.slide_types.comparison import ComparisonSlide
    from harness_pptx.slide_types.before_after import BeforeAfterSlide
    from harness_pptx.slide_types.data_insight import DataInsightSlide
    from harness_pptx.slide_types.chart_slide import ChartSlideType
    from harness_pptx.slide_types.table_slide import TableSlideType
    from harness_pptx.slide_types.case_study import CaseStudySlide
    from harness_pptx.slide_types.quote import QuoteSlide
    from harness_pptx.slide_types.team import TeamSlide
    from harness_pptx.slide_types.roadmap import RoadmapSlide
    from harness_pptx.slide_types.architecture import ArchitectureSlide
    from harness_pptx.slide_types.workflow import WorkflowSlide
    from harness_pptx.slide_types.risk import RiskSlide
    from harness_pptx.slide_types.recommendation import RecommendationSlide
    from harness_pptx.slide_types.conclusion import ConclusionSlide
    from harness_pptx.slide_types.thank_you import ThankYouSlide
    from harness_pptx.slide_types.appendix import AppendixSlide

    for cls in [
        CoverSlide, AgendaSlide, SectionDividerSlide, ExecutiveSummarySlide,
        ProblemSlide, SolutionSlide, TimelineSlide, ProcessSlide,
        FrameworkSlide, ComparisonSlide, BeforeAfterSlide, DataInsightSlide,
        ChartSlideType, TableSlideType, CaseStudySlide, QuoteSlide,
        TeamSlide, RoadmapSlide, ArchitectureSlide, WorkflowSlide,
        RiskSlide, RecommendationSlide, ConclusionSlide, ThankYouSlide,
        AppendixSlide,
    ]:
        reg.register(cls)
