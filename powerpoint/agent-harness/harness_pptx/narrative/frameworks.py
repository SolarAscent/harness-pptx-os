"""Narrative frameworks — structured story patterns for presentations."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field

from harness_pptx.models.content import Outline, OutlineItem
from harness_pptx.models.slide_spec import SlideType


@dataclass
class NarrativePhase:
    name: str
    role: str
    slide_count: int = 1
    slide_types: list[SlideType] = field(default_factory=list)


class NarrativeFramework(ABC):
    """Base class for narrative frameworks."""

    name: str = "base"
    phases: list[NarrativePhase] = []

    def apply(self, outline: Outline) -> Outline:
        """Assign narrative roles to outline items based on the framework phases."""
        items = list(outline.items)
        if not items:
            return outline

        # First and last are special
        if items:
            items[0].estimated_slide_type = SlideType.COVER
            items[0].section = "Opening"
        if len(items) > 1:
            items[-1].estimated_slide_type = SlideType.CONCLUSION
            items[-1].section = "Closing"

        # Distribute remaining items across phases
        body_items = items[1:-1] if len(items) > 2 else []
        body_phases = [p for p in self.phases if p.role not in ("opening", "closing")]

        if body_phases and body_items:
            per_phase = max(1, len(body_items) // len(body_phases))
            for i, item in enumerate(body_items):
                phase_idx = min(i // per_phase, len(body_phases) - 1)
                phase = body_phases[phase_idx]
                item.section = phase.name
                if phase.slide_types:
                    item.estimated_slide_type = phase.slide_types[i % len(phase.slide_types)]

        return outline


# ---- Concrete frameworks ----------------------------------------------------

class ProblemSolutionFramework(NarrativeFramework):
    name = "problem-solution"
    phases = [
        NarrativePhase("Context", "context", 1, [SlideType.EXECUTIVE_SUMMARY]),
        NarrativePhase("Problem", "problem", 2, [SlideType.PROBLEM]),
        NarrativePhase("Solution", "solution", 2, [SlideType.SOLUTION, SlideType.PROCESS]),
        NarrativePhase("Evidence", "evidence", 2, [SlideType.DATA_INSIGHT, SlideType.CHART]),
        NarrativePhase("Call to Action", "cta", 1, [SlideType.RECOMMENDATION]),
    ]


class SCQAFramework(NarrativeFramework):
    """Situation-Complication-Question-Answer (McKinsey-style)."""

    name = "scqa"
    phases = [
        NarrativePhase("Situation", "situation", 1, [SlideType.EXECUTIVE_SUMMARY]),
        NarrativePhase("Complication", "complication", 1, [SlideType.PROBLEM]),
        NarrativePhase("Question", "question", 1, [SlideType.FRAMEWORK]),
        NarrativePhase("Answer", "answer", 3, [SlideType.SOLUTION, SlideType.DATA_INSIGHT, SlideType.RECOMMENDATION]),
    ]


class PyramidPrincipleFramework(NarrativeFramework):
    name = "pyramid"
    phases = [
        NarrativePhase("Key Message", "top", 1, [SlideType.EXECUTIVE_SUMMARY]),
        NarrativePhase("Supporting Arguments", "mid", 3, [SlideType.FRAMEWORK, SlideType.CHART, SlideType.CASE_STUDY]),
        NarrativePhase("Details", "base", 2, [SlideType.DATA_INSIGHT, SlideType.TABLE]),
    ]


class ThreeActFramework(NarrativeFramework):
    name = "three-act"
    phases = [
        NarrativePhase("Setup", "act1", 2, [SlideType.EXECUTIVE_SUMMARY, SlideType.PROBLEM]),
        NarrativePhase("Confrontation", "act2", 4, [SlideType.SOLUTION, SlideType.PROCESS, SlideType.DATA_INSIGHT, SlideType.CHART]),
        NarrativePhase("Resolution", "act3", 2, [SlideType.RECOMMENDATION, SlideType.CONCLUSION]),
    ]


class McKinseyFlowFramework(NarrativeFramework):
    name = "mckinsey"
    phases = [
        NarrativePhase("Executive Summary", "exec", 1, [SlideType.EXECUTIVE_SUMMARY]),
        NarrativePhase("Context & Problem", "context", 1, [SlideType.PROBLEM]),
        NarrativePhase("Framework", "framework", 1, [SlideType.FRAMEWORK]),
        NarrativePhase("Analysis", "analysis", 3, [SlideType.DATA_INSIGHT, SlideType.CHART, SlideType.TABLE]),
        NarrativePhase("Synthesis", "synthesis", 1, [SlideType.COMPARISON]),
        NarrativePhase("Recommendation", "rec", 1, [SlideType.RECOMMENDATION]),
        NarrativePhase("Next Steps", "next", 1, [SlideType.ROADMAP]),
    ]


# ---- Registry ---------------------------------------------------------------

_frameworks: dict[str, type[NarrativeFramework]] = {
    "problem-solution": ProblemSolutionFramework,
    "scqa": SCQAFramework,
    "pyramid": PyramidPrincipleFramework,
    "three-act": ThreeActFramework,
    "mckinsey": McKinseyFlowFramework,
}


def get_framework(name: str) -> NarrativeFramework:
    cls = _frameworks.get(name)
    if cls is None:
        raise KeyError(f"Framework '{name}' not found. Available: {list(_frameworks.keys())}")
    return cls()


def list_frameworks() -> list[str]:
    return sorted(_frameworks.keys())
