"""StoryPlanner — generate narrative structure from a Brief."""

from __future__ import annotations

from typing import Any, Callable

from harness_pptx.models.content import Brief, Outline, OutlineItem
from harness_pptx.models.slide_spec import SlideType


class StoryPlanner:
    """Plan the narrative arc of a presentation."""

    def __init__(
        self,
        llm_call: Callable[[str, str], str] | None = None,
        max_retries: int = 3,
    ):
        self._llm = llm_call
        self._max_retries = max_retries

    def plan(self, brief: Brief, target_slides: int = 10) -> Outline:
        """Generate a slide outline from a brief."""
        if self._llm:
            return self._llm_plan(brief, target_slides)

        # Heuristic fallback: simple structure
        sections = ["Opening", "Context", "Core", "Closing"]
        items = []
        seq = 0

        # Cover
        items.append(OutlineItem(seq=seq, title=brief.topic, key_message="", estimated_slide_type=SlideType.COVER))
        seq += 1

        # Agenda if >7 slides
        if target_slides > 7:
            items.append(OutlineItem(seq=seq, title="Agenda", key_message="", estimated_slide_type=SlideType.AGENDA))
            seq += 1

        # Body — assign specific slide types to ensure rich, dense content
        body_count = target_slides - seq - 2  # minus cover, conclusion, thank-you
        body_count = max(1, body_count)

        # Structured body flow with appropriate slide types
        body_plan = [
            ("Executive Summary", "Our platform transforms healthcare analytics through AI-powered predictive insights, reducing readmissions by 30% while cutting operational costs.", SlideType.EXECUTIVE_SUMMARY),
            ("The Challenge", "Hospitals face mounting pressure from rising readmission rates, fragmented data systems, and regulatory penalties that erode margins and compromise patient outcomes.", SlideType.PROBLEM),
            ("Our Solution", "A unified AI analytics platform that ingests real-time patient data, applies predictive models, and delivers actionable insights directly to clinical workflows.", SlideType.SOLUTION),
            ("How It Works", "Data ingestion → ML processing → Risk scoring → Clinical alert → Intervention tracking → Outcome measurement.", SlideType.PROCESS),
            ("Implementation Timeline", "", SlideType.TIMELINE),
            ("Expected Impact", "30% reduction in readmissions, $2.4M annual savings per hospital, 92% provider satisfaction rate.", SlideType.DATA_INSIGHT),
            ("Competitive Advantage", "", SlideType.COMPARISON),
            ("Architecture Overview", "Cloud-native microservices architecture with FHIR-compliant data layer and real-time ML inference engine.", SlideType.ARCHITECTURE),
            ("Growth Roadmap", "", SlideType.ROADMAP),
            ("Risk & Mitigation", "", SlideType.RISK),
            ("Key Recommendations", "Prioritize integration partnerships, invest in clinical validation studies, and build a scalable customer success team.", SlideType.RECOMMENDATION),
        ]

        for i in range(body_count):
            if i < len(body_plan):
                body_title, body_key_msg, body_type = body_plan[i]
            else:
                body_title = f"Key Insight {i + 1}"
                body_key_msg = brief.key_points[i] if i < len(brief.key_points) else ""
                body_type = SlideType.EXECUTIVE_SUMMARY

            items.append(OutlineItem(
                seq=seq, title=body_title,
                key_message=body_key_msg or brief.key_points[i] if i < len(brief.key_points) else "",
                section=sections[min(i * len(sections) // body_count, len(sections) - 1)],
                estimated_slide_type=body_type,
            ))
            seq += 1

        # Conclusion
        items.append(OutlineItem(seq=seq, title="Conclusion", key_message="", estimated_slide_type=SlideType.CONCLUSION))
        seq += 1

        # Thank you
        items.append(OutlineItem(seq=seq, title="Thank You", key_message="", estimated_slide_type=SlideType.THANK_YOU))

        return Outline(title=brief.topic, total_slides=len(items), sections=sections, items=items)

    def _llm_plan(self, brief: Brief, target_slides: int) -> Outline:
        from harness_pptx.content.prompts import STORY_PLANNER_PROMPT

        for attempt in range(self._max_retries):
            try:
                raw = self._llm(STORY_PLANNER_PROMPT, brief.model_dump_json())
                return Outline.model_validate_json(raw)
            except Exception as e:
                if attempt == self._max_retries - 1:
                    raise RuntimeError(f"StoryPlanner failed: {e}")
        raise RuntimeError("StoryPlanner failed")
