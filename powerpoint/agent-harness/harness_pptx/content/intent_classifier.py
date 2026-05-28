"""SlideIntentClassifier — classify each outline item to a slide type."""

from __future__ import annotations

from typing import Callable

from harness_pptx.models.content import Outline, SlideIntent
from harness_pptx.models.slide_spec import SlideType


class SlideIntentClassifier:
    """Classify outline items into specific slide intents with types."""

    def __init__(self, llm_call: Callable[[str, str], str] | None = None):
        self._llm = llm_call

    def classify(self, outline: Outline) -> list[SlideIntent]:
        if self._llm:
            return self._llm_classify(outline)
        return self._heuristic_classify(outline)

    def _heuristic_classify(self, outline: Outline) -> list[SlideIntent]:
        intents: list[SlideIntent] = []
        for item in outline.items:
            stype = item.estimated_slide_type
            if stype == SlideType.CUSTOM:
                stype = self._guess_type(item.title, item.key_message, item.seq, outline.total_slides)
            intents.append(SlideIntent(
                slide_id=f"slide-{item.seq}",
                seq=item.seq,
                slide_type=stype,
                title=item.title,
                key_message=item.key_message,
                section=item.section,
            ))
        return intents

    @staticmethod
    def _guess_type(title: str, message: str, seq: int, total: int) -> SlideType:
        tl = title.lower()
        if seq == 0:
            return SlideType.COVER
        if seq == total - 1:
            return SlideType.THANK_YOU
        if seq == total - 2 and total > 3:
            return SlideType.CONCLUSION
        if "agenda" in tl or "outline" in tl or "contents" in tl:
            return SlideType.AGENDA
        if "problem" in tl or "challenge" in tl or "issue" in tl:
            return SlideType.PROBLEM
        if "solution" in tl or "approach" in tl or "resolve" in tl:
            return SlideType.SOLUTION
        if "timeline" in tl or "milestone" in tl or "roadmap" in tl or "schedule" in tl or "implementation" in tl:
            return SlideType.TIMELINE
        if "team" in tl or "people" in tl or "member" in tl:
            return SlideType.TEAM
        if "compare" in tl or "vs" in tl or "versus" in tl:
            return SlideType.COMPARISON
        if "conclusion" in tl or "summary" in tl or "takeaway" in tl:
            return SlideType.CONCLUSION
        if "process" in tl or "workflow" in tl or "pipeline" in tl or "flow" in tl:
            return SlideType.PROCESS
        if "architecture" in tl or "system" in tl or "design" in tl or "how it works" in tl:
            return SlideType.ARCHITECTURE
        if "result" in tl or "data" in tl or "metric" in tl or "insight" in tl or "impact" in tl:
            return SlideType.DATA_INSIGHT
        if "risk" in tl:
            return SlideType.RISK
        if "recommend" in tl or "action" in tl or "next step" in tl:
            return SlideType.RECOMMENDATION
        # Default for body slides: use executive-summary as a generic content slide
        return SlideType.EXECUTIVE_SUMMARY

    def _llm_classify(self, outline: Outline) -> list[SlideIntent]:
        from harness_pptx.content.prompts import INTENT_CLASSIFIER_PROMPT
        raw = self._llm(INTENT_CLASSIFIER_PROMPT, outline.model_dump_json())
        return [SlideIntent.model_validate(item) for item in __import__("json").loads(raw)]
