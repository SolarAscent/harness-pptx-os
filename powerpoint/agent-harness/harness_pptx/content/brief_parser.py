"""BriefParser — extract structured brief from raw user text."""

from __future__ import annotations

from typing import Any, Callable

from harness_pptx.models.content import Brief
from harness_pptx.models.deck_spec import Language, Tone


class BriefParser:
    """Parse raw text into a structured Brief.

    Uses an injected LLM call function for extraction, with Pydantic
    validation and automatic retry on schema mismatch.
    """

    def __init__(
        self,
        llm_call: Callable[[str, str], str] | None = None,
        max_retries: int = 3,
    ):
        self._llm = llm_call
        self._max_retries = max_retries

    def parse(self, text: str) -> Brief:
        """Extract brief from raw text.

        If an LLM callable is provided, uses it. Otherwise falls back to
        heuristic extraction (keyword matching + defaults).
        """
        if self._llm:
            return self._llm_parse(text)

        # Heuristic fallback
        return Brief(
            topic=self._guess_topic(text),
            audience="general",
            goal="inform",
            tone=Tone.PROFESSIONAL,
            language=Language.EN,
            source_text=text,
        )

    def _llm_parse(self, text: str) -> Brief:
        from harness_pptx.content.prompts import BRIEF_PARSER_PROMPT

        for attempt in range(self._max_retries):
            try:
                raw = self._llm(BRIEF_PARSER_PROMPT, text)
                return Brief.model_validate_json(raw)
            except Exception as e:
                if attempt == self._max_retries - 1:
                    raise RuntimeError(f"BriefParser failed after {self._max_retries} attempts: {e}")
        raise RuntimeError("BriefParser failed")

    @staticmethod
    def _guess_topic(text: str) -> str:
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        return lines[0][:120] if lines else "Untitled"
