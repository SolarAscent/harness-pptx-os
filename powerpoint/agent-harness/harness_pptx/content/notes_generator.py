"""SpeakerNotesGenerator — generate speaker notes for slides."""

from __future__ import annotations

from typing import Callable

from harness_pptx.models.content import SlideIntent


class SpeakerNotesGenerator:
    """Generate speaker notes from slide intents."""

    def __init__(self, llm_call: Callable[[str, str], str] | None = None):
        self._llm = llm_call

    def generate(self, intent: SlideIntent) -> str:
        if self._llm:
            return self._llm_generate(intent)

        # Simple fallback: restate title + bullets
        parts = [intent.key_message] if intent.key_message else []
        parts.extend(intent.bullet_points[:3])
        return " • ".join(p for p in parts if p)

    def _llm_generate(self, intent: SlideIntent) -> str:
        from harness_pptx.content.prompts import NOTES_GENERATOR_PROMPT
        return self._llm(NOTES_GENERATOR_PROMPT, intent.model_dump_json())
