"""SpeakerNotesWriter — generate speaker notes based on slide content + role."""

from __future__ import annotations

from typing import Callable

from harness_pptx.models.content import SlideIntent


class SpeakerNotesWriter:
    """Write speaker notes for each slide."""

    def __init__(self, llm_call: Callable[[str, str], str] | None = None):
        self._llm = llm_call

    def write(self, intent: SlideIntent) -> str:
        """Generate speaker notes for a single slide intent."""
        if self._llm:
            return self._llm(
                "Write 2-4 sentences of speaker notes for this slide.",
                intent.model_dump_json(),
            )

        # Fallback
        role_hint = f"[{intent.narrative_role}] " if intent.narrative_role else ""
        parts = [f"{role_hint}{intent.key_message}"] if intent.key_message else []
        parts.extend(intent.bullet_points[:3])
        return " • ".join(parts)
