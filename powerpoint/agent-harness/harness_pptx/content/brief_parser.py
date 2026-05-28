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

        # Heuristic fallback — actually parse structured text
        topic, key_points, audience, goal, tone, lang = self._heuristic_parse(text)
        return Brief(
            topic=topic,
            audience=audience,
            goal=goal,
            tone=tone,
            language=lang,
            key_points=key_points,
            source_text=text,
        )

    @staticmethod
    def _heuristic_parse(text: str) -> tuple:
        """Extract topic, key points, and metadata from structured text."""
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        topic = lines[0][:120] if lines else "Untitled"

        key_points = []
        audience = "general"
        goal = "inform"
        tone = Tone.PROFESSIONAL
        lang = Language.EN

        in_bullets = False
        for line in lines:
            ll = line.lower()

            # Detect language
            if any('一' <= c <= '鿿' for c in line):
                lang = Language.ZH

            # Detect metadata lines
            if "audience" in ll or "受众" in ll or "目标受众" in ll:
                audience = line.split("：", 1)[-1].split(":", 1)[-1].strip()[:80]
                continue
            if "goal" in ll or "目标" in ll:
                goal = "inform"
                continue
            if "tone" in ll or "风格" in ll or "语气" in ll:
                val = line.split("：", 1)[-1].split(":", 1)[-1].strip()[:80]
                if any(w in val for w in ["professional", "专业", "庄重", "大气"]):
                    tone = Tone.PROFESSIONAL
                elif any(w in val for w in ["inspire", "激励", "鼓舞"]):
                    tone = Tone.INSPIRING
                continue

            # Detect bullet points (explicit markers)
            stripped = line.lstrip("•·-*●○▪▸➤►✓✔☑✅🔸🔹")
            if stripped != line or line.startswith(("- ", "* ", "• ", "· ")):
                key_points.append(stripped.strip())
                in_bullets = True
                continue

            # Detect numbered items in structured sections
            if in_bullets and (line.startswith(("核心", "关键", "主要", "重要")) or stripped != line):
                key_points.append(stripped.strip())
                continue
            # Stop bullet mode on blank-line-like transitions
            if in_bullets and not stripped and not line.startswith(("-", "*", "•")):
                in_bullets = False

            # Detect dash-prefixed content lines (e.g. "- 北大历史：1898年建校...")
            if line.startswith("- ") and len(line) > 10:
                # Extract the label before ：or :
                content = line[2:].strip()
                key_points.append(content)
                in_bullets = True

        return topic, key_points, audience, goal, tone, lang

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
