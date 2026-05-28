"""ContentCompressor — compress long text to slide-ready copy."""

from __future__ import annotations

from typing import Callable


class ContentCompressor:
    """Compress long-form text into slide-appropriate copy (≤6 bullets)."""

    def __init__(self, llm_call: Callable[[str, str], str] | None = None):
        self._llm = llm_call

    def compress(self, text: str, max_bullets: int = 6, max_chars: int = 500) -> str:
        if self._llm:
            return self._llm_compress(text, max_bullets, max_chars)

        # Simple truncation fallback
        if len(text) <= max_chars:
            return text
        sentences = text.replace("。", ".").replace("！", ".").replace("？", ".").split(".")
        compressed = ". ".join(s.strip() for s in sentences[:max_bullets] if s.strip())
        if len(compressed) > max_chars:
            compressed = compressed[:max_chars - 3] + "..."
        return compressed

    def _llm_compress(self, text: str, max_bullets: int, max_chars: int) -> str:
        from harness_pptx.content.prompts import COMPRESSOR_PROMPT
        prompt = COMPRESSOR_PROMPT.format(max_bullets=max_bullets, max_chars=max_chars)
        return self._llm(prompt, text)
