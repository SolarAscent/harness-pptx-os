"""Vision API client — doubao-seed-2-0-pro via Volcengine ARK.

Use only for PPT visual review. Do NOT use for general-purpose image analysis.
"""

from __future__ import annotations

import os
from typing import Any


class VisionClient:
    """Client for PPT slide visual review.

    Usage::

        client = VisionClient()
        result = client.review_slide("/path/to/slide.png", "检查文字是否溢出")
    """

    base_url = "https://ark.cn-beijing.volces.com/api/v3"
    model = "doubao-seed-2-0-pro-260215"

    def __init__(self, api_key: str | None = None):
        self._api_key = api_key or os.getenv("ARK_API_KEY", "")
        self._client = None

    @property
    def _openai_client(self):
        if self._client is None:
            from openai import OpenAI
            self._client = OpenAI(
                base_url=self.base_url,
                api_key=self._api_key,
            )
        return self._client

    def check_configured(self) -> bool:
        """Return True if the API key is configured."""
        return bool(self._api_key)

    def ask(
        self,
        prompt: str,
        image_url: str | None = None,
        image_base64: str | None = None,
    ) -> str:
        """Send a prompt with optional image to the vision model."""
        content: list[dict[str, Any]] = []

        if image_url:
            content.append({"type": "input_image", "image_url": image_url})
        elif image_base64:
            content.append({"type": "input_image", "image_url": image_base64})

        content.append({"type": "input_text", "text": prompt})

        response = self._openai_client.responses.create(
            model=self.model,
            input=[{"role": "user", "content": content}],
        )
        return response.output_text

    def review_slide(self, image_path: str, aspect: str | None = None) -> str:
        """Review a slide image for visual quality issues.

        Args:
            image_path: Path to the slide PNG export.
            aspect: Specific aspect to check (text overflow, alignment, etc.).
                    If None, runs a general visual review.

        Returns:
            Review feedback as a string.
        """
        import base64

        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        data_uri = f"data:image/png;base64,{b64}"

        if aspect:
            prompt = _REVIEW_PROMPTS.get(aspect, _REVIEW_PROMPTS["general"]).format(aspect=aspect)
        else:
            prompt = _REVIEW_PROMPTS["general"].format(aspect="整体视觉效果")

        return self.ask(prompt, image_base64=data_uri)


# ---- Prompt templates ----------------------------------------------------------

_REVIEW_PROMPTS = {
    "general": """你是一位专业的 PPT 设计审阅专家。请仔细查看这张幻灯片截图，从以下维度进行评审：

1. **文字溢出**：是否有文字超出文本框边界？
2. **对齐问题**：元素之间是否有明显的对齐错误？
3. **字体大小**：正文字号是否过小（小于10pt）？
4. **对比度**：文字与背景的对比度是否足够？
5. **间距**：元素之间的留白是否合理？
6. **整体美观**：布局是否协调、专业？

请用中文简洁指出发现的问题，每个问题一行。如果没有明显问题，请回复"无明显问题"。

检查方面：{aspect}""",

    "text_overflow": "请检查这张幻灯片中是否有文字溢出文本框或超出边界的情况。用中文简洁回答。",

    "alignment": "请检查这张幻灯片中各元素的水平/垂直对齐是否一致。用中文简洁回答。",

    "font_size": "请检查这张幻灯片中的字号是否过小（小于10pt视为过小）。用中文简洁回答。",

    "contrast": "请检查这张幻灯片中文字颜色与背景的对比度是否足够。用中文简洁回答。",

    "overlap": "请检查这张幻灯片中是否有元素互相重叠覆盖。用中文简洁回答。",

    "slide_density": "请检查这张幻灯片的信息密度是否过高或过低。用中文简洁回答。",
}


# ---- Singleton ----------------------------------------------------------------

_vision_client: VisionClient | None = None


def get_vision_client() -> VisionClient:
    global _vision_client
    if _vision_client is None:
        _vision_client = VisionClient()
    return _vision_client
