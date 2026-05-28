"""StoryPlanner — generate narrative structure from a Brief."""

from __future__ import annotations

import re
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

    @staticmethod
    def _is_chinese(text: str) -> bool:
        return any('一' <= c <= '鿿' for c in text)

    @classmethod
    def _lang_labels(cls, brief: Brief) -> dict[str, str]:
        """Return localized labels based on brief language."""
        zh = cls._is_chinese(brief.topic) or any(cls._is_chinese(kp) for kp in brief.key_points[:1])
        if zh:
            return {"agenda": "内容概览", "conclusion": "总结", "thank_you": "感谢聆听"}
        return {"agenda": "Agenda", "conclusion": "Conclusion", "thank_you": "Thank You"}

    def plan(self, brief: Brief, target_slides: int = 10) -> Outline:
        """Generate a slide outline from a brief."""
        if self._llm:
            return self._llm_plan(brief, target_slides)

        labels = self._lang_labels(brief)
        items = []
        seq = 0

        # Cover
        items.append(OutlineItem(seq=seq, title=brief.topic, key_message="", estimated_slide_type=SlideType.COVER))
        seq += 1

        # Agenda if deck is large enough
        if target_slides > 7:
            items.append(OutlineItem(seq=seq, title=labels["agenda"], key_message="", estimated_slide_type=SlideType.AGENDA))
            seq += 1

        # Body — use brief.key_points when available, fall back to defaults
        reserved = 2  # conclusion + thank-you
        kp = brief.key_points

        if kp:
            for point in kp:
                if seq >= target_slides - reserved:
                    break
                title, key_msg = self._split_key_point(point)
                stype = self._infer_slide_type(title, key_msg)
                items.append(OutlineItem(
                    seq=seq, title=title, key_message=key_msg,
                    section="Body", estimated_slide_type=stype,
                ))
                seq += 1
        else:
            body_count = max(1, target_slides - seq - reserved)
            for i in range(body_count):
                items.append(OutlineItem(
                    seq=seq,
                    title=f"Key Point {i + 1}",
                    key_message="",
                    section="Body",
                    estimated_slide_type=SlideType.EXECUTIVE_SUMMARY,
                ))
                seq += 1

        # Conclusion
        items.append(OutlineItem(seq=seq, title=labels["conclusion"], key_message="", estimated_slide_type=SlideType.CONCLUSION))
        seq += 1

        # Thank you
        items.append(OutlineItem(seq=seq, title=labels["thank_you"], key_message="", estimated_slide_type=SlideType.THANK_YOU))

        return Outline(title=brief.topic, total_slides=len(items), sections=["Opening", "Body", "Closing"], items=items)

    @staticmethod
    def _split_key_point(point: str) -> tuple[str, str]:
        """Split '标题：内容' or 'Title: Content' into (title, key_message)."""
        for sep in ("：", ":", "——", "—"):
            if sep in point:
                parts = point.split(sep, 1)
                title = parts[0].strip()[:80]
                msg = parts[1].strip() if len(parts) > 1 else ""
                return title, msg
        return point[:80], ""

    @staticmethod
    def _infer_slide_type(title: str, key_msg: str) -> SlideType:
        """Infer the best slide type from Chinese and English keywords."""
        tl = title.lower()
        combined = f"{tl} {key_msg.lower()}"

        # History / timeline
        if any(w in combined for w in ["history", "历史", "timeline", "milestone", "发展历程", "沿革", "起源", "成立"]):
            return SlideType.TIMELINE

        # Data / stats / achievements
        if any(w in combined for w in ["data", "metric", "insight", "数据", "实力", "排名", "科研", "成就",
                                         "经费", "实验室", "专利", "论文", "成果", "指标", "统计"]):
            return SlideType.DATA_INSIGHT

        # Team / people
        if any(w in combined for w in ["team", "people", "member", "founder", "团队", "校友", "人物", "成员",
                                         "创始人", "领导", "师资", "人才"]):
            return SlideType.TEAM

        # Future / roadmap
        if any(w in combined for w in ["roadmap", "future", "vision", "plan", "phase", "未来", "规划", "展望",
                                         "路线图", "前景", "战略", "发展", "建设"]):
            return SlideType.ROADMAP

        # Problem / challenge
        if any(w in combined for w in ["problem", "challenge", "issue", "pain", "问题", "挑战", "困难", "痛点"]):
            return SlideType.PROBLEM

        # Solution
        if any(w in combined for w in ["solution", "approach", "resolve", "解决", "方案", "对策", "方法"]):
            return SlideType.SOLUTION

        # Process / workflow
        if any(w in combined for w in ["process", "workflow", "pipeline", "flow", "流程", "步骤", "环节"]):
            return SlideType.PROCESS

        # Architecture / system
        if any(w in combined for w in ["architecture", "system", "design", "架构", "系统", "结构", "体系"]):
            return SlideType.ARCHITECTURE

        # Comparison
        if any(w in combined for w in ["compare", "vs", "versus", "对比", "比较", "优势", "竞争"]):
            return SlideType.COMPARISON

        # Risk
        if any(w in combined for w in ["risk", "风险"]):
            return SlideType.RISK

        # Recommendation
        if any(w in combined for w in ["recommend", "action", "next step", "建议", "举措", "行动"]):
            return SlideType.RECOMMENDATION

        # Campus / facility / scenery → general content
        if any(w in combined for w in ["campus", "校园", "风光", "设施", "环境", "文化", "传统", "国际",
                                         "交流", "合作", "项目"]):
            return SlideType.EXECUTIVE_SUMMARY

        # Default body slide type
        return SlideType.EXECUTIVE_SUMMARY

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
