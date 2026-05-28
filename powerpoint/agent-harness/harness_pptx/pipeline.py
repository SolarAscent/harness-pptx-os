"""DeckPipeline — unified orchestration of the complete PPT generation pipeline.

This is the main entry point that coordinates all 13 pipeline steps:
  1. BriefParser → Brief
  2. StoryPlanner → NarrativeStructure
  3. OutlineBuilder → Outline
  4. IntentClassifier → list[SlideIntent]
  5. ContentCompressor → compressed slide content
  6. SlideTypeRegistry → SceneGraph
  7. ThemeRegistry → Theme applied
  8. LayoutEngine → bbox resolved
  9. DeckRenderer → .pptx
  10. PreviewExport → previews/
  11. QAEngine → QAReport
  12. RepairEngine → repair loop until pass
  13. FinalExport → final.pptx
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from harness_pptx.cli.workspace import ProjectWorkspace
from harness_pptx.models.content import Brief, Outline, SlideIntent
from harness_pptx.models.qa import QAReport
from harness_pptx.models.scene_graph import SceneGraph
from harness_pptx.models.theme import Theme
from harness_pptx.renderer.engine import RenderResult


@dataclass
class PipelineConfig:
    """Configuration for the DeckPipeline."""

    theme: str = "corporate"
    target_slides: int = 10
    max_repair_loops: int = 3
    output_dir: str = "."
    llm_call: Callable[[str, str], str] | None = None
    verbose: bool = False
    enable_visual_review: bool = False


@dataclass
class PipelineResult:
    """Result of a complete pipeline run."""

    success: bool = False
    output_path: str = ""
    brief: Brief | None = None
    outline: Outline | None = None
    intents: list[SlideIntent] = field(default_factory=list)
    scene_graph: SceneGraph | None = None
    theme: Theme | None = None
    render_result: RenderResult | None = None
    qa_report: QAReport | None = None
    repair_loops: int = 0
    errors: list[str] = field(default_factory=list)


# ---- Content builder: maps SlideIntent → slide-type build() content dict -------------

import re

def _is_chinese(text: str) -> bool:
    """Check if text contains Chinese characters."""
    return any('一' <= c <= '鿿' for c in text)


def _to_bullets(text: str, min_items: int = 2) -> list[str]:
    """Split text into bullet points, handling both Chinese and English."""
    if not text:
        return []
    # Normalize separators: Chinese/English commas → periods for splitting
    normalized = text.replace("。", ".").replace("；", ".").replace(";", ".").replace("，", ".").replace("、", ".")
    parts = [s.strip() for s in normalized.split(".") if s.strip()]
    parts = [p for p in parts if len(p) > 2]
    if len(parts) < min_items:
        return []
    return parts[:6]


def _default_items(seed: str, count: int = 4) -> list[str]:
    """Generate placeholder items from a seed title, in the appropriate language."""
    if _is_chinese(seed):
        templates = [
            f"{seed}概览",
            f"{seed}核心要点",
            f"现状与亮点",
            f"重要数据与指标",
            f"发展趋势与展望",
            f"总结与建议",
        ]
    else:
        templates = [
            f"Overview of {seed}",
            f"Key Highlights",
            f"Current State & Performance",
            f"Strategic Opportunities",
            f"Future Outlook",
            f"Summary & Recommendations",
        ]
    return templates[:count]


def _parse_timeline_items(key_msg: str, title: str) -> list[dict]:
    """Parse key_message into timeline milestone items."""
    parts = [p.strip() for p in re.split(r'[，,。；;、]', key_msg) if p.strip() and len(p.strip()) > 2]
    if not parts:
        return []
    items = []
    for part in parts[:6]:
        m = re.search(r'(\d{4})\s*年', part)
        if m:
            year = m.group(1)
            event = part.replace(m.group(0), "").strip().lstrip("，,")
            if not event:
                event = part
            items.append({"date": f"{year}", "event": event[:60]})
        else:
            items.append({"date": "", "event": part[:60]})
    return items if items else []


def _parse_team_members(key_msg: str) -> list[dict]:
    """Parse key_message into team member items."""
    parts = [p.strip() for p in re.split(r'[、，,]', key_msg) if p.strip() and len(p.strip()) > 1]
    if not parts:
        return []
    members = []
    for part in parts[:6]:
        part = part.rstrip("等")
        name = part
        bio = ""
        m = re.search(r'[（(]([^）)]+)[）)]', part)
        if m:
            name = part[:m.start()].strip()
            bio = m.group(1).strip()
        members.append({"name": name[:20], "role": bio[:40], "bio": bio[:60]})
    return members if members else []


def _parse_roadmap_phases(key_msg: str) -> list[dict]:
    """Parse key_message into roadmap phase items."""
    parts = [p.strip() for p in re.split(r'[，,。；;]', key_msg) if p.strip() and len(p.strip()) > 2]
    if not parts:
        return []
    phases = []
    for i, part in enumerate(parts[:5]):
        phase_num = i + 1
        phases.append({
            "name": f"阶段 {phase_num}",
            "timeline": "",
            "description": part[:80],
        })
    return phases if phases else []


def _populate_slide_content(intent, brief=None):
    """Build a complete content dict for any slide type from SlideIntent data.

    This is the central bridge between LLM-generated intents and the
    slide-type template build() methods.  It ensures every required field
    is populated — from intent.extra (when the LLM provides structured
    data), from intent.key_message (when available), or from sensible
    defaults (for heuristic / no-LLM mode).
    """
    stype = intent.slide_type.value
    key_msg = getattr(intent, "key_message", "") or ""
    bullets = _to_bullets(key_msg)
    extra = dict(getattr(intent, "extra", {}) or {})
    use_zh = _is_chinese(intent.title) or _is_chinese(key_msg)
    if not use_zh and brief:
        use_zh = _is_chinese(brief.topic) or any(_is_chinese(kp) for kp in brief.key_points[:1])

    content: dict[str, Any] = {
        "slide_id": intent.slide_id,
        "seq": intent.seq,
        "title": intent.title,
    }

    # ---- Per-type required-field population ---------------------------------

    if stype == "cover":
        title_text = intent.title
        if len(title_text) > 80:
            for sep in ["。", ".", "；", ";", "，", ","]:
                if sep in title_text[:80]:
                    title_text = title_text.split(sep)[0].strip()
                    break
            if len(title_text) > 80:
                title_text = title_text[:77] + "..."
        content["title"] = title_text
        # Use first key_point as subtitle if available, otherwise blank
        default_subtitle = ""
        if key_msg:
            default_subtitle = key_msg
        elif brief and brief.key_points:
            first_kp = brief.key_points[0]
            default_subtitle = first_kp.split("：")[-1].split(":")[-1].strip()[:80]
        content["subtitle"] = extra.get("subtitle", default_subtitle)
        content["author"] = extra.get("author", "")
        content["date"] = extra.get("date", "")

    elif stype == "agenda":
        items = extra.get("items", bullets or _default_items(intent.title, 5))
        if not items and brief and brief.key_points:
            items = [p.split("：")[0].split(":")[0].strip() for p in brief.key_points[:8]]
        content["items"] = items[:7] if items else _default_items(intent.title, 5)

    elif stype == "executive-summary":
        pts = extra.get("key_points", bullets if bullets else key_msg.split("，") if key_msg and use_zh else [])
        if not pts:
            pts = _default_items(intent.title, 4)
        content["key_points"] = [p.strip() for p in pts[:6] if p.strip()]
        content["bottom_line"] = extra.get("bottom_line", "")

    elif stype == "problem":
        content["problem_statement"] = extra.get("problem_statement") or key_msg or (
            "当前面临重大挑战，亟需新的解决方案。" if use_zh else
            "The current state presents significant challenges that demand a new approach."
        )
        content["pain_points"] = extra.get("pain_points") or bullets or (
            [f"{intent.title}相关核心痛点", "现有体系难以满足发展需求", "变革与创新势在必行"]
            if use_zh else
            ["Market inefficiency and fragmentation create hidden costs",
             "High operational overhead with limited transparency",
             "Lack of data-driven decision making"]
        )
        content["impact"] = extra.get("impact") or (
            "这些挑战对组织的可持续发展构成实质影响。" if use_zh else
            "These challenges result in measurable business impact across key metrics."
        )

    elif stype == "solution":
        content["solution_summary"] = extra.get("solution_summary") or key_msg or (
            "我们提出系统性解决方案。" if use_zh else
            "Our approach combines proven methodologies with innovative technology."
        )
        content["key_features"] = extra.get("key_features") or bullets or (
            ["核心技术能力卓越", "系统架构先进可靠", "用户体验全面优化", "安全保障完善到位"]
            if use_zh else
            ["Core technology capabilities", "Advanced system architecture", "Optimized user experience", "Comprehensive security"]
        )
        content["benefits"] = extra.get("benefits") or []
        content["how_it_works"] = extra.get("how_it_works") or ""

    elif stype == "conclusion":
        if extra.get("key_takeaways"):
            takeaways = extra["key_takeaways"]
        elif bullets:
            takeaways = bullets
        elif brief and brief.key_points:
            takeaways = [p.split("：")[0].split(":")[0].strip() for p in brief.key_points[:6]]
        else:
            takeaways = (
                ["核心观点与关键结论", "数据支撑的决策建议", "下一步行动计划", "持续优化与跟进"]
                if use_zh else
                ["Key strategic insights and findings", "Data-backed recommendations", "Immediate next steps", "Long-term roadmap"]
            )
        content["key_takeaways"] = takeaways[:6]
        content["call_to_action"] = extra.get("call_to_action", "携手共创未来" if use_zh else "Let's Build the Future Together")
        content["subtitle"] = extra.get("subtitle", "")

    elif stype == "thank-you":
        content["message"] = extra.get("message", intent.title)
        content["subtitle"] = extra.get("subtitle",
            "感谢您的关注与支持" if use_zh else "We look forward to your questions")
        content["contact"] = extra.get("contact", "")
        content["email"] = extra.get("email", "")
        content["website"] = extra.get("website", "")

    elif stype == "section-divider":
        content["section_number"] = extra.get("section_number", "")
        content["subtitle"] = extra.get("subtitle", key_msg)
        content["background_color"] = extra.get("background_color", "primary")

    elif stype == "timeline":
        items = extra.get("milestones", None)
        if not items:
            items = _parse_timeline_items(key_msg, intent.title)
        if not items:
            items = [{"date": f"阶段{i+1}", "event": f"{intent.title}关键节点{i+1}"} for i in range(5)]
        content["milestones"] = items[:6]

    elif stype == "process":
        items = extra.get("steps", None)
        if not items and bullets:
            items = [{"label": str(i+1), "name": b[:20], "description": b[:60]} for i, b in enumerate(bullets[:6])]
        if not items:
            labels = "一二三四五六"
            items = [
                {"label": labels[i] if i < len(labels) else str(i+1),
                 "name": f"{intent.title}步骤{i+1}",
                 "description": f"第{i+1}步关键环节"}
                for i in range(4)
            ]
        content["steps"] = items[:6]

    elif stype == "workflow":
        items = extra.get("steps", None)
        if not items and bullets:
            items = [{"name": b[:20], "description": b[:60]} for b in bullets[:5]]
        if not items:
            items = [
                {"name": "输入", "description": "数据与需求采集"},
                {"name": "处理", "description": "分析与加工"},
                {"name": "输出", "description": "成果交付"},
                {"name": "反馈", "description": "持续优化迭代"},
            ]
        content["steps"] = items[:5]

    elif stype == "comparison":
        content["left"] = extra.get("left", bullets[:3] if bullets else (
            ["传统模式：效率低、成本高", "信息孤岛、协同困难", "扩展性不足、响应滞后"] if use_zh else
            ["Legacy approach: manual, slow, error-prone", "High operational overhead", "Limited scalability"]
        ))
        content["right"] = extra.get("right", (
            ["新模式：高效、精准、智能", "数据互通、全面协同", "弹性扩展、快速响应"] if use_zh else
            ["Modern approach: automated, fast, reliable", "Lower TCO with cloud-native architecture", "Enterprise-scale with seamless integrations"]
        ))
        content["left_label"] = extra.get("left_label", "传统方式" if use_zh else "Traditional")
        content["right_label"] = extra.get("right_label", "创新方案" if use_zh else "Our Solution")

    elif stype == "before-after":
        content["before_points"] = extra.get("before_points", bullets[:3] if bullets else (
            ["之前面临的问题与挑战", "效率与质量有待提升", "资源利用不够充分"] if use_zh else
            ["Disconnected data silos", "Reactive decision-making", "High costs and inefficiency"]
        ))
        content["after_points"] = extra.get("after_points", (
            ["显著改善与全面提升", "效率大幅提高", "资源优化配置完成"] if use_zh else
            ["Unified platform with real-time dashboards", "Predictive analytics", "Cost reduction achieved"]
        ))
        content["before_label"] = extra.get("before_label", "改善前" if use_zh else "Before")
        content["after_label"] = extra.get("after_label", "改善后" if use_zh else "After")

    elif stype == "roadmap":
        items = extra.get("phases", None)
        if not items:
            items = _parse_roadmap_phases(key_msg)
        if not items:
            items = [
                {"name": f"第{i+1}阶段", "timeline": "", "description": f"{intent.title}规划第{i+1}步"}
                for i in range(3)
            ]
        content["phases"] = items[:5]

    elif stype == "data-insight":
        content["insight"] = extra.get("insight") or key_msg or intent.title
        content["supporting_points"] = extra.get("supporting_points") or bullets or (
            [p.strip() for p in key_msg.split("，") if p.strip() and len(p.strip()) > 3][:4]
            if key_msg and use_zh else
            ["Key metrics and indicators", "Trend analysis and patterns", "Actionable insights derived"]
        )
        content["chart_data"] = extra.get("chart_data", {})
        content["source"] = extra.get("source", "")

    elif stype == "chart":
        content["chart_data"] = extra.get("chart_data", {})
        content["chart_type"] = extra.get("chart_type", "bar")
        content["caption"] = extra.get("caption", key_msg)
        content["source"] = extra.get("source", "")

    elif stype == "table":
        content["headers"] = extra.get("headers", ["类别" if use_zh else "Category", "数据" if use_zh else "Value", "说明" if use_zh else "Notes"])
        content["rows"] = extra.get("rows", [["—", "—", "—"]])
        content["caption"] = extra.get("caption", key_msg)

    elif stype == "framework":
        content["framework_name"] = extra.get("framework_name", intent.title)
        components = extra.get("components", bullets if bullets else [
            {"title": f"{intent.title}维度{i+1}", "description": f"第{i+1}个核心要素"}
            for i in range(4)
        ])
        content["components"] = components[:5]

    elif stype == "architecture":
        content["layers"] = extra.get("layers", [
            {"name": "表现层", "description": "用户界面与交互"},
            {"name": "应用层", "description": "业务逻辑与API"},
            {"name": "数据层", "description": "存储、处理与分析"},
            {"name": "基础设施层", "description": "云服务、安全与运维"},
        ])
        content["components"] = extra.get("components", [])
        content["description"] = extra.get("description", key_msg)

    elif stype == "recommendation":
        items = extra.get("recommendations", bullets if bullets else (
            ["加强核心能力建设", "拓展战略合作伙伴关系", "持续优化组织效能"] if use_zh else
            ["Invest in core capabilities", "Expand strategic partnerships", "Optimize organizational effectiveness"]
        ))
        content["recommendations"] = items[:6]
        content["rationale"] = extra.get("rationale", key_msg)
        content["next_steps"] = extra.get("next_steps", [])
        content["priority"] = extra.get("priority", "high")

    elif stype == "risk":
        items = extra.get("risks", [
            {"name": "外部风险", "level": "中", "description": "外部环境变化带来的不确定性"},
            {"name": "技术风险", "level": "低", "description": "技术方案成熟可靠，风险可控"},
            {"name": "执行风险", "level": "中", "description": "组织协同与资源配置挑战"},
        ])
        content["risks"] = items
        content["mitigation"] = extra.get("mitigation", "")

    elif stype == "team":
        items = extra.get("members", None)
        if not items:
            items = _parse_team_members(key_msg)
        if not items:
            items = [{"name": f"{intent.title}成员{i+1}", "role": "核心成员", "bio": "专业背景深厚，经验丰富"} for i in range(3)]
        content["members"] = items[:6]

    elif stype == "case-study":
        content["company"] = extra.get("company", intent.title)
        content["challenge"] = extra.get("challenge", key_msg or "面临的核心挑战")
        content["solution"] = extra.get("solution", "系统性解决方案")
        content["results"] = extra.get("results", ["显著成效", "持续改善", "广泛认可"])

    elif stype == "quote":
        content["quote_text"] = extra.get("quote_text", key_msg or intent.title)
        content["attribution"] = extra.get("attribution", "")
        content["role"] = extra.get("role", "")
        content["context"] = extra.get("context", "")

    elif stype == "appendix":
        content["content"] = extra.get("content", bullets if bullets else [intent.title])
        content["type"] = extra.get("type", "reference")
        content["reference"] = extra.get("reference", "")

    for k, v in extra.items():
        if k not in content:
            content[k] = v

    return content


class DeckPipeline:
    """Unified orchestrator for text-to-PPTX generation.

    Usage::

        pipeline = DeckPipeline(config)
        result = pipeline.run("Our product is...")
        print(f"Output: {result.output_path}")
    """

    def __init__(self, config: PipelineConfig | None = None):
        self.config = config or PipelineConfig()

    def run(self, brief_text: str) -> PipelineResult:
        result = PipelineResult()
        cfg = self.config

        try:
            # Step 1: Parse brief
            from harness_pptx.content.brief_parser import BriefParser
            parser = BriefParser(llm_call=cfg.llm_call)
            brief = parser.parse(brief_text)
            result.brief = brief

            # Step 2: Plan narrative
            from harness_pptx.content.story_planner import StoryPlanner
            planner = StoryPlanner(llm_call=cfg.llm_call)
            outline = planner.plan(brief, target_slides=cfg.target_slides)
            result.outline = outline

            # Step 3: Build outline
            from harness_pptx.content.outline_builder import OutlineBuilder
            ob = OutlineBuilder()
            outline = ob.build(outline)

            # Step 4: Classify intents
            from harness_pptx.content.intent_classifier import SlideIntentClassifier
            classifier = SlideIntentClassifier(llm_call=cfg.llm_call)
            intents = classifier.classify(outline)
            result.intents = intents

            # Step 5-7: Theme + Slide types → SceneGraph
            from harness_pptx.themes.base import ThemeRegistry
            themes = ThemeRegistry()
            theme = themes.get(cfg.theme)
            result.theme = theme

            from harness_pptx.slide_types.base import get_slide_type_registry
            reg = get_slide_type_registry()
            sg = SceneGraph(deck_id=brief.topic, theme=theme)

            for i, intent in enumerate(intents):
                try:
                    st = reg.get(intent.slide_type.value)
                    content = _populate_slide_content(intent, brief)
                    node = st.build(content, theme, registry=sg.element_registry)
                    sg.slides[node.id] = node
                    sg.slide_count = len(sg.slides)
                except Exception as e:
                    result.errors.append(f"Slide {intent.slide_id}: {e}")

            result.scene_graph = sg

            # Step 8: Render
            from harness_pptx.backends.registry import get_backend
            from harness_pptx.renderer.engine import DeckRenderer

            output_path = str(Path(cfg.output_dir) / f"{brief.topic or 'deck'}.pptx")
            backend = get_backend()
            renderer = DeckRenderer(backend)
            rr = renderer.render(sg, output_path)
            result.render_result = rr
            result.output_path = output_path

            # Step 9-11: QA + Repair loop
            from harness_pptx.qa.engine import QAEngine
            from harness_pptx.repair.engine import RepairEngine
            from harness_pptx.repair.planner import RepairPlanner

            qa = QAEngine(theme=theme)
            repair_engine = RepairEngine()
            repair_planner = RepairPlanner()

            for loop in range(cfg.max_repair_loops):
                qa_report = qa.run(sg)
                result.qa_report = qa_report
                if qa_report.passed:
                    break
                plan = repair_planner.plan(qa_report)
                sg = repair_engine.repair(sg, plan)
                rr = renderer.render(sg, output_path)
                result.repair_loops += 1

            # Step 12: Visual review — AI-powered slide image inspection
            if cfg.enable_visual_review:
                try:
                    visual_report = qa.run_full(sg, pptx_path=output_path)
                    result.qa_report = visual_report
                except Exception as e:
                    result.errors.append(f"Visual review: {e}")

            result.success = True

        except Exception as e:
            result.errors.append(str(e))

        return result
