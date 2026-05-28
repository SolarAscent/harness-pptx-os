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

def _to_bullets(text: str, min_items: int = 3) -> list[str]:
    """Split text into bullet points, falling back to defaults if too short."""
    bullets = [s.strip() for s in text.replace("。", ".").replace("；", ";").replace("，", ",").split(".") if s.strip()]
    bullets = [b for b in bullets if len(b) > 3]
    if len(bullets) < min_items:
        return []
    return bullets[:6]


def _default_items(seed: str, count: int = 4) -> list[str]:
    """Generate placeholder items from a seed title."""
    templates = [
        f"Overview of {seed}",
        f"Key Drivers Behind {seed}",
        f"Current State & Challenges",
        f"Strategic Opportunities",
        f"Implementation Roadmap",
        f"Expected Outcomes & Impact",
    ]
    return templates[:count]


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

    content: dict[str, Any] = {
        "slide_id": intent.slide_id,
        "seq": intent.seq,
        "title": intent.title,
    }

    # ---- Per-type required-field population ---------------------------------
    # Each block receives content, extra, key_msg, bullets, brief and must
    # return a dict that satisfies its slide type's required_fields.

    if stype == "cover":
        # Truncate overly long titles (brief text used as title)
        title_text = intent.title
        if len(title_text) > 80:
            # Try to extract a reasonable title from the long text
            for sep in ["。", ".", "；", ";", "，", ","]:
                if sep in title_text[:80]:
                    title_text = title_text.split(sep)[0].strip()
                    break
            if len(title_text) > 80:
                title_text = title_text[:77] + "..."
        content["title"] = title_text
        content["subtitle"] = extra.get("subtitle", key_msg or f"A Presentation on {title_text}")
        content["author"] = extra.get("author", "")
        content["date"] = extra.get("date", "")

    elif stype == "agenda":
        items = extra.get("items", bullets or _default_items(intent.title, 5))
        content["items"] = items if items else _default_items(intent.title, 5)

    elif stype == "executive-summary":
        pts = extra.get("key_points", bullets or _default_items(intent.title, 4))
        content["key_points"] = pts if pts else _default_items(intent.title, 4)
        content["bottom_line"] = extra.get("bottom_line", key_msg if key_msg and not bullets else "")

    elif stype == "problem":
        content["problem_statement"] = extra.get("problem_statement") or key_msg or (
            "The current state presents significant challenges that demand a new approach. "
            "Without intervention, these issues will continue to compound."
        )
        content["pain_points"] = extra.get("pain_points") or bullets or [
            "Market inefficiency and fragmentation create hidden costs",
            "High operational overhead with limited transparency",
            "Lack of data-driven decision making leads to suboptimal outcomes",
        ]
        content["impact"] = extra.get("impact") or (
            "These challenges collectively result in measurable business impact "
            "across revenue, efficiency, and competitive position."
        )

    elif stype == "solution":
        content["solution_summary"] = extra.get("solution_summary") or key_msg or (
            "Our approach combines proven methodologies with innovative technology "
            "to deliver measurable results at scale."
        )
        content["key_features"] = extra.get("key_features") or bullets or [
            "Proprietary AI-powered analytics engine with real-time processing",
            "Intuitive dashboard with customizable alerts and reporting",
            "Seamless integration with existing enterprise systems",
            "Enterprise-grade security with full compliance coverage",
        ])
        content["benefits"] = extra.get("benefits") or []
        content["how_it_works"] = extra.get("how_it_works") or ""

    elif stype == "conclusion":
        takeaways = extra.get("key_takeaways", bullets if bullets else [
            "Clear market opportunity with proven demand",
            "Differentiated technology with defensible moat",
            "Strong unit economics & scalable GTM strategy",
            "Experienced team ready to execute",
        ])
        content["key_takeaways"] = takeaways[:6]
        content["call_to_action"] = extra.get("call_to_action", "Let's Build the Future Together")
        content["subtitle"] = extra.get("subtitle", "")

    elif stype == "thank-you":
        content["message"] = extra.get("message", intent.title if intent.title != "Thank You" else "Thank You")
        content["subtitle"] = extra.get("subtitle", "We look forward to your questions")
        content["contact"] = extra.get("contact", "")
        content["email"] = extra.get("email", "")
        content["website"] = extra.get("website", "")

    elif stype == "section-divider":
        content["section_number"] = extra.get("section_number", "")
        content["subtitle"] = extra.get("subtitle", key_msg)
        content["background_color"] = extra.get("background_color", "primary")

    elif stype == "timeline":
        items = extra.get("milestones", bullets)
        if not items:
            items = [
                {"date": "Q1 2025", "event": "Discovery & Research Phase"},
                {"date": "Q2 2025", "event": "MVP Development"},
                {"date": "Q3 2025", "event": "Beta Launch & Testing"},
                {"date": "Q4 2025", "event": "Market Rollout"},
                {"date": "Q1 2026", "event": "Scale & Optimize"},
            ]
        content["milestones"] = items

    elif stype == "process":
        items = extra.get("steps", bullets)
        if not items:
            items = [
                {"label": "1", "name": "Discovery", "description": "Understand needs & gather requirements"},
                {"label": "2", "name": "Design", "description": "Architect solution & create prototypes"},
                {"label": "3", "name": "Develop", "description": "Build, test & iterate rapidly"},
                {"label": "4", "name": "Deploy", "description": "Launch, monitor & continuously improve"},
            ]
        content["steps"] = items

    elif stype == "workflow":
        items = extra.get("steps", bullets)
        if not items:
            items = [
                {"name": "Input", "description": "Raw data ingestion"},
                {"name": "Process", "description": "AI-powered analysis"},
                {"name": "Output", "description": "Actionable insights"},
                {"name": "Feedback", "description": "Continuous learning loop"},
            ]
        content["steps"] = items

    elif stype == "comparison":
        content["left"] = extra.get("left", bullets[:3] if bullets else [
            "Legacy approach: manual, slow, error-prone",
            "High operational overhead & maintenance cost",
            "Limited scalability & integration capabilities",
        ])
        content["right"] = extra.get("right", [
            "Our approach: automated, fast, reliable",
            "Lower TCO with cloud-native architecture",
            "Enterprise-scale with seamless integrations",
        ])
        content["left_label"] = extra.get("left_label", "Traditional Approach")
        content["right_label"] = extra.get("right_label", "Our Solution")

    elif stype == "before-after":
        content["before_points"] = extra.get("before_points", bullets[:3] if bullets else [
            "Disconnected data silos across departments",
            "Reactive decision-making with 48-hour reporting lag",
            "15% average patient readmission rate",
        ])
        content["after_points"] = extra.get("after_points", [
            "Unified data platform with real-time dashboards",
            "Predictive analytics enabling proactive intervention",
            "Readmission rate reduced to under 10%",
        ])
        content["before_label"] = extra.get("before_label", "Before")
        content["after_label"] = extra.get("after_label", "After")

    elif stype == "roadmap":
        items = extra.get("phases", [])
        if not items:
            items = [
                {"name": "Phase 1", "timeline": "Q1-Q2", "description": "Foundation & MVP"},
                {"name": "Phase 2", "timeline": "Q3-Q4", "description": "Beta & Early Customers"},
                {"name": "Phase 3", "timeline": "Next Year", "description": "Scale & Market Expansion"},
            ]
        content["phases"] = items

    elif stype == "data-insight":
        content["insight"] = extra.get("insight", key_msg or f"Key data insight: {intent.title}")
        content["supporting_points"] = extra.get("supporting_points", bullets if bullets else [
            "Data-driven evidence supports this conclusion",
            "Statistical significance verified across cohorts",
            "Trend analysis confirms sustained improvement",
        ])
        content["chart_data"] = extra.get("chart_data", {})
        content["source"] = extra.get("source", "")

    elif stype == "chart":
        content["chart_data"] = extra.get("chart_data", {})
        content["chart_type"] = extra.get("chart_type", "bar")
        content["caption"] = extra.get("caption", key_msg)
        content["source"] = extra.get("source", "")

    elif stype == "table":
        content["headers"] = extra.get("headers", ["Category", "Value", "Change"])
        content["rows"] = extra.get("rows", [["—", "—", "—"]])
        content["caption"] = extra.get("caption", key_msg)

    elif stype == "framework":
        content["framework_name"] = extra.get("framework_name", intent.title)
        components = extra.get("components", bullets if bullets else [
            {"title": "Market Analysis", "description": "Size, growth, trends & competitive landscape"},
            {"title": "Value Proposition", "description": "Unique differentiation & customer value"},
            {"title": "Business Model", "description": "Revenue streams & unit economics"},
            {"title": "Growth Strategy", "description": "GTM plan & expansion roadmap"},
        ])
        content["components"] = components

    elif stype == "architecture":
        content["layers"] = extra.get("layers", [
            {"name": "Presentation", "description": "UI / Dashboard Layer"},
            {"name": "Application", "description": "Business Logic & APIs"},
            {"name": "Data", "description": "Storage, Processing & Analytics"},
            {"name": "Infrastructure", "description": "Cloud, Security & DevOps"},
        ])
        content["components"] = extra.get("components", [])
        content["description"] = extra.get("description", key_msg)

    elif stype == "recommendation":
        items = extra.get("recommendations", bullets if bullets else [
            "Invest in AI/ML capabilities to maintain competitive edge",
            "Expand into adjacent markets within 12 months",
            "Build strategic partnerships for distribution",
        ])
        content["recommendations"] = items
        content["rationale"] = extra.get("rationale", key_msg)
        content["next_steps"] = extra.get("next_steps", [])
        content["priority"] = extra.get("priority", "high")

    elif stype == "risk":
        items = extra.get("risks", [
            {"name": "Market Risk", "level": "Medium", "description": "Competitive pressure & market adoption speed"},
            {"name": "Technology Risk", "level": "Low", "description": "Proven architecture with fallback options"},
            {"name": "Execution Risk", "level": "Medium", "description": "Team scaling & operational complexity"},
        ])
        content["risks"] = items
        content["mitigation"] = extra.get("mitigation", "")

    elif stype == "team":
        content["members"] = extra.get("members", [
            {"name": "Leadership", "role": "Executive Team", "bio": "Industry veterans with 20+ years experience"},
            {"name": "Engineering", "role": "Tech Team", "bio": "Full-stack AI/ML engineering squad"},
            {"name": "Advisory", "role": "Board", "bio": "Domain experts & strategic advisors"},
        ])

    elif stype == "case-study":
        content["company"] = extra.get("company", "Example Corp")
        content["challenge"] = extra.get("challenge", key_msg or "The client faced significant challenges...")
        content["solution"] = extra.get("solution", "Our team deployed a customized solution...")
        content["results"] = extra.get("results", ["Result 1: 30% improvement", "Result 2: 50% cost reduction"])

    elif stype == "quote":
        content["quote_text"] = extra.get("quote_text", key_msg or intent.title)
        content["attribution"] = extra.get("attribution", "")
        content["role"] = extra.get("role", "")
        content["context"] = extra.get("context", "")

    elif stype == "appendix":
        content["content"] = extra.get("content", bullets if bullets else [intent.title])
        content["type"] = extra.get("type", "reference")
        content["reference"] = extra.get("reference", "")

    # Merge any remaining extra fields not explicitly handled
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
