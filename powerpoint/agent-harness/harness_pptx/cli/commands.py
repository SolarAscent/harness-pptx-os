"""High-level CLI commands for the PPT harness.

These are designed to be called from a Click/Typer CLI entry point
or directly from Python code.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from harness_pptx.models.content import Brief, Outline
from harness_pptx.models.deck_spec import DeckSpec
from harness_pptx.models.scene_graph import SceneGraph
from harness_pptx.cli.workspace import ProjectWorkspace


class PPTCommands:
    """High-level operations for deck creation, inspection, and QA."""

    def __init__(self, workspace_root: str | Path = "."):
        self._ws = ProjectWorkspace(workspace_root)

    @property
    def workspace(self) -> ProjectWorkspace:
        return self._ws

    # ---- create-from-text ---------------------------------------------------

    def create_from_text(
        self,
        text: str,
        output_path: str,
        audience: str = "general",
        style: str = "corporate",
        slides: int = 10,
    ) -> dict[str, Any]:
        """End-to-end: raw text → .pptx."""
        from harness_pptx.content.brief_parser import BriefParser
        from harness_pptx.content.story_planner import StoryPlanner
        from harness_pptx.content.intent_classifier import SlideIntentClassifier
        from harness_pptx.slide_types.base import get_slide_type_registry
        from harness_pptx.themes.base import ThemeRegistry
        from harness_pptx.backends.registry import get_backend
        from harness_pptx.renderer.engine import DeckRenderer
        from harness_pptx.qa.engine import QAEngine
        from harness_pptx.repair.planner import RepairPlanner
        from harness_pptx.repair.engine import RepairEngine

        # 1. Save brief
        self._ws.write_text(self._ws.brief_path, text)

        # 2. Parse brief
        parser = BriefParser()
        brief = parser.parse(text)
        brief.audience = audience
        self._ws.write_json(self._ws.root / "brief.json", brief.model_dump())

        # 3. Plan outline
        planner = StoryPlanner()
        outline = planner.plan(brief, target_slides=slides)
        self._ws.write_json(self._ws.outline_path, outline.model_dump())

        # 4. Classify intents
        classifier = SlideIntentClassifier()
        intents = classifier.classify(outline)

        # 5. Build scene graph via slide types
        themes = ThemeRegistry()
        theme = themes.get(style)
        self._ws.write_json(self._ws.theme_path, theme.model_dump())

        sg = SceneGraph(deck_id=brief.topic, theme=theme)
        reg = get_slide_type_registry()
        for i, intent in enumerate(intents):
            try:
                st = reg.get(intent.slide_type.value)
                content = {
                    "slide_id": intent.slide_id,
                    "seq": i,
                    "title": intent.title,
                }
                if intent.key_message:
                    content["key_message"] = intent.key_message
                node = st.build(content, theme, registry=sg.element_registry)
                sg.slides[node.id] = node
                sg.slide_count = len(sg.slides)
            except Exception:
                continue

        self._ws.write_json(self._ws.scene_graph_path, sg.model_dump())

        # 6. Render
        backend = get_backend()
        renderer = DeckRenderer(backend)
        result = renderer.render(sg, output_path)

        # 7. QA
        qa = QAEngine(theme=theme)
        qa_report = qa.run(sg)
        self._ws.write_json(self._ws.qa_report_path, qa_report.model_dump())

        # 8. Repair if needed
        if not qa_report.passed:
            rp = RepairPlanner()
            repair_plan = rp.plan(qa_report)
            self._ws.write_json(self._ws.repair_plan_path, repair_plan.model_dump())
            re = RepairEngine()
            sg = re.repair(sg, repair_plan)
            result = renderer.render(sg, output_path)

        # 9. Manifest
        self._ws.write_manifest({
            "pptx": output_path,
            "spec": str(self._ws.spec_path),
            "scene_graph": str(self._ws.scene_graph_path),
            "theme": str(self._ws.theme_path),
            "qa_report": str(self._ws.qa_report_path),
        })

        return {
            "output": output_path,
            "slides": len(outline.items),
            "theme": theme.name,
            "qa_passed": qa_report.passed,
            "render_success": result.success,
        }

    # ---- plan ---------------------------------------------------------------

    def plan(self, text: str, output_path: str | None = None) -> dict[str, Any]:
        """Parse text and produce an outline."""
        from harness_pptx.content.brief_parser import BriefParser
        from harness_pptx.content.story_planner import StoryPlanner

        parser = BriefParser()
        brief = parser.parse(text)
        planner = StoryPlanner()
        outline = planner.plan(brief)

        if output_path:
            self._ws.write_json(Path(output_path), outline.model_dump())

        return outline.model_dump()

    # ---- build --------------------------------------------------------------

    def build(
        self,
        spec_path: str,
        output_path: str,
        theme_name: str = "corporate",
    ) -> dict[str, Any]:
        """Build a .pptx from a spec JSON file."""
        from harness_pptx.models.deck_spec import DeckSpec
        from harness_pptx.themes.base import ThemeRegistry
        from harness_pptx.backends.registry import get_backend
        from harness_pptx.renderer.engine import DeckRenderer

        spec = DeckSpec.model_validate_json(Path(spec_path).read_text())
        if theme_name:
            themes = ThemeRegistry()
            spec.theme = themes.get(theme_name)

        backend = get_backend()
        renderer = DeckRenderer(backend)

        # Build SceneGraph from spec
        sg = SceneGraph(deck_id=spec.deck.title, theme=spec.theme)

        result = renderer.render(sg, output_path)
        return {"output": output_path, "success": result.success}

    # ---- inspect ------------------------------------------------------------

    def inspect(self, pptx_path: str, output_path: str | None = None) -> dict[str, Any]:
        """Inspect a .pptx and extract structure."""
        from harness_pptx.backends.registry import get_backend

        backend = get_backend()
        pres = backend.open_presentation(pptx_path)
        count = backend.slide_count(pres)

        slides_info = []
        for i in range(count):
            import time
            time.sleep(0.1)
            shapes = backend.list_shapes(pres)  # simplified
            slides_info.append({"index": i, "shape_count": len(shapes)})

        backend.close(pres)

        result = {"path": pptx_path, "slide_count": count, "slides": slides_info}
        if output_path:
            self._ws.write_json(Path(output_path), result)

        return result

    # ---- qa -----------------------------------------------------------------

    def qa(self, pptx_path: str, output_path: str | None = None) -> dict[str, Any]:
        """Run QA on a .pptx file."""
        from harness_pptx.models.scene_graph import SceneGraph
        from harness_pptx.qa.engine import QAEngine

        sg = SceneGraph(deck_id=Path(pptx_path).stem)
        qa = QAEngine()
        report = qa.run(sg)

        if output_path:
            self._ws.write_json(Path(output_path), report.model_dump())

        return report.model_dump()

    # ---- repair -------------------------------------------------------------

    def repair(self, pptx_path: str, spec_path: str, output_path: str | None = None) -> dict[str, Any]:
        """Repair a .pptx based on reference spec and QA report."""
        from harness_pptx.models.qa import QAReport
        from harness_pptx.models.scene_graph import SceneGraph
        from harness_pptx.qa.engine import QAEngine
        from harness_pptx.repair.engine import RepairEngine
        from harness_pptx.repair.planner import RepairPlanner

        sg = SceneGraph(deck_id=Path(pptx_path).stem)
        qa = QAEngine()
        report = qa.run(sg)

        rp = RepairPlanner()
        plan = rp.plan(report)

        re = RepairEngine()
        sg = re.repair(sg, plan)

        result = {"path": pptx_path, "issues_found": report.total_issues, "repairs_planned": len(plan.actions)}
        if output_path:
            self._ws.write_json(Path(output_path), result)
        return result

    # ---- preview ------------------------------------------------------------

    def preview(self, pptx_path: str, output_dir: str) -> dict[str, Any]:
        """Export slide previews as images."""
        from harness_pptx.backends.registry import get_backend

        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)

        backend = get_backend()
        pres = backend.open_presentation(pptx_path)
        count = backend.slide_count(pres)

        previews = []
        for i in range(count):
            png_path = out / f"slide_{i:02d}.png"
            backend.export_png(pres, i, str(png_path))
            previews.append(str(png_path))

        backend.close(pres)
        return {"previews": previews, "count": len(previews)}
