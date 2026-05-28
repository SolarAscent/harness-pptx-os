"""QAEngine — run all quality checks on a scene graph or rendered deck."""

from __future__ import annotations

from harness_pptx.models.qa import QACategory, QAIssue, QAReport, Severity
from harness_pptx.models.scene_graph import SceneGraph, SlideNode
from harness_pptx.models.theme import Theme


class QAEngine:
    """Run all quality checks on a scene graph.

    Checks are modular; each check function takes (SceneGraph, SlideNode)
    and returns a list of QAIssue objects.
    """

    def __init__(self, theme: Theme | None = None):
        self._theme = theme
        self._checks = [
            self._check_text_overflow,
            self._check_overlap,
            self._check_font_size,
            self._check_margin,
            self._check_slide_density,
            self._check_contrast,
            self._check_color_theme,
        ]

    def run(self, scene_graph: SceneGraph) -> QAReport:
        """Run all model-based checks and return a QAReport."""
        all_issues: list[QAIssue] = []

        for slide in scene_graph.slide_order():
            for check in self._checks:
                try:
                    issues = check(scene_graph, slide)
                    all_issues.extend(issues)
                except Exception:
                    continue

        report = QAReport(deck_id=scene_graph.deck_id, issues=all_issues)
        return report

    def run_visual(
        self,
        pptx_path: str,
        scene_graph: SceneGraph,
        aspects: list[str] | None = None,
    ) -> QAReport:
        """Run vision-based review on rendered slide images.

        Exports slides to PNG via backend, sends each to the vision API,
        and returns issues found by the AI.

        Args:
            pptx_path: Path to the rendered .pptx file.
            scene_graph: SceneGraph with slide metadata.
            aspects: Specific aspects to review (all if None).

        Returns:
            QAReport with vision-detected issues.
        """
        from harness_pptx.qa.visual_reviewer import VisualReviewer
        from harness_pptx.backends.registry import get_backend

        backend = get_backend()
        reviewer = VisualReviewer(backend=backend)
        visual_report = reviewer.review(pptx_path, scene_graph, aspects=aspects)

        all_issues = visual_report.to_issues()
        return QAReport(deck_id=scene_graph.deck_id, issues=all_issues)

    def run_full(
        self,
        scene_graph: SceneGraph,
        pptx_path: str | None = None,
        aspects: list[str] | None = None,
    ) -> QAReport:
        """Run both model-based and vision-based checks.

        Args:
            scene_graph: SceneGraph for model checks.
            pptx_path: Path to rendered .pptx (required for vision check).
            aspects: Specific visual aspects to check.

        Returns:
            Combined QAReport.
        """
        # Model-based checks
        model_report = self.run(scene_graph)

        # Vision-based checks (if pptx available)
        if pptx_path:
            vision_report = self.run_visual(pptx_path, scene_graph, aspects=aspects)
            all_issues = model_report.issues + vision_report.issues
        else:
            all_issues = list(model_report.issues)

        return QAReport(deck_id=scene_graph.deck_id, issues=all_issues)

    # ---- Individual checks --------------------------------------------------

    def _check_text_overflow(self, sg: SceneGraph, slide: SlideNode) -> list[QAIssue]:
        issues = []
        for eid in slide.all_element_ids():
            el = sg.get_element(eid)
            if el is None or el.bbox is None:
                continue
            # Heuristic: text length vs bbox capacity
            if hasattr(el, "text") and el.text and el.bbox:
                est_chars_per_line = el.bbox.w / (el.style.font_size or 14 * 0.55) if el.style.font_size else el.bbox.w / 7.7
                est_lines = max(1, len(el.text) / max(est_chars_per_line, 1))
                est_h = est_lines * (el.style.font_size or 14) * 1.3
                if est_h > el.bbox.h * 1.2:
                    issues.append(QAIssue(
                        id=f"overflow-{eid}",
                        slide_id=slide.id,
                        element_id=eid,
                        category=QACategory.TEXT_OVERFLOW,
                        severity=Severity.WARNING,
                        message=f"Text may overflow: est {est_h:.0f}pt > box {el.bbox.h:.0f}pt",
                    ))
        return issues

    def _check_overlap(self, sg: SceneGraph, slide: SlideNode) -> list[QAIssue]:
        issues = []
        elements = []
        for eid in slide.all_element_ids():
            el = sg.get_element(eid)
            if el and el.bbox:
                elements.append((eid, el.bbox))

        for i in range(len(elements)):
            for j in range(i + 1, len(elements)):
                eid_a, ba = elements[i]
                eid_b, bb = elements[j]
                if _bboxes_overlap(ba, bb):
                    issues.append(QAIssue(
                        id=f"overlap-{eid_a}-{eid_b}",
                        slide_id=slide.id,
                        category=QACategory.ELEMENT_OVERLAP,
                        severity=Severity.WARNING,
                        message=f"Elements {eid_a} and {eid_b} overlap",
                    ))
        return issues

    def _check_font_size(self, sg: SceneGraph, slide: SlideNode) -> list[QAIssue]:
        issues = []
        for eid in slide.all_element_ids():
            el = sg.get_element(eid)
            if el is None:
                continue
            fs = el.style.font_size
            if fs and fs < 9:
                issues.append(QAIssue(
                    id=f"fontsize-{eid}",
                    slide_id=slide.id,
                    element_id=eid,
                    category=QACategory.FONT_SIZE,
                    severity=Severity.WARNING,
                    message=f"Font size {fs}pt is below readable minimum (9pt)",
                ))
        return issues

    def _check_margin(self, sg: SceneGraph, slide: SlideNode) -> list[QAIssue]:
        issues = []
        for eid in slide.all_element_ids():
            el = sg.get_element(eid)
            if el is None or el.bbox is None:
                continue
            b = el.bbox
            if b.x < 20 or b.y < 20 or b.right > slide.canvas.w - 20 or b.bottom > slide.canvas.h - 20:
                issues.append(QAIssue(
                    id=f"margin-{eid}",
                    slide_id=slide.id,
                    element_id=eid,
                    category=QACategory.MARGIN,
                    severity=Severity.WARNING,
                    message=f"Element {eid} is within 20pt of canvas edge",
                ))
        return issues

    def _check_slide_density(self, sg: SceneGraph, slide: SlideNode) -> list[QAIssue]:
        element_count = len(slide.all_element_ids())
        if element_count > 15:
            return [QAIssue(
                id=f"density-{slide.id}",
                slide_id=slide.id,
                category=QACategory.SLIDE_DENSITY,
                severity=Severity.WARNING,
                message=f"Slide has {element_count} elements (>15) — consider splitting",
            )]
        return []

    def _check_contrast(self, sg: SceneGraph, slide: SlideNode) -> list[QAIssue]:
        issues = []
        for eid in slide.all_element_ids():
            el = sg.get_element(eid)
            if el is None or el.bbox is None:
                continue
            fg = getattr(el.style, "font_color", None)
            if fg is None:
                continue
            # Resolve token to hex via theme
            if self._theme and not fg.startswith("#"):
                try:
                    fg = self._theme.color(fg)
                except (AttributeError, KeyError):
                    continue
            bg = self._theme.colors.background if self._theme else "#FFFFFF"
            ratio = _wcag_contrast_ratio(fg, bg)
            if ratio < 3.0:
                issues.append(QAIssue(
                    id=f"contrast-{eid}",
                    slide_id=slide.id,
                    element_id=eid,
                    category=QACategory.CONTRAST,
                    severity=Severity.WARNING,
                    message=f"Low contrast {ratio:.1f}:1 between {fg} and {bg}",
                ))
        return issues

    def _check_color_theme(self, sg: SceneGraph, slide: SlideNode) -> list[QAIssue]:
        if self._theme is None:
            return []
        issues = []
        theme_colors = {
            self._theme.colors.primary,
            self._theme.colors.accent,
            self._theme.colors.background,
            self._theme.colors.text,
            self._theme.colors.muted,
            self._theme.colors.surface,
            self._theme.colors.border,
            self._theme.colors.success,
            self._theme.colors.warning,
            self._theme.colors.error,
        }
        for eid in slide.all_element_ids():
            el = sg.get_element(eid)
            if el is None or el.bbox is None:
                continue
            for attr in ("font_color", "fill_color", "line_color"):
                color = getattr(el.style, attr, None)
                if color is None:
                    continue
                if self._theme and not color.startswith("#"):
                    try:
                        color = self._theme.color(color)
                    except (AttributeError, KeyError):
                        continue
                if not color.startswith("#"):
                    continue
                if color.upper() not in {c.upper() for c in theme_colors}:
                    issues.append(QAIssue(
                        id=f"colortheme-{eid}-{attr}",
                        slide_id=slide.id,
                        element_id=eid,
                        category=QACategory.COLOR_THEME,
                        severity=Severity.WARNING,
                        message=f"Color {color} on {attr} not in theme palette",
                    ))
        return issues


def _bboxes_overlap(a, b) -> bool:
    """True if two BBox objects overlap."""
    return (a.x < b.x + b.w and a.x + a.w > b.x and
            a.y < b.y + b.h and a.y + a.h > b.y)


def _wcag_contrast_ratio(fg: str, bg: str) -> float:
    """Compute WCAG 2.0 contrast ratio between two hex colors."""
    def _luminance(hex_color: str) -> float:
        h = hex_color.lstrip("#")
        rgb = tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))
        linear = tuple(
            c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
            for c in rgb
        )
        return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]
    l1 = _luminance(fg)
    l2 = _luminance(bg)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)
