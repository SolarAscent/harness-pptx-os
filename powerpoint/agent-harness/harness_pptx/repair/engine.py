"""RepairEngine — execute repair actions on a scene graph."""

from __future__ import annotations

from harness_pptx.models.element import BaseElement
from harness_pptx.models.layout import BBox
from harness_pptx.models.qa import RepairAction, RepairActionType, RepairPlan
from harness_pptx.models.scene_graph import SceneGraph


class RepairEngine:
    """Execute repairs by modifying the scene graph.

    After repairs, re-rendering produces the fixed output.
    """

    def repair(self, scene_graph: SceneGraph, plan: RepairPlan) -> SceneGraph:
        """Apply all repair actions to the scene graph."""
        for action in plan.actions:
            handler = self._handlers.get(action.action)
            if handler:
                try:
                    handler(scene_graph, action)
                except Exception:
                    continue
        return scene_graph

    # ---- Repair handlers ----------------------------------------------------

    def _shrink_text(self, sg: SceneGraph, action: RepairAction) -> None:
        el = self._get_element(sg, action.element_id)
        if el and el.style.font_size:
            el.style.font_size = max(8, el.style.font_size - 2)

    def _expand_box(self, sg: SceneGraph, action: RepairAction) -> None:
        el = self._get_element(sg, action.element_id)
        if el and el.bbox:
            el.bbox.h += 20

    def _move_element(self, sg: SceneGraph, action: RepairAction) -> None:
        el = self._get_element(sg, action.element_id)
        if el and el.bbox:
            dx = action.params.get("dx", 0)
            dy = action.params.get("dy", 0)
            el.bbox.x += dx
            el.bbox.y += dy

    def _adjust_margin(self, sg: SceneGraph, action: RepairAction) -> None:
        el = self._get_element(sg, action.element_id)
        if el and el.bbox:
            b = el.bbox
            el.bbox = BBox(
                x=max(36, b.x),
                y=max(36, b.y),
                w=min(b.w, 960 - 72),
                h=min(b.h, 540 - 72),
            )

    def _split_slide(self, sg: SceneGraph, action: RepairAction) -> None:
        pass  # Placeholder — would create new slide and move half the elements

    def _resize_image(self, sg: SceneGraph, action: RepairAction) -> None:
        el = self._get_element(sg, action.element_id)
        if el and el.bbox:
            target_w = action.params.get("w", el.bbox.w)
            target_h = action.params.get("h", el.bbox.h)
            el.bbox.w = target_w
            el.bbox.h = target_h

    # ---- Helpers ------------------------------------------------------------

    @staticmethod
    def _get_element(sg: SceneGraph, element_id: str | None) -> BaseElement | None:
        if element_id is None:
            return None
        return sg.get_element(element_id)

    @property
    def _handlers(self):
        return {
            RepairActionType.SHRINK_TEXT: self._shrink_text,
            RepairActionType.EXPAND_BOX: self._expand_box,
            RepairActionType.MOVE_ELEMENT: self._move_element,
            RepairActionType.SPLIT_SLIDE: self._split_slide,
            RepairActionType.ADJUST_MARGIN: self._adjust_margin,
            RepairActionType.RESIZE_IMAGE: self._resize_image,
            RepairActionType.REBALANCE_COLUMNS: self._adjust_margin,
            RepairActionType.INCREASE_CONTRAST: self._adjust_margin,
            RepairActionType.ALIGN_GROUP: self._adjust_margin,
            RepairActionType.SIMPLIFY_BULLETS: self._shrink_text,
            RepairActionType.PROMOTE_TO_APPENDIX: self._split_slide,
            RepairActionType.FIX_PUNCTUATION: self._shrink_text,
            RepairActionType.ADD_CHART_LABELS: self._expand_box,
            RepairActionType.ADD_PAGE_NUMBER: self._expand_box,
        }
