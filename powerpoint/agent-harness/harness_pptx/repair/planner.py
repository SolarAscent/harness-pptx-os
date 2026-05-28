"""RepairPlanner — convert QA issues into ordered repair actions."""

from __future__ import annotations

from harness_pptx.models.qa import (
    QACategory,
    QAIssue,
    QAReport,
    RepairAction,
    RepairActionType,
    RepairPlan,
)


class RepairPlanner:
    """Plan repairs from a QA report."""

    def plan(self, report: QAReport) -> RepairPlan:
        actions: list[RepairAction] = []
        for issue in report.issues:
            action = self._map_issue(issue)
            if action:
                actions.append(action)
        return RepairPlan(deck_id=report.deck_id, actions=actions)

    def _map_issue(self, issue: QAIssue) -> RepairAction | None:
        mapping = {
            QACategory.TEXT_OVERFLOW: RepairActionType.SHRINK_TEXT,
            QACategory.ELEMENT_OVERLAP: RepairActionType.MOVE_ELEMENT,
            QACategory.MARGIN: RepairActionType.ADJUST_MARGIN,
            QACategory.FONT_SIZE: RepairActionType.SHRINK_TEXT,
            QACategory.SLIDE_DENSITY: RepairActionType.SPLIT_SLIDE,
            QACategory.IMAGE_STRETCH: RepairActionType.RESIZE_IMAGE,
            QACategory.CONTRAST: RepairActionType.INCREASE_CONTRAST,
            QACategory.COLOR_THEME: RepairActionType.ALIGN_GROUP,
            QACategory.PUNCTUATION: RepairActionType.FIX_PUNCTUATION,
            QACategory.CHART_COMPLETENESS: RepairActionType.ADD_CHART_LABELS,
            QACategory.PAGE_NUMBER: RepairActionType.ADD_PAGE_NUMBER,
        }

        action_type = mapping.get(issue.category)
        if action_type is None:
            return None

        return RepairAction(
            id=f"repair-{issue.id}",
            issue_id=issue.id,
            action=action_type,
            slide_id=issue.slide_id,
            element_id=issue.element_id,
            params=issue.detail,
            description=issue.message,
        )
