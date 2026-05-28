"""RoleAssigner — assign narrative roles to each slide."""

from __future__ import annotations

from harness_pptx.models.content import SlideIntent
from harness_pptx.narrative.frameworks import NarrativeFramework


class RoleAssigner:
    """Assign narrative roles to slide intents based on a framework."""

    def assign(self, intents: list[SlideIntent], framework: NarrativeFramework) -> list[SlideIntent]:
        """Annotate intents with narrative roles from the framework."""
        phases = framework.phases
        if not phases:
            return intents

        for i, intent in enumerate(intents):
            # First → opening, Last → closing
            if i == 0:
                intent.narrative_role = "opening"
            elif i == len(intents) - 1:
                intent.narrative_role = "closing"
            else:
                # Distribute across body phases
                body_phases = [p for p in phases if p.role not in ("opening", "closing")]
                if body_phases:
                    body_items = len(intents) - 2
                    per_phase = max(1, body_items // len(body_phases))
                    phase_idx = min((i - 1) // per_phase, len(body_phases) - 1)
                    intent.narrative_role = body_phases[phase_idx].role

        return intents
