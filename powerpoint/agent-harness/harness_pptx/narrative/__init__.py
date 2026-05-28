"""Narrative package."""

from harness_pptx.narrative.frameworks import (
    NarrativeFramework,
    get_framework,
    list_frameworks,
)
from harness_pptx.narrative.role_assigner import RoleAssigner
from harness_pptx.narrative.notes_writer import SpeakerNotesWriter

__all__ = [
    "NarrativeFramework",
    "get_framework",
    "list_frameworks",
    "RoleAssigner",
    "SpeakerNotesWriter",
]
