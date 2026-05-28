"""ErrorIsolation — per-slide build with failure isolation."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class IsolatedResult:
    success: bool = True
    failed_slides: list[str] = field(default_factory=list)
    failed_elements: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


class ErrorIsolation:
    """Run an operation per-slide; failure on one slide doesn't stop others."""

    def run_per_slide(
        self,
        slide_ids: list[str],
        handler: callable,
    ) -> IsolatedResult:
        result = IsolatedResult()
        for sid in slide_ids:
            try:
                handler(sid)
            except Exception as e:
                result.failed_slides.append(sid)
                result.errors.append(f"Slide {sid}: {e}")
        result.success = len(result.failed_slides) == 0
        return result
