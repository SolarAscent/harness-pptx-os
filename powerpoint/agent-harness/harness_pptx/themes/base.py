"""Theme registry — load, validate, and query design themes."""

from __future__ import annotations

import json
from pathlib import Path

from harness_pptx.models.theme import Theme


class ThemeRegistry:
    """Registry of named Theme presets, loaded from JSON files."""

    def __init__(self, presets_dir: Path | None = None):
        if presets_dir is None:
            presets_dir = Path(__file__).resolve().parent / "presets"
        self._presets_dir = Path(presets_dir)
        self._cache: dict[str, Theme] = {}

    @property
    def presets_dir(self) -> Path:
        return self._presets_dir

    def list(self) -> list[str]:
        """Return sorted list of available theme names."""
        if not self._presets_dir.is_dir():
            return []
        names = sorted(
            p.stem for p in self._presets_dir.glob("*.json")
        )
        return names

    def get(self, name: str) -> Theme:
        """Load a theme by name. Caches after first load."""
        if name in self._cache:
            return self._cache[name]

        path = self._presets_dir / f"{name}.json"
        if not path.is_file():
            available = self.list()
            raise FileNotFoundError(
                f"Theme '{name}' not found. Available: {available}. "
                f"Expected path: {path}"
            )

        data = json.loads(path.read_text(encoding="utf-8"))
        theme = Theme.model_validate(data)
        self._cache[name] = theme
        return theme

    def load_all(self) -> dict[str, Theme]:
        """Pre-load all presets from disk."""
        for name in self.list():
            self.get(name)
        return dict(self._cache)

    def register(self, theme: Theme) -> None:
        """Programmatically register a theme (not persisted)."""
        self._cache[theme.name] = theme

    def default(self) -> Theme:
        """Return the default (corporate) theme."""
        return self.get("corporate")
