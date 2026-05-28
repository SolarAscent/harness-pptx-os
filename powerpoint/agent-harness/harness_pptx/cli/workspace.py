"""ProjectWorkspace — manage the project directory structure for a deck."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any


class ProjectWorkspace:
    """Manages a deck project directory with standard sub-folders.

    Structure::

        project/
          brief.md
          outline.json
          deck.spec.json
          scene.graph.json
          theme.json
          assets/
          previews/
          qa/
            report.json
            repair_plan.json
          exports/
          logs/
            build.log
            qa.log
          manifest.json
    """

    SUBDIRS = ["assets", "previews", "qa", "exports", "logs"]

    def __init__(self, root: str | Path):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)
        for sub in self.SUBDIRS:
            (self.root / sub).mkdir(exist_ok=True)

    # ---- Path helpers -------------------------------------------------------

    @property
    def brief_path(self) -> Path:
        return self.root / "brief.md"

    @property
    def outline_path(self) -> Path:
        return self.root / "outline.json"

    @property
    def spec_path(self) -> Path:
        return self.root / "deck.spec.json"

    @property
    def scene_graph_path(self) -> Path:
        return self.root / "scene.graph.json"

    @property
    def theme_path(self) -> Path:
        return self.root / "theme.json"

    @property
    def qa_dir(self) -> Path:
        return self.root / "qa"

    @property
    def qa_report_path(self) -> Path:
        return self.qa_dir / "report.json"

    @property
    def repair_plan_path(self) -> Path:
        return self.qa_dir / "repair_plan.json"

    @property
    def exports_dir(self) -> Path:
        return self.root / "exports"

    @property
    def logs_dir(self) -> Path:
        return self.root / "logs"

    @property
    def manifest_path(self) -> Path:
        return self.root / "manifest.json"

    # ---- Read / write helpers -----------------------------------------------

    def write_json(self, path: Path, data: Any) -> None:
        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )

    def read_json(self, path: Path) -> Any:
        return json.loads(path.read_text(encoding="utf-8"))

    def write_text(self, path: Path, text: str) -> None:
        path.write_text(text, encoding="utf-8")

    def read_text(self, path: Path) -> str:
        return path.read_text(encoding="utf-8")

    # ---- Manifest -----------------------------------------------------------

    def write_manifest(self, artifacts: dict[str, str]) -> None:
        manifest = {
            "created": datetime.now().isoformat(),
            "artifacts": artifacts,
        }
        self.write_json(self.manifest_path, manifest)
