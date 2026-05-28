"""ManifestGenerator — generate artifact manifests."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any


class ManifestGenerator:
    """Generate a manifest listing all generated artifacts."""

    def generate(self, artifacts: dict[str, str], output_path: str | Path) -> dict[str, Any]:
        manifest = {
            "generated_at": datetime.now().isoformat(),
            "artifacts": artifacts,
        }
        path = Path(output_path)
        path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False))
        return manifest
