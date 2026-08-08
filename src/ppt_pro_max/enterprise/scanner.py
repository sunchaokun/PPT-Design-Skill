"""ProjectScanner — scan project folder for assets."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional


_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".svg", ".webp"}

_LOGO_PATTERN = re.compile(r"^logo([_-]|$)", re.IGNORECASE)


@dataclass
class ProjectAsset:
    project_dir: str
    template_path: Optional[str] = None
    logo_path: Optional[str] = None
    brand_raw: Optional[dict[str, Any]] = None
    content_raw: Optional[dict[str, Any]] = None
    readme_path: Optional[str] = None
    image_pool: list[str] = field(default_factory=list)


_TEMPLATE_PATTERN = re.compile(r"template", re.IGNORECASE)


class ProjectScanner:

    def scan(self, project_dir: str) -> ProjectAsset:
        project = Path(project_dir)
        asset = ProjectAsset(project_dir=project_dir)

        if not project.is_dir():
            return asset

        output_dir = project / "output"
        pptx_candidates: list[Path] = []

        for entry in sorted(project.iterdir()):
            if not entry.is_file():
                continue
            name = entry.name
            stem = entry.stem
            suffix = entry.suffix.lower()

            if suffix == ".pptx" and entry.parent != output_dir:
                pptx_candidates.append(entry)
            elif self._is_logo(stem, suffix):
                asset.logo_path = str(entry)
            elif name == "brand.json":
                asset.brand_raw = self._load_json(entry)
            elif name == "content.json":
                asset.content_raw = self._load_json(entry)
            elif name.lower() in ("readme.md", "readme.markdown", "readme"):
                asset.readme_path = str(entry)
            elif suffix in _IMAGE_EXTENSIONS:
                asset.image_pool.append(str(entry))

        if pptx_candidates:
            template = self._select_template(pptx_candidates)
            if template:
                asset.template_path = str(template)

        return asset

    def _select_template(self, candidates: list[Path]) -> Path | None:
        for c in candidates:
            if _TEMPLATE_PATTERN.search(c.stem):
                return c
        return candidates[0]

    def _is_logo(self, stem: str, suffix: str) -> bool:
        if suffix not in _IMAGE_EXTENSIONS:
            return False
        return bool(_LOGO_PATTERN.match(stem))

    def _load_json(self, path: Path) -> dict[str, Any] | None:
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return None
