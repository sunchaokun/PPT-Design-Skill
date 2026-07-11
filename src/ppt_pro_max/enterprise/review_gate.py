"""ReviewGate — user confirmation mechanism."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class ReviewGate:

    def generate_proposal(
        self,
        project_dir: str,
        brand_spec: dict[str, Any],
        story_plan: dict[str, Any],
        page_designs: list[dict[str, Any]],
        assets: dict[str, Any],
    ) -> dict[str, Any]:
        return {
            "project_dir": project_dir,
            "brand_spec": brand_spec,
            "story_plan": story_plan,
            "pages": page_designs,
            "assets": assets,
        }

    def write_proposal(self, proposal: dict[str, Any], path: str) -> None:
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(
            json.dumps(proposal, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def read_proposal(self, path: str) -> dict[str, Any] | None:
        try:
            return json.loads(Path(path).read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError, FileNotFoundError):
            return None

    def format_cli(self, proposal: dict[str, Any]) -> str:
        lines = []
        lines.append("=" * 50)
        lines.append("  PPT 设计方案")
        lines.append("=" * 50)

        bs = proposal.get("brand_spec", {})
        colors = bs.get("colors", {})
        if colors:
            lines.append("")
            lines.append("设计规范")
            for k, v in colors.items():
                lines.append(f"  {k}: {v}")

        assets = proposal.get("assets", {})
        if assets:
            lines.append("")
            lines.append("资产识别结果")
            for k, v in assets.items():
                mark = "✓" if v else "✗"
                lines.append(f"  {mark} {k}: {v}")

        pages = proposal.get("pages", [])
        if pages:
            lines.append("")
            lines.append(f"页面规划 ({len(pages)}页)")
            for i, p in enumerate(pages, 1):
                notes_preview = ""
                if p.get("notes"):
                    notes_preview = f" [{p['notes'][:30]}...]"
                lines.append(f"  {i}. {p.get('goal', '?'):12s} → {p.get('title', '')}{notes_preview}")

        lines.append("")
        lines.append("请确认方案 (y/n/edit):")
        return "\n".join(lines)
