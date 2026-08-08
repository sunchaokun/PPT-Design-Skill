"""SKILL.md template renderer — platform-specific content generation."""

from __future__ import annotations

from pathlib import Path


def render_skill_md(skill_md_path: Path, platform_config: dict) -> str:
    content = skill_md_path.read_text(encoding="utf-8")
    frontmatter = platform_config.get("frontmatter", {})
    if not frontmatter:
        return content
    lines = content.splitlines(keepends=True)
    in_frontmatter = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped == "---":
            if in_frontmatter:
                break
            in_frontmatter = True
            continue
        if in_frontmatter and ":" in stripped:
            key = stripped.split(":", 1)[0].strip()
            if key in frontmatter:
                lines[i] = f"{key}: {frontmatter[key]}\n"
    return "".join(lines)
