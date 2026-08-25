"""Skill roots for supported coding assistants."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Platform:
    key: str
    display_name: str
    global_root: str
    project_root: str


PLATFORMS = {
    "claude": Platform("claude", "Claude Code", ".claude/skills", ".claude/skills"),
    "codex": Platform("codex", "Codex", ".agents/skills", ".codex/skills"),
    "opencode": Platform("opencode", "OpenCode", ".config/opencode/skills", ".opencode/skills"),
    "deepseek-harness": Platform("deepseek-harness", "DeepSeek Harness", ".dsh/skills", ".dsh/skills"),
    "cursor": Platform("cursor", "Cursor", ".agents/skills", ".cursor/skills"),
    "windsurf": Platform("windsurf", "Windsurf", ".windsurf/skills", ".windsurf/skills"),
    "roocode": Platform("roocode", "Roo Code", ".roo/skills", ".roo/skills"),
    "gemini": Platform("gemini", "Gemini CLI", ".gemini/skills", ".gemini/skills"),
    "trae": Platform("trae", "Trae", ".trae/skills", ".trae/skills"),
    "continue": Platform("continue", "Continue", ".continue/skills", ".continue/skills"),
    "droid": Platform("droid", "Droid", ".factory/skills", ".factory/skills"),
    "kilocode": Platform("kilocode", "KiloCode", ".kilocode/skills", ".kilocode/skills"),
    "augment": Platform("augment", "Augment", ".augment/skills", ".augment/skills"),
    "copilot": Platform("copilot", "GitHub Copilot", ".github/skills", ".github/skills"),
}

ALIASES = {
    "deepseek": "deepseek-harness",
    "deepseek-harness": "deepseek-harness",
    "deepseekhaness": "deepseek-harness",
    "deepseek_harness": "deepseek-harness",
    "claude-code": "claude",
    "github-copilot": "copilot",
}


def normalize(name: str) -> str:
    key = name.strip().lower()
    return ALIASES.get(key, key)


def global_path(platform: Platform, home: Path) -> Path:
    return home / Path(platform.global_root)


def project_path(platform: Platform, root: Path) -> Path:
    return root / Path(platform.project_root)
