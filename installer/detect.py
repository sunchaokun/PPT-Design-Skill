"""Platform detection logic for AI coding assistants."""

from __future__ import annotations

from pathlib import Path

PLATFORMS: dict[str, dict[str, str]] = {
    "claude": {
        "project_dir": ".claude",
        "global_path": ".claude/skills",
        "desc": "Claude Code",
    },
    "opencode": {
        "project_dir": ".opencode",
        "global_path": ".config/opencode/skills",
        "desc": "OpenCode",
    },
    "codex": {
        "project_dir": ".codex",
        "global_path": ".agents/skills",
        "desc": "Codex / OpenClaw",
    },
    "cursor": {
        "project_dir": ".cursor",
        "global_path": ".agents/plugins",
        "desc": "Cursor",
    },
    "windsurf": {
        "project_dir": ".windsurf",
        "global_path": ".windsurf/skills",
        "desc": "Windsurf",
    },
    "roocode": {
        "project_dir": ".roo",
        "global_path": ".roo/skills",
        "desc": "RooCode",
    },
    "gemini": {
        "project_dir": ".gemini",
        "global_path": ".gemini/skills",
        "desc": "Gemini CLI",
    },
    "trae": {
        "project_dir": ".trae",
        "global_path": ".trae/skills",
        "desc": "Trae",
    },
    "continue": {
        "project_dir": ".continue",
        "global_path": ".continue/skills",
        "desc": "Continue",
    },
    "droid": {
        "project_dir": ".factory",
        "global_path": ".factory/skills",
        "desc": "Droid (Factory)",
    },
    "kilocode": {
        "project_dir": ".kilocode",
        "global_path": ".kilocode/skills",
        "desc": "KiloCode",
    },
    "augment": {
        "project_dir": ".augment",
        "global_path": ".augment/skills",
        "desc": "Augment",
    },
    "copilot": {
        "project_dir": ".github",
        "global_path": ".github/skills",
        "desc": "GitHub Copilot",
    },
}


def detect_project_platforms(target_dir: Path) -> list[str]:
    detected: list[str] = []
    for name, info in PLATFORMS.items():
        if (target_dir / info["project_dir"]).exists():
            detected.append(name)
    return detected


def detect_global_platforms() -> list[str]:
    detected: list[str] = []
    for name, info in PLATFORMS.items():
        global_base = Path.home() / info["global_path"]
        if global_base.exists():
            detected.append(name)
    return detected
