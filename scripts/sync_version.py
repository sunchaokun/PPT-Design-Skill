#!/usr/bin/env python3
"""Sync version number across all manifest files."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

MANIFEST_FILES = [
    "skill.json",
    ".claude-plugin/plugin.json",
    ".claude-plugin/marketplace.json",
    "pyproject.toml",
    "skill/SKILL.md",
    "src/ppt_pro_max/__init__.py",
]


def _sync_yml(rel_path: str, file_path: Path, new_version: str) -> None:
    """Sync version in YAML front-matter (e.g. skill/SKILL.md 'version: X.Y.Z')."""
    content = file_path.read_text(encoding="utf-8")
    old_match = re.search(r'^version:\s*[\d.]+', content, re.MULTILINE)
    old_version = old_match.group(0) if old_match else "?"
    content = re.sub(r'^version:\s*[\d.]+', f"version: {new_version}", content, count=1, flags=re.MULTILINE)
    file_path.write_text(content, encoding="utf-8")
    print(f"  [OK] {rel_path}: {old_version} -> {new_version}")


def _sync_py(rel_path: str, file_path: Path, new_version: str) -> None:
    """Sync __version__ in a Python file."""
    content = file_path.read_text(encoding="utf-8")
    old_match = re.search(r'__version__\s*=\s*"([^"]+)"', content)
    old_version = old_match.group(1) if old_match else "?"
    content = re.sub(r'__version__\s*=\s*"[^"]+"', f'__version__ = "{new_version}"', content, count=1)
    file_path.write_text(content, encoding="utf-8")
    print(f"  [OK] {rel_path}: {old_version} -> {new_version}")


def sync_version(new_version: str) -> None:
    project_root = Path(__file__).resolve().parent.parent

    for rel_path in MANIFEST_FILES:
        file_path = project_root / rel_path
        if not file_path.exists():
            print(f"  [SKIP] {rel_path} not found")
            continue

        content = file_path.read_text(encoding="utf-8")

        if rel_path.endswith(".json"):
            data = json.loads(content)
            old_version = data.get("version", "")
            data["version"] = new_version
            if "metadata" in data and "version" in data["metadata"]:
                data["metadata"]["version"] = new_version
            if "plugins" in data:
                for plugin in data["plugins"]:
                    if "version" in plugin:
                        plugin["version"] = new_version
            file_path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
            print(f"  [OK] {rel_path}: {old_version} -> {new_version}")
        elif rel_path.endswith(".toml"):
            old_match = re.search(r'^version\s*=\s*"([^"]+)"', content, re.MULTILINE)
            old_version = old_match.group(1) if old_match else "?"
            content = re.sub(r'^version\s*=\s*"[^"]+"', f'version = "{new_version}"',
                             content, count=1, flags=re.MULTILINE)
            file_path.write_text(content, encoding="utf-8")
            print(f"  [OK] {rel_path}: {old_version} -> {new_version}")
        elif rel_path.endswith(".md"):
            _sync_yml(rel_path, file_path, new_version)
        elif rel_path.endswith(".py"):
            _sync_py(rel_path, file_path, new_version)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python scripts/sync_version.py <version>")
        print("Example: python scripts/sync_version.py 0.10.0")
        sys.exit(1)
    sync_version(sys.argv[1])
