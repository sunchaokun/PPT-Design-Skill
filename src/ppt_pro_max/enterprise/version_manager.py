"""Version manager — output directory and version numbering."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def next_version(output_dir: str) -> int:
    existing: list[int] = []
    output = Path(output_dir)
    if not output.is_dir():
        return 1
    for entry in output.iterdir():
        if entry.is_dir() and entry.name.startswith("v") and entry.name[1:].isdigit():
            existing.append(int(entry.name[1:]))
    return max(existing, default=0) + 1


def write_meta(version_dir: str, meta: dict[str, Any]) -> None:
    vdir = Path(version_dir)
    vdir.mkdir(parents=True, exist_ok=True)
    (vdir / "meta.json").write_text(
        json.dumps(meta, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def read_meta(version_dir: str) -> dict[str, Any] | None:
    meta_path = Path(version_dir) / "meta.json"
    if not meta_path.is_file():
        return None
    try:
        return json.loads(meta_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
