"""Create an isolated presentation-task workspace from the bundled template."""

from __future__ import annotations

import argparse
import json
import re
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from shutil import copytree
from typing import NoReturn


SKILL_VERSION = "1.3"
TASK_TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "task-init"
SLUG_RE = re.compile(r"^[a-z0-9][a-z0-9_-]{0,63}$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Initialize an isolated presentation task directory."
    )
    parser.add_argument(
        "--project", required=True, type=Path,
        help="Existing project directory; no files outside it are touched.",
    )
    parser.add_argument(
        "--name", required=True,
        help="Task slug used below <project>/ppt_tasks/.",
    )
    parser.add_argument(
        "--destination", type=Path,
        help="Optional destination inside the project directory.",
    )
    return parser.parse_args()


def fail(message: str) -> NoReturn:
    print(f"[ERROR] {message}", file=sys.stderr)
    raise SystemExit(2)


def main() -> int:
    args = parse_args()
    project = args.project.expanduser().resolve()
    if not project.is_dir():
        fail(f"project directory does not exist: {project}")
    if not SLUG_RE.fullmatch(args.name):
        fail("name must be a lowercase slug using letters, digits, '-' or '_'")
    if not TASK_TEMPLATE.is_dir():
        fail(f"task template is missing: {TASK_TEMPLATE}")

    if args.destination:
        destination = args.destination.expanduser().resolve()
        try:
            destination.relative_to(project)
        except ValueError:
            fail("destination must be inside the project directory")
    else:
        destination = project / "ppt_tasks" / args.name

    if destination.exists():
        fail(f"refusing to overwrite existing task: {destination}")

    destination.parent.mkdir(parents=True, exist_ok=True)
    copytree(TASK_TEMPLATE, destination)
    manifest = {
        "task_id": str(uuid.uuid4()),
        "task_name": args.name,
        "skill": "ppt-design-skill",
        "skill_version": SKILL_VERSION,
        "template_source": "skill/templates/task-init",
        "project_root": str(project),
        "task_root": str(destination),
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
    (destination / ".task-manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"[OK] initialized task -> {destination}")
    print(json.dumps(manifest, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
