from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def run_script(name: str, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(ROOT / "skill/scripts" / name), *args],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )


def test_visual_schemas_and_empty_indexes_are_valid() -> None:
    result = run_script("validate_visual_pack.py")
    assert result.returncode == 0, result.stdout + result.stderr


def test_asset_inventory_is_machine_readable() -> None:
    result = run_script("inspect_visual_assets.py")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "packs=" in result.stdout
    assert "status=inventory_only" in result.stdout


def test_workflow_state_machine_supports_resume(tmp_path: Path) -> None:
    state = tmp_path / "state.json"
    for action in ("route", "directions", "confirm", "p01", "pass", "review", "pass"):
        result = run_script("run_visual_workflow.py", str(state), action)
        assert result.returncode == 0, result.stdout + result.stderr
    data = json.loads(state.read_text(encoding="utf-8"))
    assert data["status"] == "PASS"
    assert len(data["history"]) == 7


def test_workflow_rejects_invalid_transition(tmp_path: Path) -> None:
    state = tmp_path / "state.json"
    result = run_script("run_visual_workflow.py", str(state), "pass")
    assert result.returncode == 1
    assert "invalid" in result.stdout
