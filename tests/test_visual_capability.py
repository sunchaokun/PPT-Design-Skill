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


def test_refresh_writes_atomic_state_for_registered_case(tmp_path: Path) -> None:
    case = tmp_path / "cases" / "demo-case"
    case.mkdir(parents=True)
    (case / "marker.txt").write_text("v1", encoding="utf-8")
    index = tmp_path / "index.json"
    index.write_text(json.dumps({"schema_version": 1, "cases_root": "cases", "prototypes": [{
        "prototype_id": "demo-case-p01", "case_id": "demo-case", "case_root": "cases/demo-case"
    }]}), encoding="utf-8")
    cache = tmp_path / "cache.json"
    lock = tmp_path / "cache.lock"
    result = run_script("refresh_case_prototypes.py", "--root", str(tmp_path), "--index", str(index), "--cache", str(cache), "--lock", str(lock))
    assert result.returncode == 0, result.stdout + result.stderr
    state = json.loads(cache.read_text(encoding="utf-8"))
    assert state["records"]["demo-case-p01"]["status"] == "valid"
    assert not lock.exists()


def test_refresh_fails_cleanly_when_lock_is_busy(tmp_path: Path) -> None:
    index = tmp_path / "index.json"
    index.write_text(json.dumps({"schema_version": 1, "cases_root": "cases", "prototypes": []}), encoding="utf-8")
    lock = tmp_path / "cache.lock"
    lock.write_text("busy", encoding="utf-8")
    result = run_script("refresh_case_prototypes.py", "--root", str(tmp_path), "--index", str(index), "--cache", str(tmp_path / "cache.json"), "--lock", str(lock))
    assert result.returncode == 1
    assert "busy" in result.stderr or "busy" in result.stdout
