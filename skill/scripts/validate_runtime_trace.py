"""Validate runtime trace structure and gate consistency."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
from validate_visual_pack import ValidationError, check_schema, load_json  # noqa: E402


def validate_trace(trace: dict) -> None:
    schema = load_json(ROOT / "skill/schemas/runtime-trace.schema.json")
    check_schema(trace, schema, root=schema)
    for field in ("prototype_ids", "recipe_ids", "composition_ids"):
        values = trace.get(field, [])
        if len(values) != len(set(values)):
            raise ValidationError(f"{field} must not contain duplicates")
    p01 = trace["p01_gate"]
    final = trace["final_visual_gate"]
    if p01 == "BLOCKED" and final == "PASS":
        raise ValidationError("final_visual_gate cannot PASS when p01_gate is BLOCKED")
    if final == "PASS" and p01 != "PASS":
        raise ValidationError("final_visual_gate PASS requires p01_gate PASS")
    if final != "PASS" and not trace.get("failure_reason"):
        raise ValidationError("failure_reason is required when final_visual_gate is not PASS")
    if trace["mode"] == "vi_build" and not trace.get("package_module_path"):
        raise ValidationError("vi_build traces require package_module_path")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("trace", type=Path)
    args = parser.parse_args()
    try:
        value = json.loads(args.trace.read_text(encoding="utf-8"))
        if not isinstance(value, dict):
            raise ValidationError("trace root must be an object")
        validate_trace(value)
    except (OSError, json.JSONDecodeError, ValidationError) as exc:
        print(f"ERROR: {exc}")
        return 1
    print(f"OK: runtime trace is valid: {args.trace}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
