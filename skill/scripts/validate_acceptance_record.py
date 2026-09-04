"""Validate an acceptance record's structure and scoring semantics."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
from validate_visual_pack import ValidationError, check_schema, load_json  # noqa: E402


def validate_record(record: dict) -> None:
    schema = load_json(ROOT / "skill/schemas/acceptance-record.schema.json")
    check_schema(record, schema, root=schema)
    criteria = record["criteria"]
    ids = [item["id"] for item in criteria]
    if len(ids) != len(set(ids)):
        raise ValidationError("criteria IDs must be unique")
    score_sum = sum(item["score"] for item in criteria)
    if record["score"] != score_sum:
        raise ValidationError(f"score must equal criteria sum ({score_sum})")
    required_ids = {"visual_thesis", "composition", "material", "readability", "editability", "reproducibility"}
    if set(ids) != required_ids:
        raise ValidationError(f"criteria IDs must be exactly {sorted(required_ids)}")
    if record["status"] == "PASS":
        if record["score"] < 9:
            raise ValidationError("PASS requires score >= 9")
        lower_bounds = {item["id"]: item["score"] for item in criteria}
        if lower_bounds["readability"] < 1 or lower_bounds["editability"] < 1:
            raise ValidationError("PASS requires readability and editability >= 1")
    elif record["status"] == "NEEDS_REVISION" and record["score"] >= 9:
        raise ValidationError("score >= 9 cannot be marked NEEDS_REVISION")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("record", type=Path)
    args = parser.parse_args()
    try:
        value = json.loads(args.record.read_text(encoding="utf-8"))
        if not isinstance(value, dict):
            raise ValidationError("record root must be an object")
        validate_record(value)
    except (OSError, json.JSONDecodeError, ValidationError) as exc:
        print(f"ERROR: {exc}")
        return 1
    print(f"OK: acceptance record is valid: {args.record}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
