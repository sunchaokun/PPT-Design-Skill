from __future__ import annotations

import json
from pathlib import Path

from skill.scripts.validate_visual_pack import check_schema, load_json


ROOT = Path(__file__).resolve().parents[1]


def test_all_declared_pack_assets_match_their_index() -> None:
    index = load_json(ROOT / "skill/references/visual-rendering-packs/rendering_packs_index.json")
    assert len(index["packs"]) == 2
    for item in index["packs"]:
        asset = load_json(ROOT / item["path"])
        assert asset["pack_id"] == item["pack_id"]


def test_all_declared_composition_assets_match_their_index() -> None:
    index = load_json(ROOT / "skill/references/composition-recipes/composition_index.json")
    assert len(index["entries"]) == 4
    for item in index["entries"]:
        asset = load_json(ROOT / item["path"])
        assert asset["composition_id"] == item["composition_id"]


def test_invalid_rendering_pack_is_rejected() -> None:
    schema = load_json(ROOT / "skill/schemas/rendering-pack.schema.json")
    invalid = {"pack_id": "bad", "version": 1, "summary": "x", "compatible_domains": ["technical"], "composition_families": ["c1-01"], "anti_patterns": ["x"]}
    try:
        check_schema(invalid, schema)
    except ValueError as exc:
        assert "missing required field" in str(exc)
    else:
        raise AssertionError("incomplete rendering pack must fail validation")
