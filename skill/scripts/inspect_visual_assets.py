"""Print a compact inventory of visual-capability assets."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def read(path: Path) -> dict:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    packs = read(ROOT / "skill/references/visual-rendering-packs/rendering_packs_index.json")
    compositions = read(ROOT / "skill/references/composition-recipes/composition_index.json")
    prototypes = read(ROOT / "skill/templates/case-prototype-index.json")
    print(f"packs={len(packs.get('packs', []))}")
    print(f"compositions={len(compositions.get('entries', []))}")
    print(f"prototypes={len(prototypes.get('prototypes', []))}")
    print(f"blocked_prototypes={len(prototypes.get('blocked', []))}")
    registered_case_ids = {
        item.get("case_id")
        for group in (prototypes.get("prototypes", []), prototypes.get("blocked", []))
        for item in group
        if isinstance(item, dict) and item.get("case_id")
    }
    print(f"registered_cases={len(registered_case_ids)}")
    print(f"prototype_cases_root={prototypes.get('cases_root', '')}")
    print("status=inventory_only; no aesthetic score is produced")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
