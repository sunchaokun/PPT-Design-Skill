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
    missing_paths = []
    for group in (packs.get("packs", []), compositions.get("entries", []), prototypes.get("blocked", []), prototypes.get("prototypes", [])):
        for item in group:
            if isinstance(item, dict) and item.get("path") and not (ROOT / item["path"]).is_file():
                missing_paths.append(item["path"])
    print(f"missing_index_paths={len(missing_paths)}")
    family_ids = {item.get("composition_id") for item in compositions.get("entries", []) if isinstance(item, dict)}
    pack_assets = [read(ROOT / item["path"]) for item in packs.get("packs", []) if isinstance(item, dict) and item.get("path")]
    referenced_families = {family for pack in pack_assets for family in pack.get("composition_families", [])}
    print(f"unreferenced_compositions={sorted(family_ids - referenced_families)}")
    roles_by_pack = {pack.get("pack_id"): set() for pack in packs.get("packs", []) if isinstance(pack, dict)}
    for group in (prototypes.get("prototypes", []), prototypes.get("blocked", [])):
        for item in group:
            record = item
            if isinstance(item, dict) and item.get("path"):
                record = read(ROOT / item["path"])
            if isinstance(record, dict) and record.get("pack_id") in roles_by_pack:
                roles_by_pack[record["pack_id"]].update(record.get("page_roles", []))
    print(f"pack_role_coverage={json.dumps({key: sorted(value) for key, value in roles_by_pack.items()}, ensure_ascii=False, sort_keys=True)}")
    role_gaps = sorted(key for key, roles in roles_by_pack.items() if len(roles) < 3)
    print(f"pack_role_gaps={role_gaps}")
    print(f"prototype_cases_root={prototypes.get('cases_root', '')}")
    print("status=inventory_only; no aesthetic score is produced")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
