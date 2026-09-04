"""Audit baseline/upgraded regression pairs without inventing a baseline."""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DOMAINS = {"technical", "scientific", "brand_architecture"}


class AuditError(ValueError):
    pass


def slide_count(path: Path) -> int:
    try:
        with zipfile.ZipFile(path) as archive:
            return sum(bool(re.fullmatch(r"ppt/slides/slide\d+\.xml", name)) for name in archive.namelist())
    except (OSError, zipfile.BadZipFile) as exc:
        raise AuditError(f"invalid PPTX: {path}") from exc


def audit(root: Path, manifest_path: Path) -> tuple[list[dict], bool]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    entries = manifest.get("entries")
    if manifest.get("schema_version") != 1 or not isinstance(entries, list) or len(entries) != 3:
        raise AuditError("manifest must contain schema_version 1 and exactly three entries")
    seen_domains: set[str] = set()
    results: list[dict] = []
    all_ready = True
    for entry in entries:
        if entry.get("domain") not in DOMAINS:
            raise AuditError(f"unsupported domain: {entry.get('domain')}")
        if entry["domain"] in seen_domains:
            raise AuditError(f"duplicate domain: {entry['domain']}")
        seen_domains.add(entry["domain"])
        result = {"regression_id": entry.get("regression_id"), "domain": entry["domain"], "case_id": entry.get("case_id")}
        baseline = root / entry["baseline"] if entry.get("baseline") else None
        upgraded = root / entry["upgraded"] if entry.get("upgraded") else None
        reasons = []
        if baseline is None or not baseline.is_file():
            reasons.append("missing baseline PPTX")
        if upgraded is None or not upgraded.is_file():
            reasons.append("missing upgraded PPTX")
        if not reasons:
            base_pages = slide_count(baseline)
            upgraded_pages = slide_count(upgraded)
            result["baseline_slide_count"] = base_pages
            result["upgraded_slide_count"] = upgraded_pages
            if base_pages != upgraded_pages:
                reasons.append("baseline/upgraded slide counts differ")
        result["status"] = "PASS" if not reasons and entry.get("status") == "PASS" else "BLOCKED"
        result["blocking_reasons"] = reasons or ([] if result["status"] == "PASS" else [entry.get("reason", "manifest is not marked PASS")])
        all_ready = all_ready and result["status"] == "PASS"
        results.append(result)
    if seen_domains != DOMAINS:
        raise AuditError(f"manifest must cover domains {sorted(DOMAINS)}")
    return results, all_ready


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--manifest", type=Path, default=ROOT / "skill/templates/regression-manifest.json")
    args = parser.parse_args()
    try:
        results, all_ready = audit(args.root, args.manifest)
    except (OSError, json.JSONDecodeError, AuditError) as exc:
        print(f"ERROR: {exc}")
        return 1
    print(json.dumps({"entries": results, "status": "PASS" if all_ready else "BLOCKED"}, ensure_ascii=False, indent=2))
    return 0 if all_ready else 1


if __name__ == "__main__":
    raise SystemExit(main())
