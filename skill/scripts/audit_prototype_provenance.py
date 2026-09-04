"""Audit prototype source availability and hashes without overclaiming provenance."""
from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
INDEX = ROOT / "skill/templates/case-prototype-index.json"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_content_sha256(path: Path) -> str:
    """Hash ZIP entry names and bytes, excluding container metadata."""
    digest = hashlib.sha256()
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            name_bytes = name.encode("utf-8")
            payload = archive.read(name)
            digest.update(len(name_bytes).to_bytes(8, "big"))
            digest.update(name_bytes)
            digest.update(len(payload).to_bytes(8, "big"))
            digest.update(payload)
    return digest.hexdigest()


def audit(root: Path, index_path: Path) -> tuple[list[dict], bool]:
    index = json.loads(index_path.read_text(encoding="utf-8"))
    results: list[dict] = []
    complete = True
    entries = [(item, False) for item in index.get("prototypes", [])]
    entries.extend((item, True) for item in index.get("blocked", []))
    for entry, blocked in entries:
        record_path = root / entry["path"]
        record = json.loads(record_path.read_text(encoding="utf-8"))
        case_root = root / record["case_root"]
        files = {"build": case_root / "build.py"}
        for field in ("source_paths", "recipe_paths", "object_map_paths", "preview_paths"):
            for rel in record.get(field, []):
                files[f"{field}:{rel}"] = case_root / rel
        missing = [name for name, path in files.items() if not path.is_file()]
        hashes = {name: sha256(path) for name, path in files.items() if path.is_file()}
        content_hashes = {
            name: canonical_content_sha256(path)
            for name, path in files.items()
            if path.is_file() and path.suffix.lower() == ".pptx"
        }
        declared = record.get("extensions", {}).get("provenance", {})
        verified = (
            not missing
            and declared.get("status") == "verified"
            and declared.get("method") == "source_authoritative_rebuild"
            and declared.get("rebuild_count") == 2
        )
        result = {
            "prototype_id": record["prototype_id"],
            "case_id": record["case_id"],
            "blocked": blocked,
            "status": "VERIFIED" if verified else ("UNVERIFIED" if not missing else "INVALID"),
            "missing": missing,
            "hashes": hashes,
            "canonical_content_hashes": content_hashes,
        }
        complete = complete and verified
        results.append(result)
    return results, complete


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--index", type=Path, default=INDEX)
    args = parser.parse_args()
    try:
        results, complete = audit(args.root, args.index)
    except (OSError, json.JSONDecodeError, KeyError) as exc:
        print(f"ERROR: {exc}")
        return 1
    print(json.dumps({"status": "VERIFIED" if complete else "UNVERIFIED", "records": results}, ensure_ascii=False, indent=2))
    return 0 if complete else 1


if __name__ == "__main__":
    raise SystemExit(main())
