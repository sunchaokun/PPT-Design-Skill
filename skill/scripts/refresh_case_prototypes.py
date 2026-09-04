"""Refresh prototype freshness state without copying case assets."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
INDEX = ROOT / "skill/templates/case-prototype-index.json"
CACHE = ROOT / ".cache/case-prototype-state.json"


def fingerprint(path: Path) -> str:
    h = hashlib.sha256()
    if path.is_file():
        h.update(path.read_bytes())
    elif path.is_dir():
        for child in sorted(p for p in path.rglob("*") if p.is_file()):
            h.update(str(child.relative_to(path)).encode())
            h.update(str(child.stat().st_size).encode())
            h.update(str(child.stat().st_mtime_ns).encode())
    return h.hexdigest()


def atomic_write(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--cases-root", type=Path)
    args = parser.parse_args()
    index = json.loads(INDEX.read_text(encoding="utf-8"))
    cases_root = args.cases_root or (args.root / index.get("cases_root", "examples/new_examplex"))
    old = json.loads(CACHE.read_text(encoding="utf-8")) if CACHE.exists() else {}
    records = {}
    for prototype in index.get("prototypes", []):
        case_id = prototype.get("case_id", "")
        case_root = args.root / prototype.get("case_root", cases_root / case_id)
        records[prototype.get("prototype_id", case_id)] = {
            "case_fingerprint": fingerprint(case_root),
            "status": "valid" if case_root.exists() else "invalid",
            "checked_at": datetime.now(timezone.utc).isoformat(),
            "previous_status": old.get(prototype.get("prototype_id", case_id), {}).get("status"),
        }
    atomic_write(CACHE, {"schema_version": 1, "cases_root": str(cases_root), "records": records})
    print(json.dumps({"updated": len(records), "cache": str(CACHE)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
