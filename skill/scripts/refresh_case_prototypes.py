"""Refresh prototype freshness state without copying case assets."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
INDEX = ROOT / "skill/templates/case-prototype-index.json"
CACHE = ROOT / ".cache/case-prototype-state.json"
LOCK = ROOT / ".cache/case-prototype-state.lock"


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


@contextmanager
def exclusive_lock(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    for _ in range(120):
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            break
        except FileExistsError:
            time.sleep(0.05)
    else:
        raise RuntimeError(f"refresh lock is busy: {path}")
    try:
        yield
    finally:
        try:
            path.unlink()
        except FileNotFoundError:
            pass


def record_status(root: Path, entry: dict, blocked: bool) -> tuple[str, str | None]:
    """Return a conservative cache status for one indexed prototype."""
    case_root_value = entry.get("case_root")
    case_root = root / case_root_value if case_root_value else root / "examples/new_examplex" / entry.get("case_id", "")
    if not case_root.is_dir():
        return "invalid", f"missing case root: {case_root}"
    if blocked:
        return "blocked", entry.get("reason")
    record_path = entry.get("path")
    if not record_path:
        return "valid", None
    try:
        record = json.loads((root / record_path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        return "invalid", f"invalid prototype record: {record_path} ({exc})"
    if record.get("license_status") != "verified" or record.get("context_allowed") is not True:
        return "blocked", "license or context permission is not verified"
    for field in ("source_paths", "preview_paths", "recipe_paths", "object_map_paths"):
        for rel in record.get(field, []):
            candidate = (case_root / rel).resolve()
            if case_root.resolve() not in candidate.parents or not candidate.is_file():
                return "invalid", f"missing or escaping evidence: {rel}"
    return "valid", None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--cases-root", type=Path)
    parser.add_argument("--index", type=Path, default=INDEX)
    parser.add_argument("--cache", type=Path, default=CACHE)
    parser.add_argument("--lock", type=Path, default=LOCK)
    args = parser.parse_args()
    index = json.loads(args.index.read_text(encoding="utf-8"))
    cases_root = args.cases_root or (args.root / index.get("cases_root", "examples/new_examplex"))
    with exclusive_lock(args.lock):
        old_state = json.loads(args.cache.read_text(encoding="utf-8")) if args.cache.exists() else {}
        old = old_state.get("records", {})
        records = {}
        entries = [(prototype, False) for prototype in index.get("prototypes", [])]
        entries.extend((prototype, True) for prototype in index.get("blocked", []))
        for prototype, blocked in entries:
            case_id = prototype.get("case_id", "")
            case_root_value = prototype.get("case_root")
            case_root = args.root / case_root_value if case_root_value else cases_root / case_id
            case_root = Path(case_root)
            status, error = record_status(args.root, prototype, blocked)
            prototype_id = prototype.get("prototype_id", case_id)
            records[prototype.get("prototype_id", case_id)] = {
                "case_fingerprint": fingerprint(case_root),
                "status": status,
                "error": error,
                "blocked": blocked,
                "checked_at": datetime.now(timezone.utc).isoformat(),
                "previous_status": old.get(prototype_id, {}).get("status"),
            }
        atomic_write(args.cache, {"schema_version": 1, "cases_root": str(cases_root), "records": records})
    print(json.dumps({"updated": len(records), "cache": str(args.cache)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
