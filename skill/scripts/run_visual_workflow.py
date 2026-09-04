"""Minimal persistent workflow state machine for visual-capability runs."""
from __future__ import annotations

import argparse
import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

STATES = {"INIT", "ROUTED", "DIRECTION_PENDING", "PLANNED", "P01_PENDING", "P01_PASS", "P01_REVISION", "BLOCKED", "REVIEW_PENDING", "PASS"}
TRANSITIONS = {
    "INIT": {"route": "ROUTED"},
    "ROUTED": {"directions": "DIRECTION_PENDING"},
    "DIRECTION_PENDING": {"confirm": "PLANNED", "block": "BLOCKED"},
    "PLANNED": {"p01": "P01_PENDING"},
    "P01_PENDING": {"pass": "P01_PASS", "revise": "P01_REVISION", "block": "BLOCKED"},
    "P01_REVISION": {"p01": "P01_PENDING"},
    "P01_PASS": {"review": "REVIEW_PENDING"},
    "REVIEW_PENDING": {"pass": "PASS", "revise": "PLANNED", "block": "BLOCKED"},
    "BLOCKED": {"resume": "ROUTED"},
}


def write_atomic(path: Path, value: dict) -> None:
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
    parser.add_argument("state_file", type=Path)
    parser.add_argument("action", choices=sorted({a for actions in TRANSITIONS.values() for a in actions}))
    args = parser.parse_args()
    now = datetime.now(timezone.utc).isoformat()
    state = json.loads(args.state_file.read_text(encoding="utf-8")) if args.state_file.exists() else {"status": "INIT", "attempt": 0, "history": []}
    current = state.get("status", "INIT")
    if current not in STATES or args.action not in TRANSITIONS.get(current, {}):
        print(f"ERROR: action {args.action!r} is invalid from {current!r}")
        return 1
    next_state = TRANSITIONS[current][args.action]
    state.update({"status": next_state, "attempt": state.get("attempt", 0) + 1, "updated_at": now})
    state.setdefault("history", []).append({"from": current, "action": args.action, "to": next_state, "at": now})
    write_atomic(args.state_file, state)
    print(json.dumps(state, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
