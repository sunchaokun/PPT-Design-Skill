"""Check that visual-review claims point to the current case outputs."""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def pptx_pages(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(bool(re.fullmatch(r"ppt/slides/slide\d+\.xml", name)) for name in archive.namelist())


def audit_case(case_root: Path) -> dict[str, object]:
    review = case_root / "visual-review.md"
    pptx_files = sorted((case_root / "output").glob("*.pptx"))
    errors: list[str] = []
    if not review.is_file():
        return {"case_id": case_root.name, "status": "BLOCKED", "errors": ["missing visual-review.md"]}
    if len(pptx_files) != 1:
        errors.append(f"expected one output PPTX, found {len(pptx_files)}")
    if len(pptx_files) == 1:
        pages = pptx_pages(pptx_files[0])
        text = review.read_text(encoding="utf-8")
        page_claims = [int(value) for value in re.findall(r"(?:\b|/)(\d+) slides\b|(?:\b|/)(\d+) pages\b", text) for value in value if value]
        if page_claims and pages not in page_claims:
            errors.append(f"review page claim {page_claims} does not match PPTX page count {pages}")
        for rel in re.findall(r"`(rendered/[^`]*slide0?1\.png)`", text):
            if not (case_root / rel).is_file():
                errors.append(f"missing review render path: {rel}")
        result = {"case_id": case_root.name, "pptx_slide_count": pages, "status": "PASS" if not errors else "BLOCKED", "errors": errors}
        return result
    return {"case_id": case_root.name, "status": "BLOCKED", "errors": errors}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--cases-root", type=Path)
    args = parser.parse_args()
    cases_root = args.cases_root or args.root / "examples" / "new_examplex"
    results = [audit_case(path) for path in sorted(cases_root.iterdir()) if path.is_dir() and not path.name.startswith(".")]
    ok = bool(results) and all(item["status"] == "PASS" for item in results)
    print(json.dumps({"status": "PASS" if ok else "BLOCKED", "cases": results}, ensure_ascii=False, indent=2))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
