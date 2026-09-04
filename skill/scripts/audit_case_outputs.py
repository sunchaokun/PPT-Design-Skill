"""Audit case output completeness without claiming visual or license approval."""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


def slide_count(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(bool(re.fullmatch(r"ppt/slides/slide\d+\.xml", name)) for name in archive.namelist())


def audit_case(case_root: Path) -> dict[str, object]:
    pptx_files = sorted((case_root / "output").glob("*.pptx"))
    pdf_files = sorted(set((case_root / "output").glob("*.pdf")) | set((case_root / "rendered").glob("*.pdf")))
    png_files = sorted((case_root / "rendered").glob("slide*.png"))
    result: dict[str, object] = {
        "case_id": case_root.name,
        "pptx": [str(path.relative_to(case_root)) for path in pptx_files],
        "pdf": [str(path.relative_to(case_root)) for path in pdf_files],
        "rendered_png_count": len(png_files),
        "status": "BLOCKED",
        "blocking_reasons": [],
    }
    reasons = result["blocking_reasons"]
    assert isinstance(reasons, list)
    if len(pptx_files) != 1:
        reasons.append(f"expected exactly one output PPTX, found {len(pptx_files)}")
    if not pdf_files:
        reasons.append("missing output PDF")
    if len(pptx_files) == 1:
        try:
            pages = slide_count(pptx_files[0])
            result["pptx_slide_count"] = pages
            if not pdf_files:
                pass
            if len(png_files) != pages:
                reasons.append(f"expected {pages} rendered PNGs, found {len(png_files)}")
        except (OSError, zipfile.BadZipFile) as exc:
            reasons.append(f"invalid PPTX: {exc}")
    if not reasons:
        result["status"] = "render_ready"
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--cases-root", type=Path, default=None)
    args = parser.parse_args()
    cases_root = args.cases_root or args.root / "examples" / "new_examplex"
    cases = [path for path in sorted(cases_root.iterdir()) if path.is_dir() and not path.name.startswith(".")]
    results = [audit_case(path) for path in cases]
    print(json.dumps({"cases_root": str(cases_root), "cases": results}, ensure_ascii=False, indent=2))
    return 0 if all(item["status"] == "render_ready" for item in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
