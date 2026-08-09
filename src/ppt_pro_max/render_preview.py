"""Render a .pptx to PNG previews for fast visual debugging of build scripts.

Solves the "改坐标 → 打开 PowerPoint → 肉眼检查 → 再改" loop by batch-rendering
every slide to PNG plus a single HTML contact sheet, so you see all pages at
once without manually opening PowerPoint.

Backends (auto-detected in order):
  1. Microsoft PowerPoint (COM)  — best fidelity, Windows only
  2. LibreOffice (soffice)       — cross-platform fallback

Usage:
    python -m ppt_pro_max.render_preview build_crispr.pptx [--out output/preview]
    python -m ppt_pro_max.render_preview build_crispr.pptx --open

Python API:
    from ppt_pro_max.render_preview import render_preview
    result = render_preview("build_crispr.pptx")
    # result == {"pngs": [Path, ...], "html": Path, "engine": "powerpoint"|"libreoffice"}
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


def _find_libreoffice() -> str | None:
    """Locate the soffice executable, or None if LibreOffice is not installed."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"/usr/bin/libreoffice",
        r"/usr/bin/soffice",
        r"/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def _powerpoint_available() -> bool:
    """True if Microsoft PowerPoint COM is importable on this machine."""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001 - any failure means no PowerPoint
        return False
    return True


def detect_engine() -> str:
    """Return the best available rendering engine: 'powerpoint' | 'libreoffice' | 'none'."""
    if _powerpoint_available():
        return "powerpoint"
    if _find_libreoffice():
        return "libreoffice"
    return "none"


def _render_powerpoint(pptx: Path, out_dir: Path, width: int, height: int) -> list[Path]:
    """Render every slide via PowerPoint COM. Returns sorted PNG paths."""
    import time

    import win32com.client

    pngs: list[Path] = []
    ppt: Any = None
    pres: Any = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = 1
        time.sleep(1.2)
        # Open read-only, no window, don't render fonts
        pres = ppt.Presentations.Open(
            str(pptx),
            ReadOnly=-1,
            Untitled=-1,
            WithWindow=-1,
        )
        count = pres.Slides.Count
        for i in range(1, count + 1):
            out_png = out_dir / f"slide{i}.png"
            pres.Slides.Item(i).Export(str(out_png), "PNG", width, height)
            pngs.append(out_png)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:  # noqa: BLE001, S110 - best-effort cleanup
                pass
        if ppt is not None:
            try:
                ppt.Quit()
            except Exception:  # noqa: BLE001, S110 - best-effort cleanup
                pass
    return sorted(pngs)


def _render_libreoffice(pptx: Path, out_dir: Path) -> list[Path]:
    """Render via LibreOffice: convert pptx→PDF, then each PDF page→PNG."""
    soffice = _find_libreoffice()
    if soffice is None:
        raise RuntimeError("LibreOffice not found")
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(pptx)],
            check=True,
            capture_output=True,
            timeout=300,
        )
        pdfs = list(tmp.glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("LibreOffice produced no PDF")
        pdf = pdfs[0]
        from pdf2image import convert_from_path  # type: ignore

        pages = convert_from_path(str(pdf), dpi=110)
        pngs: list[Path] = []
        for idx, page in enumerate(pages, start=1):
            out_png = out_dir / f"slide{idx}.png"
            page.save(out_png, "PNG")
            pngs.append(out_png)
    return sorted(pngs)


def _build_html(pngs: list[Path], pptx: Path, title: str = "Slide Preview") -> Path:
    """Build a single HTML contact sheet embedding every slide PNG."""
    html_dir = pngs[0].parent if pngs else Path(".")
    rel = [p.name for p in pngs]
    items = "\n".join(
        f'  <div class="slide"><img src="{name}" alt="{name}"></div>' for name in rel
    )
    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{title} — {pptx.name}</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 24px; background: #f5f5f5; }}
  h1 {{ font-size: 18px; color: #333; }}
  .slide {{ margin-bottom: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.25); }}
  .slide img {{ width: 100%; display: block; background: #fff; }}
</style>
</head>
<body>
<h1>{title} — {pptx.name} ({len(pngs)} slides)</h1>
{items}
</body>
</html>
"""
    html_path = html_dir / "index.html"
    html_path.write_text(doc, encoding="utf-8")
    return html_path


def render_preview(
    pptx_path: str,
    out_dir: str | None = None,
    width: int = 1280,
    height: int = 720,
    engine: str | None = None,
    title: str = "Slide Preview",
) -> dict[str, Any]:
    """Render every slide of a .pptx to PNG plus an HTML contact sheet.

    Args:
        pptx_path: Path to the .pptx file.
        out_dir: Output directory for PNGs + index.html. Defaults to
            <pptx_dir>/preview/<stem>/.
        width/height: Pixel size for PowerPoint export (LibreOffice ignores).
        engine: Force backend: 'powerpoint', 'libreoffice', or None=auto.
        title: Heading shown in the HTML preview page.

    Returns:
        {"pngs": [Path, ...], "html": Path, "engine": str}
    """
    src = Path(pptx_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"File not found: {pptx_path}")

    if out_dir is None:
        out_dir = str(src.parent / "preview" / src.stem)
    out = Path(out_dir)
    if not out.is_absolute():
        out = (src.parent / out).resolve()
    out.mkdir(parents=True, exist_ok=True)

    selected = engine or detect_engine()
    if selected == "powerpoint":
        pngs = _render_powerpoint(src, out, width, height)
    elif selected == "libreoffice":
        pngs = _render_libreoffice(src, out)
    else:
        raise RuntimeError(
            "No rendering engine available. Install PowerPoint (Windows) or "
            "LibreOffice, or run from a machine that has one."
        )

    if not pngs:
        raise RuntimeError("Rendering produced no slide images")

    html = _build_html(pngs, src, title=title)
    return {"pngs": pngs, "html": html, "engine": selected}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Render a .pptx to PNG previews")
    parser.add_argument("pptx", help="Path to the .pptx file")
    parser.add_argument("--out", default=None, help="Output directory (default: <pptx>/preview/<stem>)")
    parser.add_argument("--width", type=int, default=1280, help="Export width (PowerPoint only)")
    parser.add_argument("--height", type=int, default=720, help="Export height (PowerPoint only)")
    parser.add_argument("--engine", choices=["powerpoint", "libreoffice", "auto"], default="auto")
    parser.add_argument("--open", action="store_true", help="Open the HTML preview in browser")
    parser.add_argument("--title", default="Slide Preview", help="Preview page heading")
    args = parser.parse_args()

    result = render_preview(
        args.pptx,
        out_dir=args.out,
        width=args.width,
        height=args.height,
        engine=None if args.engine == "auto" else args.engine,
        title=args.title,
    )
    print(f"[render_preview] engine={result['engine']}")
    print(f"[render_preview] {len(result['pngs'])} slides -> {result['pngs'][0].parent}")
    print(f"[render_preview] HTML: {result['html']}")
    if args.open:
        import webbrowser

        webbrowser.open(str(result["html"]))
