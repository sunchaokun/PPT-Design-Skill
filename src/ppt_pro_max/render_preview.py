"""Render a .pptx to PNG previews for fast visual debugging of build scripts.

Solves the "改坐标 → 打开 PowerPoint → 肉眼检查 → 再改" loop by batch-rendering
every slide to PNG plus a single HTML contact sheet, so you see all pages at
once without manually opening PowerPoint.

Backends (auto-detected in order, graceful fallback if one fails):
  1. Microsoft PowerPoint (COM)  — best fidelity, Windows interactive session
  2. LibreOffice headless        — soffice.bin, works in sandbox/CI/no desktop
                                     (e.g. Codex sandbox where COM fails 1312)

If PowerPoint COM fails (e.g. WinError 1312 in a background/CI session), the
preview automatically falls back to LibreOffice so the caller can still judge
layout from real renders.

Usage:
    python -m ppt_pro_max.render_preview build.pptx [--out output/preview]
    python -m ppt_pro_max.render_preview build.pptx --open
    python -m ppt_pro_max.render_preview build.pptx --engine libreoffice

Python API:
    from ppt_pro_max.render_preview import render_preview
    result = render_preview("build.pptx")
    # result == {"pngs": [Path, ...], "html": Path,
    #            "engine": "powerpoint"|"libreoffice", "warnings": [...]}
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


def _find_libreoffice() -> str | None:
    """Locate the soffice executable, preferring soffice.bin.

    On Windows, `soffice.exe` is a launcher that can hang in headless / CI /
    sandbox sessions; `soffice.bin` is the real worker and works reliably.
    """
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.bin",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.bin",
        r"/usr/lib/libreoffice/program/soffice.bin",
        r"/usr/bin/libreoffice",
        r"/usr/bin/soffice",
        r"/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    for name in ("soffice.bin", "soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
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


def describe_com_error(exc: Exception) -> str:
    """Human-readable description of a PowerPoint COM failure.

    Maps common HRESULT codes (e.g. 0x80070520 = no logon session) to
    actionable hints so the caller knows WHY the backend failed.
    """
    hresult = getattr(exc, "hresult", None)
    if not hresult:
        return repr(exc)
    try:
        import winerror

        code = winerror.HRESULT_CODE(hresult)
    except Exception:  # noqa: BLE001
        code = None
    detail = f"COM error 0x{hresult:08X}"
    if code is not None:
        detail += f" (WinError {code})"
    if code == 1312:  # ERROR_NO_SUCH_LOGON_SESSION
        detail += (
            " — no interactive desktop. This happens in a background/CI/sandbox "
            "session (e.g. Codex). Use engine='libreoffice' or run interactively."
        )
    elif code in (5, 1314):
        detail += " — access denied. Run under the Windows-signed-in account."
    return detail


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


def _find_pdftoppm() -> str | None:
    """Locate the REAL pdftoppm.exe (poppler), or None if not installed.

    Skips `.CMD`/`.bat` shims: on Codex sandbox the PATH may expose an
    override wrapper (pdftoppm.cmd) whose chained exe can't find its DLLs.
    We return the actual .exe path so its directory can be added to PATH
    (poppler needs its sibling DLLs on the DLL search path).

    Search order:
      1. Well-known codex-runtimes native poppler (Library\\bin\\pdftoppm.exe)
      2. A direct pdftoppm.exe on PATH (skip .CMD shims)
      3. winget per-user poppler install
      4. Program Files poppler
    """
    exe_names = ("pdftoppm.exe", "pdftoppm")  # prefer .exe
    home = os.path.expanduser("~")

    # 1) codex-runtimes native poppler (the real install with DLLs)
    codex_roots = [os.path.join(home, ".cache", "codex-runtimes")]
    for root in codex_roots:
        if os.path.isdir(root):
            for base, _dirs, files in os.walk(root):
                if "pdftoppm.exe" in files:
                    return os.path.join(base, "pdftoppm.exe")

    # 2) direct .exe on PATH (skip .CMD/.bat shims)
    for name in exe_names:
        found = shutil.which(name)
        if found and found.lower().endswith(".exe"):
            return found

    # 3) winget per-user + 4) Program Files
    candidates = [
        os.path.join(home, "AppData", "Local", "Microsoft", "WinGet", "Packages"),
        r"C:\Program Files\poppler\Library\bin\pdftoppm.exe",
        r"C:\Program Files (x86)\poppler\Library\bin\pdftoppm.exe",
    ]
    for root in candidates:
        if os.path.isfile(root):
            return root
        if os.path.isdir(root):
            for base, _dirs, files in os.walk(root):
                if "pdftoppm.exe" in files:
                    return os.path.join(base, "pdftoppm.exe")
    return None


def _run_pdftoppm(pdftoppm: str, pdf: Path, out_dir: Path) -> list[Path]:
    """Run pdftoppm with the exe's directory added to PATH (DLL resolution).

    poppler's pdftoppm.exe needs sibling DLLs (poppler-*.dll etc.) on the DLL
    search path. In a sandbox the chained wrapper may run without those dirs,
    producing a bare `exit status 3`. Adding the exe dir to PATH fixes it.
    """
    import copy

    env = copy.deepcopy(os.environ)
    exe_dir = os.path.dirname(os.path.abspath(pdftoppm))
    env["PATH"] = exe_dir + os.pathsep + env.get("PATH", "")
    base = out_dir / "lo_slide"
    subprocess.run(
        [pdftoppm, "-png", "-r", "110", str(pdf), str(base)],
        check=True,
        capture_output=True,
        timeout=300,
        env=env,
    )
    return sorted(out_dir.glob("lo_slide-*.png"))


def _pdf_to_pngs(pdf: Path, out_dir: Path) -> list[Path]:
    """Convert a PDF to one PNG per page, named slide1..N (same as PowerPoint).

    Tries (in order):
      1. pdftoppm (poppler)  — best quality/control, fast
      2. pdf2image           — Python wrapper over poppler

    The engine writes temp files (lo_slide-*) then renames to slide{i}.png so
    the output dir is identical regardless of engine.
    """
    from PIL import Image as PILImage

    tmp_pngs: list[Path] = []

    # Path 1: pdftoppm (with DLL-path injection + pre-flight diagnostics)
    pdftoppm = _find_pdftoppm()
    if pdftoppm:
        if not os.path.isfile(pdftoppm):
            raise RuntimeError(
                f"pdftoppm resolved to a non-existent file: {pdftoppm}. "
                "Install poppler (winget install oschwartz10612.Poppler)."
            )
        try:
            tmp_pngs = _run_pdftoppm(pdftoppm, pdf, out_dir)
        except subprocess.CalledProcessError as e:
            raise RuntimeError(
                f"pdftoppm failed (exit {e.returncode}). File: {pdftoppm}\n"
                f"stderr: {(e.stderr or '').decode(errors='replace')[:300]}\n"
                "This usually means poppler DLLs are not resolvable. The exe "
                "directory is auto-added to PATH; verify poppler is installed."
            ) from e

    if not tmp_pngs:
        # Path 2: pdf2image
        from pdf2image import convert_from_path  # type: ignore

        pages = convert_from_path(str(pdf), dpi=110)
        for idx, page in enumerate(pages, start=1):
            out_png = out_dir / f"lo_slide{idx}.png"
            PILImage.frombytes(page.mode, page.size, page.tobytes()).save(out_png, "PNG")
            tmp_pngs.append(out_png)

    # Rename to fixed slide{i}.png
    final: list[Path] = []
    for idx, tmp in enumerate(sorted(tmp_pngs), start=1):
        final_path = out_dir / f"slide{idx}.png"
        if tmp != final_path:
            final_path.write_bytes(tmp.read_bytes())
            tmp.unlink(missing_ok=True)
        final.append(final_path)
    return final


def _render_libreoffice(pptx: Path, out_dir: Path, width: int = 1280, height: int = 720) -> list[Path]:
    """Render via LibreOffice headless: pptx→PDF (soffice.bin), then PDF→PNG.

    Uses `soffice.bin` directly (the `soffice.exe` launcher can hang in
    headless/sandbox sessions) with a FIXED user profile and --norestore.
    width/height are accepted for interface parity with the PowerPoint engine
    but LibreOffice output resolution is controlled by pdftoppm's -r flag.

    LibreOffice quirk: a brand-new profile fails on FIRST launch (RC 81) but
    initializes enough that a retry with the same profile succeeds. We use a
    stable profile dir and retry once on failure.
    """
    soffice = _find_libreoffice()
    if soffice is None:
        raise RuntimeError("LibreOffice not found")
    import tempfile

    _kill_libreoffice()

    profile = Path(tempfile.gettempdir()) / "ppt_lo_profile"
    profile.mkdir(parents=True, exist_ok=True)
    prof_arg = "-env:UserInstallation=file:///" + str(profile).replace("\\", "/")

    def _convert() -> list[Path]:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            cmd = [
                soffice, prof_arg,
                "--headless", "--norestore",
                "--convert-to", "pdf",
                "--outdir", str(tmp),
                str(pptx),
            ]
            subprocess.run(cmd, check=True, capture_output=True, timeout=300)
            pdfs = list(tmp.glob("*.pdf"))
            if not pdfs:
                raise RuntimeError("LibreOffice produced no PDF")
            return _pdf_to_pngs(pdfs[0], out_dir)

    try:
        return _convert()
    except subprocess.CalledProcessError:
        # First-run profile init often fails once; retry with initialized profile.
        return _convert()


def _kill_libreoffice() -> None:
    """Terminate stray soffice processes (Windows) to avoid profile lock / RC 81."""
    if sys.platform != "win32":
        return
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "soffice.bin", "/T"],
            capture_output=True, timeout=30,
        )
    except Exception:  # noqa: BLE001, S110 - best-effort cleanup
        pass


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

    Multi-tier fallback so a preview is ALWAYS produced:
      1. 'powerpoint'  — PowerPoint COM (Windows, interactive session)
      2. 'libreoffice' — soffice.bin headless (sandbox/CI/no-desktop safe)

    If the default engine fails, the next tier is tried automatically.

    Args:
        pptx_path: Path to the .pptx file.
        out_dir: Output directory for PNGs + index.html. Defaults to
            <pptx_dir>/preview/<stem>/.
        width/height: Pixel size for PowerPoint export.
        engine: Force backend: 'powerpoint' | 'libreoffice', or None=auto
            (tries tiers in order, falls back gracefully).
        title: Heading shown in the HTML preview page.

    Returns:
        {"pngs": [Path, ...], "html": Path, "engine": str, "warnings": [str]}
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

    warnings: list[str] = []
    pngs: list[Path] | None = None
    used_engine: str | None = None

    # Clear stale PNGs from previous renders so the output dir always holds
    # exactly one fresh set (fixed names slide1..N). Prevents the LLM from
    # accumulating / re-reading old images across runs.
    for stale in out.glob("*.png"):
        try:
            stale.unlink()
        except OSError:  # noqa: S110 - ignore transient lock
            pass

    if engine:
        order = [engine]
    else:
        order = ["powerpoint", "libreoffice"]

    for eng in order:
        try:
            if eng == "powerpoint":
                pngs = _render_powerpoint(src, out, width, height)
            elif eng == "libreoffice":
                pngs = _render_libreoffice(src, out, width, height)
            else:
                raise RuntimeError(f"Unknown engine: {eng}")
            used_engine = eng
            break
        except Exception as e:  # noqa: BLE001 - any engine failure falls back
            warnings.append(f"engine '{eng}' failed: {e}")
            pngs = None

    if not pngs or not used_engine:
        raise RuntimeError(
            "All rendering engines failed.\n  " + "\n  ".join(warnings)
        )

    html = _build_html(pngs, src, title=title)
    return {"pngs": pngs, "html": html, "engine": used_engine, "warnings": warnings}


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
    for w in result.get("warnings", []):
        print(f"[render_preview] warn: {w}")
    print(f"[render_preview] {len(result['pngs'])} slides -> {result['pngs'][0].parent}")
    print(f"[render_preview] HTML: {result['html']}")
    if args.open:
        import webbrowser

        webbrowser.open(str(result["html"]))
