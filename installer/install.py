#!/usr/bin/env python3
"""PPT Design Skill — One-click installer (refactored).

Source of truth: skill/ directory at project root.
Copies skill/ contents + root src/ (via junction) + root pyproject.toml +
root component_library/ to each platform's skill directory.

Usage:
    python installer/install.py                  # Auto-detect + install to global dirs
    python installer/install.py --project        # Also install to project-level dirs
    python installer/install.py --no-global      # Skip global, only --project
    python installer/install.py --platform claude  # Install for specific platform
    python installer/install.py --check          # Check current installation status
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

for _stream_name in ("stdout", "stderr"):
    _stream = getattr(sys, _stream_name)
    if _stream and hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

SKILL_NAME = "ppt-design-skill"

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

EXCLUDE_NAMES = {
    "__pycache__",
    ".pyc",
    ".egg-info",
    ".env",
}


def _get_version(project_root: Path) -> str:
    content = (project_root / "pyproject.toml").read_text(encoding="utf-8")
    match = re.search(r'version\s*=\s*"([^"]+)"', content)
    return match.group(1) if match else "0.0.0"


def _get_project_root() -> Path:
    start = Path(__file__).resolve().parent.parent
    for candidate in [start] + list(start.parents):
        if (candidate / "src" / "ppt_pro_max").exists() and (candidate / "skill" / "SKILL.md").exists():
            return candidate
    return start


def _get_skill_dir(project_root: Path) -> Path:
    return project_root / "skill"


def get_skill_sources(project_root: Path) -> dict[str, Path]:
    skill_dir = _get_skill_dir(project_root)
    sources: dict[str, Path] = {}

    for name in ("SKILL.md", "AGENTS.md", ".env.example"):
        if (skill_dir / name).exists():
            sources[name] = skill_dir / name

    for dir_name in ("scripts", "data"):
        skill_path = skill_dir / dir_name
        if skill_path.exists() and not skill_path.is_symlink():
            real = os.path.realpath(skill_path)
            if real != str(skill_path):
                sources[dir_name] = Path(real)
            else:
                sources[dir_name] = skill_path

    skill_src = skill_dir / "src"
    if skill_src.exists():
        real = os.path.realpath(skill_src)
        sources["src"] = Path(real)
    else:
        root_src = project_root / "src"
        if root_src.exists():
            sources["src"] = root_src

    if (project_root / "pyproject.toml").exists():
        sources["pyproject.toml"] = project_root / "pyproject.toml"

    return sources


def _ignore_patterns(directory: str, contents: list[str]) -> list[str]:
    ignored: list[str] = []
    for name in contents:
        if name in EXCLUDE_NAMES:
            ignored.append(name)
        elif name.endswith(".pyc"):
            ignored.append(name)
        elif name.endswith(".egg-info"):
            ignored.append(name)
    return ignored


def _copy_skill_sources(sources: dict[str, Path], dest_dir: Path) -> None:
    dest_dir.mkdir(parents=True, exist_ok=True)

    for name, src_path in sources.items():
        dst = dest_dir / name
        if src_path.is_dir():
            if dst.exists():
                if dst.resolve() == src_path.resolve():
                    continue
                shutil.rmtree(dst, ignore_errors=True)
            shutil.copytree(src_path, dst, ignore=_ignore_patterns)
        else:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src_path, dst)


def _copy_component_library(project_root: Path, dest_dir: Path) -> None:
    for asset_name in ("component_library",):
        src = project_root / asset_name
        if not src.exists():
            continue
        dst = dest_dir / asset_name
        if dst.exists():
            if dst.resolve() == src.resolve():
                continue
            shutil.rmtree(dst, ignore_errors=True)
        shutil.copytree(src, dst, ignore=_ignore_patterns)
        print(f"    + {asset_name}/ -> {dst}")


def copy_skill(project_root: Path, dest_dir: Path, force: bool = False) -> bool:
    if dest_dir.exists():
        if not force:
            return False
        source_skill = _get_skill_dir(project_root)
        if dest_dir.resolve() == source_skill.resolve():
            return True
        shutil.rmtree(dest_dir)

    sources = get_skill_sources(project_root)
    _copy_skill_sources(sources, dest_dir)
    return True


def _read_skill_version(skill_dir: Path) -> str:
    skill_md = skill_dir / "SKILL.md"
    if not skill_md.exists():
        return ""
    try:
        for line in skill_md.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("version:"):
                return line.split(":", 1)[1].strip().strip('"')
    except Exception:
        pass
    return ""


def _skill_status(skill_dir: Path) -> str:
    skill_md = skill_dir / "SKILL.md"
    has_src = (skill_dir / "src" / "ppt_pro_max").exists()
    has_data = (skill_dir / "data").exists()
    if skill_md.exists() and has_src:
        ver = _read_skill_version(skill_dir)
        parts: list[str] = []
        if not has_data:
            parts.append("no data")
        suffix = f" [{', '.join(parts)}]" if parts else ""
        return f"INSTALLED v{ver}{suffix}" if ver else f"INSTALLED{suffix}"
    elif skill_md.exists():
        return "PARTIAL (no src/)"
    return "NOT INSTALLED"


def install_to_dir(project_root: Path, dest_dir: Path, label: str,
                   force: bool = False) -> str | None:
    source_skill = _get_skill_dir(project_root)
    version = _get_version(project_root)

    if dest_dir.resolve() == source_skill.resolve():
        return None
    if dest_dir.exists() and not force:
        ver = _read_skill_version(dest_dir)
        print(f"  [SKIP] {label} — already installed")
        if ver:
            print(f"         Version: {ver}, Latest: {version}")
        print("         Use --force to overwrite")
        return None

    copy_skill(project_root, dest_dir, force=force)
    print(f"  [OK] {label} -> {dest_dir}")

    _copy_component_library(project_root, dest_dir)

    return str(dest_dir)


def install_python_package(skill_dir: Path, project_root: Path | None = None) -> bool:
    if not (skill_dir / "pyproject.toml").exists():
        print(f"  [SKIP] No pyproject.toml in {skill_dir}")
        return False
    install_dir = project_root if project_root and (project_root / "pyproject.toml").exists() and (project_root / "src" / "ppt_pro_max").exists() else skill_dir
    print(f"  Installing Python package from {install_dir}...")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", str(install_dir)],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode == 0:
            success = False
            for line in result.stdout.splitlines():
                if "Successfully installed" in line and "ppt-design-skill" in line:
                    success = True
                    break
            if not success:
                for line in result.stdout.splitlines():
                    if "Requirement already satisfied" in line and "ppt-design-skill" in line:
                        success = True
                        break
            if success:
                print("  [OK] ppt-design-skill installed to site-packages")
            else:
                print("  [WARN] pip install succeeded but package name unexpected")
                print(f"  {result.stdout[:300]}")
            return True
        else:
            print(f"  [WARN] pip install failed (exit {result.returncode})")
            print(f"  {result.stderr[:200]}")
            print(f"  You can install manually: pip install {install_dir}")
            return False
    except Exception as e:
        print(f"  [WARN] pip install failed: {e}")
        print(f"  You can install manually: pip install {install_dir}")
        return False


def install_ui_ux_pro_max(target_dir: Path, force: bool = False) -> bool:
    print("\n  Checking bundled design knowledge database...")
    # Since v2, the design knowledge database is BUNDLED inside ppt_pro_max
    # (adapters/design_search.py + data/*.csv). No external ui-ux-pro-max
    # skill / uipro CLI is required — is_available() reflects the bundled DB.
    # Ensure we inspect THIS project's src (not a stale site-packages copy).
    _root = Path(__file__).resolve().parent.parent
    _src = _root / "src"
    if _src.is_dir() and str(_src) not in sys.path:
        sys.path.insert(0, str(_src))
    from ppt_pro_max.adapters.ui_ux_adapter import is_available

    if is_available():
        print("  [OK] design knowledge search is bundled (no external install needed)")
        return True

    print("  [WARN] Bundled design database is missing — ui-ux-pro-max not available.")
    print("  [WARN] Reinstall the python package so data/*.csv is included:")
    print("          pip install -e .")
    return False


# ── Render dependencies (LibreOffice + poppler) ────────────────────────────
def _find_soffice_bin() -> str | None:
    """Locate LibreOffice's soffice.bin (the real worker; soffice.exe launcher
    can hang headless)."""
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.bin",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.bin",
        r"/usr/lib/libreoffice/program/soffice.bin",
        r"/usr/bin/soffice",
        r"/usr/bin/libreoffice",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    for name in ("soffice.bin", "soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return None


def _find_pdftoppm_exe() -> str | None:
    """Locate a real pdftoppm.exe (skip .CMD/.bat shims)."""
    # user-local codex-runtimes native poppler
    home = os.path.expanduser("~")
    for root in (os.path.join(home, ".cache", "codex-runtimes"),):
        if os.path.isdir(root):
            for base, _dirs, files in os.walk(root):
                if "pdftoppm.exe" in files:
                    return os.path.join(base, "pdftoppm.exe")
    # direct .exe on PATH
    found = shutil.which("pdftoppm")
    if found and found.lower().endswith(".exe"):
        return found
    return None


def detect_render_deps() -> dict[str, bool]:
    """Check which render dependencies are present.

    Returns:
        {"libreoffice": bool, "poppler": bool}
    """
    return {
        "libreoffice": _find_soffice_bin() is not None,
        "poppler": _find_pdftoppm_exe() is not None,
    }


def _winget_install(package_id: str, label: str) -> bool:
    """Install a package via winget (silent). Returns True on success."""
    winget = shutil.which("winget")
    if not winget:
        print(f"  [WARN] winget not found — cannot auto-install {label}.")
        return False
    print(f"  Installing {label} via winget (this may take a few minutes)...")
    try:
        result = subprocess.run(
            [winget, "install", package_id,
             "--accept-source-agreements", "--accept-package-agreements",
             "--silent"],
            capture_output=True, text=True, timeout=1200,
        )
        if result.returncode == 0:
            print(f"  [OK] {label} installed")
            return True
        print(f"  [WARN] {label} install failed (exit {result.returncode})")
        if result.stderr.strip():
            print(f"          {result.stderr.strip()[:300]}")
        return False
    except subprocess.TimeoutExpired:
        print(f"  [WARN] {label} install timed out (1200s)")
        return False
    except Exception as e:  # noqa: BLE001
        print(f"  [WARN] {label} install failed: {e}")
        return False


def ensure_render_deps() -> dict[str, str]:
    """Detect render deps and auto-install any that are missing.

    Uses winget to install LibreOffice + Poppler so the user is ready to
    render previews (Codex sandbox / headless) right after installation.

    Returns:
        {"libreoffice": "ok"|"installed"|"failed"|"skipped",
         "poppler":    "ok"|"installed"|"failed"|"skipped"}
    """
    print("\n  Checking render dependencies (LibreOffice + poppler)...")
    status = {"libreoffice": "ok", "poppler": "ok"}
    deps = detect_render_deps()

    if not deps["libreoffice"]:
        if _winget_install("TheDocumentFoundation.LibreOffice", "LibreOffice"):
            status["libreoffice"] = "installed"
        else:
            status["libreoffice"] = "failed"
    else:
        print("  [OK] LibreOffice found")

    if not deps["poppler"]:
        if _winget_install("oschwartz10612.Poppler", "Poppler (pdftoppm)"):
            status["poppler"] = "installed"
        else:
            status["poppler"] = "failed"
    else:
        print("  [OK] poppler (pdftoppm) found")

    # Re-detect after install to confirm (installed-but-unfindable → failed)
    for key, locator in (("libreoffice", _find_soffice_bin),
                         ("poppler", _find_pdftoppm_exe)):
        if status[key] == "installed":
            if not locator():
                status[key] = "failed"
                print(f"  [WARN] {key} installed but not found after install — check PATH")
    return status


def check_installation(target_dir: Path, project_root: Path) -> None:
    version = _get_version(project_root)
    from installer.detect import PLATFORMS, detect_project_platforms

    print(f"\nPPT Design Skill v{version} — Installation Check")
    print(f"Project dir : {target_dir}")
    print(f"Home dir    : {Path.home()}")
    print()

    print("Platform detections (project-level):")
    detected = detect_project_platforms(target_dir)
    if detected:
        for p in detected:
            info = PLATFORMS.get(p, {})
            skill_dir = target_dir / info["project_dir"] / "skills" / SKILL_NAME
            status = _skill_status(skill_dir)
            print(f"  {info.get('desc', p):25s} [{status}]")
    else:
        print("  No AI coding platforms detected in current directory")

    print("\nGlobal installations:")
    for p, info in PLATFORMS.items():
        global_dir = Path.home() / info["global_path"] / SKILL_NAME
        status = _skill_status(global_dir)
        if status != "NOT INSTALLED":
            print(f"  {info['desc']:25s} [{status}]  ({global_dir})")

    print("\nPython dependencies:")
    import importlib.util
    for pkg, label in [("pptx", "python-pptx"), ("PIL", "Pillow"), ("ppt_pro_max", "ppt_pro_max")]:
        spec = importlib.util.find_spec(pkg)
        print(f"  {label:20s} [{'OK' if spec else 'MISSING'}]")

    print()


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="ppt-design-install",
        description="PPT Design Skill — One-click installer",
    )
    parser.add_argument("--platform", "-p", help="Specific platform (claude, opencode, codex, cursor, all)")
    parser.add_argument("--target", "-t", help="Target project directory (default: current)")
    parser.add_argument("--no-global", dest="no_global", action="store_true",
                        help="Skip global installation (global is installed by default)")
    parser.add_argument("--project", dest="project_install", action="store_true",
                        help="Also install to project-level .claude/.opencode etc dirs")
    parser.add_argument("--force", "-f", action="store_true", help="Overwrite existing installation")
    parser.add_argument("--check", "-c", action="store_true", help="Check installation status only")
    parser.add_argument("--no-pip", action="store_true", help="Skip pip install")
    parser.add_argument("--no-render-deps", action="store_true",
                        help="Skip auto-installing LibreOffice + poppler render deps")
    args = parser.parse_args()

    from installer.detect import PLATFORMS, detect_project_platforms, detect_global_platforms

    target_dir = Path(args.target) if args.target else Path.cwd()
    project_root = _get_project_root()
    version = _get_version(project_root)

    if args.check:
        check_installation(target_dir, project_root)
        print()
        deps = detect_render_deps()
        print("Render dependencies:")
        for name, ok in deps.items():
            mark = "OK" if ok else "MISSING"
            print(f"  {name:12s} [{mark}]")
        print()
        return

    print(f"\n{'='*60}")
    print(f"  PPT Design Skill v{version} — Installer")
    print("  3 modes · 40,000+ style combos · 28 design quality upgrades · 6 image engines")
    print(f"{'='*60}\n")

    skill_dir = _get_skill_dir(project_root)
    print(f"  Source: {skill_dir}\n")

    all_installed: list[str] = []

    # --- Global install (always, unless --no-global) ---
    if not args.no_global:
        global_platforms = detect_global_platforms()
        if args.platform and args.platform != "all":
            global_platforms = [p for p in global_platforms if p == args.platform]
        if global_platforms:
            print(f"Installing to global: {', '.join(PLATFORMS[p]['desc'] for p in global_platforms)}")
            print()
            for platform in global_platforms:
                info = PLATFORMS[platform]
                dest_dir = Path.home() / info["global_path"] / SKILL_NAME
                result = install_to_dir(project_root, dest_dir, f"{info['desc']} (global)", force=args.force)
                if result:
                    all_installed.append(result)
        else:
            print("No matching global AI platforms detected. Skipping global install.")
        print()

    # --- Project-level install (only if --project specified) ---
    if args.project_install:
        if args.platform == "all":
            platforms_to_install = list(PLATFORMS.keys())
        elif args.platform:
            platforms_to_install = [args.platform]
        else:
            detected = detect_project_platforms(target_dir)
            if detected:
                platforms_to_install = detected
                print(f"Auto-detected project platforms: {', '.join(PLATFORMS[p]['desc'] for p in detected)}")
            else:
                platforms_to_install = ["claude", "opencode", "codex", "cursor"]
                print(f"No project platforms detected. Installing for: {', '.join(PLATFORMS[p]['desc'] for p in platforms_to_install)}")

        print(f"\nInstalling to project: {', '.join(PLATFORMS.get(p, {}).get('desc', p) for p in platforms_to_install)}")
        print()

        for platform in platforms_to_install:
            info = PLATFORMS.get(platform)
            if not info:
                print(f"  [SKIP] Unknown platform: {platform}")
                continue
            dest_dir = target_dir / info["project_dir"] / "skills" / SKILL_NAME
            result = install_to_dir(project_root, dest_dir, f"{info['desc']} (project)",
                                    force=args.force)
            if result:
                all_installed.append(result)

        agents_src = skill_dir / "AGENTS.md"
        if agents_src.exists():
            root_agents = target_dir / "AGENTS.md"
            if not root_agents.exists() or args.force:
                try:
                    shutil.copy2(agents_src, root_agents)
                except PermissionError:
                    print(f"  [WARN] Cannot overwrite {root_agents} (file locked)")

    # --- pip install ---
    if not args.no_pip:
        print()
        pip_dirs = list(set(all_installed))
        if pip_dirs:
            for d in pip_dirs:
                install_python_package(Path(d), project_root=project_root)
        else:
            for platform in (global_platforms if not args.no_global else []):
                info = PLATFORMS[platform]
                skill_dest = Path.home() / info["global_path"] / SKILL_NAME
                if skill_dest.exists():
                    install_python_package(skill_dest, project_root=project_root)
                    break
            else:
                install_python_package(project_root, project_root=project_root)

    # --- bundled design knowledge database ---
    install_ui_ux_pro_max(target_dir, force=args.force)

    # --- render dependencies (LibreOffice + poppler) auto-install ---
    if not args.no_render_deps:
        ensure_render_deps()

    # --- Summary ---
    print(f"\n{'='*60}")
    if all_installed:
        print("  Installed successfully!")
        print("\n  Installed to:")
        for path in all_installed:
            print(f"    {path}")
    else:
        print("  No new installations (use --force to overwrite)")

    print("\n  Next steps:")
    print("    1. Restart your AI coding assistant")
    print("    2. Try: 'Generate a PPT for AI startup investor pitch'")
    print("    3. Or: python -m ppt_pro_max 'AI pitch' --style 'dark cyberpunk'")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
