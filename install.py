#!/usr/bin/env python3
"""PPT Design Skill — One-click installer.

Auto-detects AI coding platform, installs skill files + dependencies.

Usage:
    python install.py                  # Auto-detect + install to project dir
    python install.py --global         # Install to global dirs for all platforms
    python install.py --platform claude  # Install for specific platform
    python install.py --platform all    # Install for all platforms
    python install.py --check          # Check current installation status
"""

from __future__ import annotations

import argparse
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

VERSION = "0.7.0"
SKILL_NAME = "ppt-design-skill"

# Source of truth: .opencode/skills/ppt-design-skill/
# This directory is the single canonical skill package.
# install.py copies it as-is to destination, only excluding build artifacts.

EXCLUDE_NAMES = {
    "__pycache__",
    ".pyc",
    ".egg-info",
    ".env",
}

PLATFORMS: dict[str, dict] = {
    "claude": {
        "project_dir": ".claude",
        "global_path": ".claude/skills",
        "desc": "Claude Code",
    },
    "opencode": {
        "project_dir": ".opencode",
        "global_path": ".config/opencode/skills",
        "desc": "OpenCode",
    },
    "codex": {
        "project_dir": ".codex",
        "global_path": ".agents/skills",
        "desc": "Codex / OpenClaw",
    },
    "cursor": {
        "project_dir": ".cursor",
        "global_path": ".agents/plugins",
        "desc": "Cursor",
    },
    "windsurf": {
        "project_dir": ".windsurf",
        "global_path": ".windsurf/skills",
        "desc": "Windsurf",
    },
    "roocode": {
        "project_dir": ".roo",
        "global_path": ".roo/skills",
        "desc": "RooCode",
    },
    "gemini": {
        "project_dir": ".gemini",
        "global_path": ".gemini/skills",
        "desc": "Gemini CLI",
    },
    "trae": {
        "project_dir": ".trae",
        "global_path": ".trae/skills",
        "desc": "Trae",
    },
    "continue": {
        "project_dir": ".continue",
        "global_path": ".continue/skills",
        "desc": "Continue",
    },
    "droid": {
        "project_dir": ".factory",
        "global_path": ".factory/skills",
        "desc": "Droid (Factory)",
    },
    "kilocode": {
        "project_dir": ".kilocode",
        "global_path": ".kilocode/skills",
        "desc": "KiloCode",
    },
    "augment": {
        "project_dir": ".augment",
        "global_path": ".augment/skills",
        "desc": "Augment",
    },
    "copilot": {
        "project_dir": ".github",
        "global_path": ".github/skills",
        "desc": "GitHub Copilot",
    },
}


def get_source_skill_dir() -> Path:
    script_dir = Path(__file__).resolve().parent
    candidate = script_dir / ".opencode" / "skills" / SKILL_NAME
    if candidate.exists() and (candidate / "SKILL.md").exists():
        return candidate
    candidate2 = script_dir / ".claude" / "skills" / SKILL_NAME
    if candidate2.exists() and (candidate2 / "SKILL.md").exists():
        return candidate2
    return script_dir


def get_project_root() -> Path:
    start = Path(__file__).resolve().parent
    for candidate in [start] + list(start.parents):
        if (candidate / "src" / "ppt_pro_max").exists() and (candidate / "component_library").exists():
            return candidate
    return start


def _ignore_patterns(directory: str, contents: list[str]) -> list[str]:
    ignored = []
    for name in contents:
        if name in EXCLUDE_NAMES:
            ignored.append(name)
        elif name.endswith(".pyc"):
            ignored.append(name)
        elif name.endswith(".egg-info"):
            ignored.append(name)
    return ignored


def copy_skill(source_dir: Path, dest_dir: Path, force: bool = False) -> bool:
    if dest_dir.exists():
        if not force:
            _read_skill_version(dest_dir)
            return False
        if dest_dir.resolve() == source_dir.resolve():
            return True
        shutil.rmtree(dest_dir)

    shutil.copytree(source_dir, dest_dir, ignore=_ignore_patterns)
    return True


def _copy_data_assets(project_root: Path, dest_dir: Path) -> None:
    for asset_name in ("component_library", "data"):
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


def detect_platforms(target_dir: Path) -> list[str]:
    detected = []
    for name, info in PLATFORMS.items():
        if (target_dir / info["project_dir"]).exists():
            detected.append(name)
    return detected


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
    has_clib = (skill_dir / "component_library" / "index.db").exists()
    has_data = (skill_dir / "data").exists()
    if skill_md.exists() and has_src:
        ver = _read_skill_version(skill_dir)
        parts = []
        if not has_clib:
            parts.append("no component_library")
        if not has_data:
            parts.append("no data")
        suffix = f" [{', '.join(parts)}]" if parts else ""
        return f"INSTALLED v{ver}{suffix}" if ver else f"INSTALLED{suffix}"
    elif skill_md.exists():
        return "PARTIAL (no src/)"
    return "NOT INSTALLED"


def install_to_dir(source_dir: Path, dest_dir: Path, label: str, force: bool = False, project_root: Path | None = None) -> str | None:
    if dest_dir.resolve() == source_dir.resolve():
        return None
    if dest_dir.exists() and not force:
        ver = _read_skill_version(dest_dir)
        print(f"  [SKIP] {label} — already installed")
        if ver:
            print(f"         Version: {ver}, Latest: {VERSION}")
        print("         Use --force to overwrite")
        return None

    copy_skill(source_dir, dest_dir, force=force)
    print(f"  [OK] {label} -> {dest_dir}")

    if project_root:
        _copy_data_assets(project_root, dest_dir)

    return str(dest_dir)


def install_python_package(source_dir: Path) -> bool:
    project_root = source_dir.parent.parent.parent
    print("  Installing Python package...")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-e", str(project_root)],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode == 0:
            print("  [OK] ppt_pro_max installed")
            return True
        else:
            print(f"  [WARN] pip install failed (exit {result.returncode})")
            print(f"  {result.stderr[:200]}")
            print(f"  You can install manually: pip install -e {project_root}")
            return False
    except Exception as e:
        print(f"  [WARN] pip install failed: {e}")
        print(f"  You can install manually: pip install -e {project_root}")
        return False


def install_ui_ux_pro_max(target_dir: Path, force: bool = False) -> bool:
    print("\n  Checking ui-ux-pro-max skill (required dependency)...")
    from ppt_pro_max.adapters.ui_ux_adapter import is_available, found_path, _UX_FOUND_PATH

    if is_available():
        print(f"  [OK] ui-ux-pro-max found at: {found_path()}")
        return True

    print("  ui-ux-pro-max NOT found. Installing via npx...")
    try:
        result = subprocess.run(
            ["npx", "ui-ux-pro-max-cli", "init", "--ai", "opencode", "--force"],
            capture_output=True, text=True, timeout=120,
            cwd=str(target_dir),
        )
        if result.returncode == 0:
            print("  [OK] ui-ux-pro-max installed via npx")
            return True
        else:
            print(f"  [WARN] npx install failed (exit {result.returncode})")
            if result.stderr:
                print(f"  {result.stderr[:300]}")
    except FileNotFoundError:
        print("  [WARN] npx not found. Install Node.js first: https://nodejs.org/")
    except Exception as e:
        print(f"  [WARN] npx install failed: {e}")

    print()
    print("  *** ui-ux-pro-max is REQUIRED. Install it manually: ***")
    print("  npx ui-ux-pro-max-cli init --ai opencode")
    print("  Or set UX_PRO_MAX_DIR environment variable to its location.")
    return False


def check_installation(target_dir: Path) -> None:
    print(f"\nPPT Design Skill v{VERSION} — Installation Check")
    print(f"Project dir : {target_dir}")
    print(f"Home dir    : {Path.home()}")
    print()

    print("Platform detections (project-level):")
    detected = detect_platforms(target_dir)
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


def main():
    parser = argparse.ArgumentParser(
        prog="ppt-design-install",
        description=f"PPT Design Skill v{VERSION} — One-click installer",
    )
    parser.add_argument("--platform", "-p", help="Specific platform (claude, opencode, codex, cursor, all)")
    parser.add_argument("--target", "-t", help="Target project directory (default: current)")
    parser.add_argument("--global", dest="global_install", action="store_true",
                        help="Install to global home dirs for all platforms")
    parser.add_argument("--force", "-f", action="store_true", help="Overwrite existing installation")
    parser.add_argument("--check", "-c", action="store_true", help="Check installation status only")
    parser.add_argument("--no-pip", action="store_true", help="Skip pip install")
    args = parser.parse_args()

    target_dir = Path(args.target) if args.target else Path.cwd()
    source_skill_dir = get_source_skill_dir()
    project_root = get_project_root()

    if args.check:
        check_installation(target_dir)
        return

    print(f"\n{'='*60}")
    print(f"  PPT Design Skill v{VERSION} — Installer")
    print("  40,000+ style combos · 28 design quality upgrades · 5 image engines")
    print(f"{'='*60}\n")

    print(f"  Source: {source_skill_dir}\n")

    all_installed = []

    # --- Global install ---
    if args.global_install:
        print("Installing to global directories...\n")
        for platform, info in PLATFORMS.items():
            dest_dir = Path.home() / info["global_path"] / SKILL_NAME
            result = install_to_dir(source_skill_dir, dest_dir, f"{info['desc']} (global)", force=args.force, project_root=project_root)
            if result:
                all_installed.append(result)
        print()

    # --- Project-level install ---
    if args.platform == "all":
        platforms_to_install = list(PLATFORMS.keys())
    elif args.platform:
        platforms_to_install = [args.platform]
    else:
        detected = detect_platforms(target_dir)
        if detected:
            platforms_to_install = detected
            print(f"Auto-detected platforms: {', '.join(PLATFORMS[p]['desc'] for p in detected)}")
        else:
            print("No AI coding platforms detected in current directory.")
            print("Installing for common platforms (claude, opencode, codex, cursor)...")
            platforms_to_install = ["claude", "opencode", "codex", "cursor"]

    print(f"\nInstalling for: {', '.join(PLATFORMS.get(p, {}).get('desc', p) for p in platforms_to_install)}")
    print()

    for platform in platforms_to_install:
        info = PLATFORMS.get(platform)
        if not info:
            print(f"  [SKIP] Unknown platform: {platform}")
            continue
        dest_dir = target_dir / info["project_dir"] / "skills" / SKILL_NAME
        result = install_to_dir(source_skill_dir, dest_dir, f"{info['desc']} (project)", force=args.force, project_root=project_root)
        if result:
            all_installed.append(result)

    # Also copy AGENTS.md to project root
    agents_src = source_skill_dir / "AGENTS.md"
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
        install_python_package(source_skill_dir)

    # --- ui-ux-pro-max skill (required) ---
    install_ui_ux_pro_max(target_dir, force=args.force)

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
