#!/usr/bin/env python3
"""PPT Design Skill — One-click installer.

Auto-detects AI coding platform, installs skill files + dependencies.

Usage:
    python install.py                  # Auto-detect + install
    python install.py --platform claude  # Install for specific platform
    python install.py --platform all    # Install for all platforms
    python install.py --global         # Install to home directory
    python install.py --check          # Check current installation status
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

REPO_URL = "https://github.com/sunchaokun/PPT-Design-Skill"
SKILL_NAME = "ppt-design-skill"
VERSION = "0.2.0"

PLATFORMS: dict[str, dict[str, str]] = {
    "claude":    {"dir": ".claude",    "desc": "Claude Code"},
    "opencode":  {"dir": ".opencode",  "desc": "OpenCode"},
    "codex":     {"dir": ".codex",     "desc": "Codex / OpenClaw"},
    "cursor":    {"dir": ".cursor",    "desc": "Cursor / Windsurf"},
    "windsurf":  {"dir": ".windsurf",  "desc": "Windsurf"},
    "roocode":   {"dir": ".roo",       "desc": "RooCode"},
    "gemini":    {"dir": ".gemini",    "desc": "Gemini CLI"},
    "trae":      {"dir": ".trae",      "desc": "Trae"},
    "continue":  {"dir": ".continue",  "desc": "Continue"},
    "droid":     {"dir": ".factory",   "desc": "Droid (Factory)"},
    "kilocode":  {"dir": ".kilocode",  "desc": "KiloCode"},
    "augment":   {"dir": ".augment",   "desc": "Augment"},
    "copilot":   {"dir": ".github",    "desc": "GitHub Copilot"},
}


def detect_platforms(target_dir: Path) -> list[str]:
    detected = []
    for name, info in PLATFORMS.items():
        if (target_dir / info["dir"]).exists():
            detected.append(name)
    return detected


def get_skill_source_dir() -> Path:
    script_dir = Path(__file__).resolve().parent
    for candidate in [script_dir, script_dir.parent]:
        if (candidate / "SKILL.md").exists() and (candidate / "src" / "ppt_pro_max").exists():
            return candidate
    return script_dir


def install_skill_files(target_dir: Path, platform: str, source_dir: Path, force: bool = False) -> list[str]:
    info = PLATFORMS.get(platform)
    if not info:
        print(f"  [SKIP] Unknown platform: {platform}")
        return []

    platform_dir = target_dir / info["dir"] / "skills" / SKILL_NAME
    scripts_dir = platform_dir / "scripts"

    if platform_dir.exists() and not force:
        print(f"  [SKIP] {info['desc']} — already installed ({platform_dir})")
        print(f"         Use --force to overwrite")
        return []

    platform_dir.mkdir(parents=True, exist_ok=True)
    scripts_dir.mkdir(parents=True, exist_ok=True)

    skill_md_source = source_dir / "SKILL.md"
    if skill_md_source.exists():
        shutil.copy2(skill_md_source, platform_dir / "SKILL.md")

    script_source = source_dir / ".claude" / "skills" / SKILL_NAME / "scripts" / "generate_ppt.py"
    if script_source.exists():
        shutil.copy2(script_source, scripts_dir / "generate_ppt.py")

    agents_source = source_dir / "AGENTS.md"
    if agents_source.exists():
        shutil.copy2(agents_source, target_dir / "AGENTS.md")

    claude_source = source_dir / "CLAUDE.md"
    if claude_source.exists() and platform == "claude":
        shutil.copy2(claude_source, target_dir / "CLAUDE.md")

    installed = [str(platform_dir)]
    print(f"  [OK] {info['desc']} → {platform_dir}")
    return installed


def install_python_package(source_dir: Path) -> bool:
    print("  Installing Python package...")
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "-e", str(source_dir)],
            capture_output=True, text=True, timeout=120,
        )
        print("  [OK] ppt_pro_max installed")
        return True
    except Exception as e:
        print(f"  [WARN] pip install failed: {e}")
        print(f"  You can install manually: pip install -e {source_dir}")
        return False


def check_dependencies() -> dict[str, bool]:
    deps = {}
    try:
        import pptx
        deps["python-pptx"] = True
    except ImportError:
        deps["python-pptx"] = False

    try:
        from PIL import Image
        deps["Pillow"] = True
    except ImportError:
        deps["Pillow"] = False

    return deps


def check_installation(target_dir: Path) -> None:
    print(f"\nPPT Design Skill v{VERSION} — Installation Check")
    print(f"Target: {target_dir}\n")

    print("Platform detections:")
    detected = detect_platforms(target_dir)
    if detected:
        for p in detected:
            info = PLATFORMS.get(p, {})
            skill_dir = target_dir / info.get("dir", "") / "skills" / SKILL_NAME
            skill_md = skill_dir / "SKILL.md"
            status = "INSTALLED" if skill_md.exists() else "NOT INSTALLED"
            print(f"  {info.get('desc', p):25s} [{status}]")
    else:
        print("  No AI coding platforms detected")

    print("\nPython dependencies:")
    deps = check_dependencies()
    for name, ok in deps.items():
        print(f"  {name:20s} [{'OK' if ok else 'MISSING'}]")

    try:
        from ppt_pro_max import generate_ppt
        print(f"  {'ppt_pro_max':20s} [OK]")
    except ImportError:
        print(f"  {'ppt_pro_max':20s} [MISSING]")

    print()


def main():
    parser = argparse.ArgumentParser(
        prog="ppt-design-install",
        description=f"PPT Design Skill v{VERSION} — One-click installer",
    )
    parser.add_argument("--platform", "-p", help="Specific platform (claude, opencode, codex, cursor, all)")
    parser.add_argument("--target", "-t", help="Target directory (default: current)")
    parser.add_argument("--global", dest="global_install", action="store_true", help="Install to home directory")
    parser.add_argument("--force", "-f", action="store_true", help="Overwrite existing installation")
    parser.add_argument("--check", "-c", action="store_true", help="Check installation status only")
    parser.add_argument("--no-pip", action="store_true", help="Skip pip install")
    args = parser.parse_args()

    target_dir = Path(args.target) if args.target else Path.cwd()
    if args.global_install:
        target_dir = Path.home()

    if args.check:
        check_installation(target_dir)
        return

    source_dir = get_skill_source_dir()

    print(f"\n{'='*60}")
    print(f"  PPT Design Skill v{VERSION} — Installer")
    print(f"  40,000+ style combinations · 5 image engines")
    print(f"{'='*60}\n")

    # Detect platforms
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
            print("Installing for all common platforms...")
            platforms_to_install = ["claude", "opencode", "codex", "cursor"]

    print(f"\nInstalling for: {', '.join(PLATFORMS.get(p, {}).get('desc', p) for p in platforms_to_install)}")
    print()

    # Install skill files
    all_installed = []
    for platform in platforms_to_install:
        installed = install_skill_files(target_dir, platform, source_dir, force=args.force)
        all_installed.extend(installed)

    # Install Python package
    if not args.no_pip:
        print()
        install_python_package(source_dir)

    # Summary
    print(f"\n{'='*60}")
    if all_installed:
        print(f"  Installed successfully!")
        print(f"\n  Installed to:")
        for path in all_installed:
            print(f"    {path}")
    else:
        print(f"  No new installations (use --force to overwrite)")

    print(f"\n  Next steps:")
    print(f"    1. Restart your AI coding assistant")
    print(f"    2. Try: 'Generate a PPT for AI startup investor pitch'")
    print(f"    3. Or: python -m ppt_pro_max 'AI pitch' --style 'dark cyberpunk'")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
