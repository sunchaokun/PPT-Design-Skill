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

SKILL_FILES = [
    "SKILL.md",
    "AGENTS.md",
    "pyproject.toml",
    ".env.example",
]

SKILL_DIRS = [
    "src",
    "docs",
    "scripts",
]

# Each platform: project-level dir, global skill path (relative to HOME), description
# Global paths based on official docs research:
#   Claude Code: ~/.claude/skills/<name>/SKILL.md
#   OpenCode:    ~/.config/opencode/skills/<name>/SKILL.md
#   Codex:       ~/.agents/skills/<name>/SKILL.md  (current standard)
#   Cursor:      ~/.agents/plugins/<name>/SKILL.md  (Open Plugins spec)
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


def get_source_dir() -> Path:
    script_dir = Path(__file__).resolve().parent
    if (script_dir / "SKILL.md").exists() and (script_dir / "src" / "ppt_pro_max").exists():
        return script_dir
    return script_dir


def detect_platforms(target_dir: Path) -> list[str]:
    detected = []
    for name, info in PLATFORMS.items():
        if (target_dir / info["project_dir"]).exists():
            detected.append(name)
    return detected


def get_project_skill_dir(target_dir: Path, platform: str) -> Path:
    info = PLATFORMS[platform]
    return target_dir / info["project_dir"] / "skills" / SKILL_NAME


def get_global_skill_dir(platform: str) -> Path:
    info = PLATFORMS[platform]
    return Path.home() / info["global_path"] / SKILL_NAME


def _copy_skill_files(dest_dir: Path, source_dir: Path) -> None:
    dest_dir.mkdir(parents=True, exist_ok=True)

    for fname in SKILL_FILES:
        src = source_dir / fname
        if src.exists():
            shutil.copy2(src, dest_dir / fname)

    for dname in SKILL_DIRS:
        src_d = source_dir / dname
        dst_d = dest_dir / dname
        if src_d.exists():
            if dst_d.exists():
                shutil.rmtree(dst_d)
            shutil.copytree(src_d, dst_d)


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
    if skill_md.exists() and has_src:
        ver = _read_skill_version(skill_dir)
        return f"INSTALLED v{ver}" if ver else "INSTALLED"
    elif skill_md.exists():
        return "PARTIAL (no src/)"
    return "NOT INSTALLED"


def install_to_dir(dest_dir: Path, source_dir: Path, label: str, force: bool = False) -> str | None:
    if dest_dir.exists() and not force:
        ver = _read_skill_version(dest_dir)
        print(f"  [SKIP] {label} — already installed")
        print(f"         Path: {dest_dir}")
        if ver:
            print(f"         Version: {ver}, Latest: {VERSION}")
        print(f"         Use --force to overwrite")
        return None

    _copy_skill_files(dest_dir, source_dir)
    print(f"  [OK] {label} -> {dest_dir}")
    return str(dest_dir)


def install_project_level(target_dir: Path, platform: str, source_dir: Path, force: bool = False) -> str | None:
    info = PLATFORMS.get(platform)
    if not info:
        print(f"  [SKIP] Unknown platform: {platform}")
        return None

    skill_dir = get_project_skill_dir(target_dir, platform)
    return install_to_dir(skill_dir, source_dir, f"{info['desc']} (project)", force)


def install_global_level(platform: str, source_dir: Path, force: bool = False) -> str | None:
    info = PLATFORMS.get(platform)
    if not info:
        return None

    skill_dir = get_global_skill_dir(platform)
    return install_to_dir(skill_dir, source_dir, f"{info['desc']} (global)", force)


def install_python_package(source_dir: Path) -> bool:
    print("  Installing Python package...")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-e", str(source_dir)],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode == 0:
            print("  [OK] ppt_pro_max installed")
            return True
        else:
            print(f"  [WARN] pip install failed (exit {result.returncode})")
            print(f"  {result.stderr[:200]}")
            print(f"  You can install manually: pip install -e {source_dir}")
            return False
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
    print(f"Project dir : {target_dir}")
    print(f"Home dir    : {Path.home()}")
    print()

    print("Platform detections (project-level):")
    detected = detect_platforms(target_dir)
    if detected:
        for p in detected:
            info = PLATFORMS.get(p, {})
            skill_dir = get_project_skill_dir(target_dir, p)
            status = _skill_status(skill_dir)
            print(f"  {info.get('desc', p):25s} [{status}]")
    else:
        print("  No AI coding platforms detected in current directory")

    print("\nGlobal installations:")
    for p, info in PLATFORMS.items():
        global_dir = get_global_skill_dir(p)
        status = _skill_status(global_dir)
        if status != "NOT INSTALLED":
            print(f"  {info['desc']:25s} [{status}]  ({global_dir})")

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
    parser.add_argument("--target", "-t", help="Target project directory (default: current)")
    parser.add_argument("--global", dest="global_install", action="store_true",
                        help="Install to global home dirs for all platforms")
    parser.add_argument("--force", "-f", action="store_true", help="Overwrite existing installation")
    parser.add_argument("--check", "-c", action="store_true", help="Check installation status only")
    parser.add_argument("--no-pip", action="store_true", help="Skip pip install")
    args = parser.parse_args()

    target_dir = Path(args.target) if args.target else Path.cwd()
    source_dir = get_source_dir()

    if args.check:
        check_installation(target_dir)
        return

    print(f"\n{'='*60}")
    print(f"  PPT Design Skill v{VERSION} — Installer")
    print(f"  40,000+ style combos · 28 design quality upgrades · 5 image engines")
    print(f"{'='*60}\n")

    all_installed = []

    # --- Global install ---
    if args.global_install:
        print("Installing to global directories...\n")
        for platform in PLATFORMS:
            result = install_global_level(platform, source_dir, force=args.force)
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
        result = install_project_level(target_dir, platform, source_dir, force=args.force)
        if result:
            all_installed.append(result)

    # Also copy AGENTS.md to project root
    agents_source = source_dir / "AGENTS.md"
    if agents_source.exists():
        root_agents = target_dir / "AGENTS.md"
        if not root_agents.exists() or args.force:
            try:
                shutil.copy2(agents_source, root_agents)
            except PermissionError:
                print(f"  [WARN] Cannot overwrite {root_agents} (file locked)")

    # --- pip install ---
    if not args.no_pip:
        print()
        install_python_package(source_dir)

    # --- Summary ---
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
