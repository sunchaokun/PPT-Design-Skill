"""Install the PPT Design Skill bundle for coding assistants."""

from __future__ import annotations

import argparse
import importlib.util
from importlib.metadata import PackageNotFoundError, version as package_version
import platform as host_platform
import shutil
import subprocess
import sys
from pathlib import Path

try:
    from installer.platforms import PLATFORMS, global_path, normalize, project_path
except ModuleNotFoundError:  # direct execution: python installer/install.py
    from platforms import PLATFORMS, global_path, normalize, project_path

SKILL_NAME = "ppt-design-skill"
SKILL_VERSION = "1.3"
ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "skill"


def copy_bundle(destination: Path, force: bool) -> bool:
    if destination.exists() and not force:
        print(f"[SKIP] {destination} exists; use --force to replace it")
        return False
    destination.parent.mkdir(parents=True, exist_ok=True)
    # A skill installation is a complete bundle, not an additive overlay.
    # Older releases shipped a ``src/ppt_pro_max`` tree and legacy entrypoint
    # files. Replace the dedicated skill directory so those files cannot be
    # discovered as part of the new Skill runtime path.
    staging = destination.with_name(f".{destination.name}.installing")
    if staging.exists():
        shutil.rmtree(staging)
    shutil.copytree(
        SOURCE,
        staging,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
    )
    # Build the replacement completely before removing the current bundle. If
    # copying fails, the existing installation is left intact.
    if destination.exists():
        shutil.rmtree(destination)
    staging.replace(destination)
    print(f"[OK] skill bundle -> {destination}")
    return True


def install_python_package(with_images: bool = False, with_ai_images: bool = False) -> bool:
    package = "pptx-designer"
    extras = []
    if with_images:
        extras.append("images")
    if with_ai_images:
        extras.append("ai-images")
    if extras:
        package += "[" + ",".join(extras) + "]"
    before = installed_package_version("pptx-designer")
    print(f"Checking pptx-designer (installed: {before or 'missing'})...")
    result = subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", package], check=False)
    if result.returncode == 0:
        after = installed_package_version("pptx-designer") or "unknown version"
        status = "updated" if before and before != after else "ready"
        print(f"[OK] pptx-designer {after} ({status})")
    return result.returncode == 0


def installed_package_version(distribution_name: str) -> str | None:
    """Return the installed distribution version without changing the environment."""
    try:
        return package_version(distribution_name)
    except PackageNotFoundError:
        return None


def render_dependency_status() -> dict[str, bool]:
    scripts_dir = ROOT / "skill" / "scripts"
    if str(scripts_dir) not in sys.path:
        sys.path.insert(0, str(scripts_dir))
    from runtime_deps import resolve_executable
    return {name: resolve_executable(command) is not None for name, command in (("libreoffice", "soffice"), ("poppler", "pdftoppm"))}


def install_render_deps() -> None:
    if host_platform.system() != "Windows":
        print("[WARN] --render-deps currently supports Windows winget only")
        return
    winget = shutil.which("winget")
    if not winget:
        print("[WARN] winget not found; install LibreOffice and Poppler manually")
        return
    for package_id, label in (("TheDocumentFoundation.LibreOffice", "LibreOffice"), ("oschwartz10612.Poppler", "Poppler")):
        print(f"Installing {label}...")
        subprocess.run([winget, "install", "--id", package_id, "--exact", "--accept-source-agreements", "--accept-package-agreements"], check=False)


def check() -> int:
    package_ok = True
    print("Python packages:")
    for module, label, distribution in (
        ("pptx_designer", "pptx-designer", "pptx-designer"),
        ("pptx", "python-pptx", "python-pptx"),
        ("PIL", "Pillow", "Pillow"),
    ):
        found = importlib.util.find_spec(module) is not None
        package_ok = package_ok and found
        installed = installed_package_version(distribution) if found else None
        suffix = f" {installed}" if installed else ""
        print(f"  {'OK' if found else 'MISSING'} {label}{suffix}")
    print("Render dependencies:")
    for name, available in render_dependency_status().items():
        package_ok = package_ok and available
        print(f"  {'OK' if available else 'MISSING'} {name}")
    return 0 if package_ok else 1


def selected_platforms(name: str | None) -> list[str]:
    if not name or name.lower() == "all":
        return list(PLATFORMS)
    key = normalize(name)
    if key not in PLATFORMS:
        raise SystemExit(f"Unknown platform: {name}. Available: {', '.join(PLATFORMS)}")
    return [key]


def main() -> int:
    parser = argparse.ArgumentParser(description="Install PPT Design Skill")
    parser.add_argument("--platform", "-p", help="Platform key or all")
    parser.add_argument("--all", action="store_true", help="Install to all supported platforms")
    parser.add_argument("--target", type=Path, default=Path.cwd(), help="Project root for --project")
    parser.add_argument("--project", action="store_true", help="Install to project-local skill roots")
    parser.add_argument("--no-global", action="store_true", help="Skip global installation")
    parser.add_argument("--force", "-f", action="store_true")
    parser.add_argument("--no-pip", action="store_true")
    parser.add_argument("--images", action="store_true")
    parser.add_argument("--ai-images", action="store_true")
    parser.add_argument("--render-deps", action="store_true")
    parser.add_argument("--check", "-c", action="store_true")
    args = parser.parse_args()
    if args.check:
        return check()

    keys = selected_platforms("all" if args.all else args.platform)
    installed = []
    home = Path.home()
    if not args.no_global:
        for key in keys:
            destination = global_path(PLATFORMS[key], home) / SKILL_NAME
            if copy_bundle(destination, args.force):
                installed.append(destination)
    if args.project:
        for key in keys:
            destination = project_path(PLATFORMS[key], args.target) / SKILL_NAME
            if copy_bundle(destination, args.force):
                installed.append(destination)
    if not args.no_pip and not install_python_package(args.images, args.ai_images):
        print("[WARN] pip installation failed")
    if args.render_deps:
        install_render_deps()
    print(f"Installed {SKILL_NAME} {SKILL_VERSION}: {len(installed)} skill bundle(s). Restart the coding assistant to reload skills.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
