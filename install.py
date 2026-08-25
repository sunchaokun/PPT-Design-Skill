"""Install the Python runtime for PPT Design Skill.

This installs Python packages only. Desktop renderers are detected but are not
silently installed because they require system-level user approval.
"""

from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--images", action="store_true", help="Install stock-image optional dependencies")
    parser.add_argument("--ai-images", action="store_true", help="Install AI-image optional dependencies")
    parser.add_argument(
        "--render-deps",
        action="store_true",
        help="On Windows, install LibreOffice and Poppler through winget",
    )
    args = parser.parse_args()

    package = "pptx-designer"
    if args.images and args.ai_images:
        package = "pptx-designer[images,ai-images]"
    elif args.images:
        package = "pptx-designer[images]"
    elif args.ai_images:
        package = "pptx-designer[ai-images]"

    command = [sys.executable, "-m", "pip", "install", "--upgrade", package]
    print("Installing:", package)
    result = subprocess.run(command, check=False)
    if result.returncode:
        return result.returncode

    if args.render_deps:
        if platform.system() != "Windows":
            print("--render-deps currently supports Windows winget only; install LibreOffice and Poppler manually.")
        else:
            winget = shutil.which("winget")
            if not winget:
                print("winget was not found; install LibreOffice and Poppler manually.")
            else:
                for package_id, label in (
                    ("TheDocumentFoundation.LibreOffice", "LibreOffice"),
                    ("oschwartz10612.Poppler", "Poppler"),
                ):
                    print("Installing:", label)
                    dep = subprocess.run(
                        [winget, "install", "--id", package_id, "--exact", "--accept-source-agreements", "--accept-package-agreements"],
                        check=False,
                    )
                    if dep.returncode:
                        print(f"Warning: {label} installation failed with exit code {dep.returncode}.")

    print("Python runtime installed. Run skill/scripts/check_runtime.py next.")
    print("LibreOffice and Poppler are optional system dependencies for the headless render fallback.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
