"""Check Python and PPTX->PDF->PNG runtime dependencies without mutation."""

from __future__ import annotations

import importlib
import shutil
import sys


def main() -> int:
    print(f"Python: {sys.version.split()[0]}")
    ok = sys.version_info >= (3, 10)
    print(f"  {'OK' if ok else 'MISSING'} Python >= 3.10")

    for module, label in (
        ("pptx_designer", "pptx-designer"),
        ("pptx", "python-pptx"),
        ("PIL", "Pillow"),
    ):
        try:
            loaded = importlib.import_module(module)
            version = getattr(loaded, "__version__", "")
            suffix = f" {version}" if version else ""
            print(f"  OK {label}{suffix}")
        except Exception as exc:  # pragma: no cover - diagnostic path
            ok = False
            print(f"  MISSING {label}: {exc}")

    for command, label in (("soffice", "LibreOffice"), ("pdftoppm", "Poppler pdftoppm")):
        found = shutil.which(command)
        print(f"  {'OK' if found else 'MISSING'} {label}{f' ({found})' if found else ''}")

    print("  INFO PowerPoint COM is checked by render_pptx.ps1 on Windows.")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
