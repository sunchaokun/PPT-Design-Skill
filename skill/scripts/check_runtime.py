"""Check Python and PPTX->PDF->PNG runtime dependencies without mutation."""

from __future__ import annotations

import importlib
import sys

from runtime_deps import resolve_executable


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

    fallback_ready = True
    for command, label in (("soffice", "LibreOffice"), ("pdftoppm", "Poppler pdftoppm")):
        found = resolve_executable(command)
        fallback_ready = fallback_ready and bool(found)
        print(f"  {'OK' if found else 'MISSING'} {label}{f' ({found})' if found else ''}")

    if fallback_ready:
        print("  OK LibreOffice + Poppler fallback renderer is available.")
    else:
        print("  INFO Fallback renderer is unavailable; PowerPoint COM is checked only by render_pptx.ps1.")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
