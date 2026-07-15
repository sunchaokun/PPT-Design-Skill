import os
import sys
from pathlib import Path

_ux_dir = os.environ.get("UX_PRO_MAX_DIR")
if not _ux_dir:
    for candidate in [
        Path(__file__).resolve().parents[1] / ".opencode" / "skills" / "ui-ux-pro-max",
        Path(__file__).resolve().parents[1] / ".claude" / "skills" / "ui-ux-pro-max",
        Path("E:/ui-ux-pro-max-skill-main/src/ui-ux-pro-max"),
    ]:
        if (candidate / "scripts" / "core.py").exists():
            os.environ["UX_PRO_MAX_DIR"] = str(candidate)
            break

if os.environ.get("UX_PRO_MAX_DIR"):
    scripts_dir = str(Path(os.environ["UX_PRO_MAX_DIR"]) / "scripts")
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
