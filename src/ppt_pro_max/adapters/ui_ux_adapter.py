"""Adapter for ui-ux-pro-max search engine and design knowledge.

**ui-ux-pro-max is a required dependency.** It must be installed first:

    npm install -g ui-ux-pro-max-cli
    uipro init --ai <your-platform>

Search order for locating the skill:
  1. Environment variable UX_PRO_MAX_DIR (explicit override)
  2. Project-local .opencode/skills/ui-ux-pro-max/
  3. Project-local .claude/skills/ui-ux-pro-max/
  4. Skill-sibling: <ppt-design-skill's parent>/ui-ux-pro-max/
  5. User-global ~/.opencode/skills/ui-ux-pro-max/
  6. User-global ~/.claude/skills/ui-ux-pro-max/
  7. User-global ~/.config/opencode/skills/ui-ux-pro-max/
  8. sys.path / pip / conda

Once found, scripts/ is added to sys.path so that the relative imports
inside ui-ux-pro-max (``from core import search``) work correctly.
The package is NOT copied into this project.
"""

from __future__ import annotations

import csv
import os
import sys
from pathlib import Path
from typing import Any

_UX_AVAILABLE = False
_UX_DATA_DIR: Path | None = None
_UX_SEARCH = None
_UX_DETECT_DOMAIN = None
_UX_GEN = None
_UX_FOUND_PATH: str | None = None

_INSTALL_INSTRUCTIONS = (
    "ui-ux-pro-max is required but not found.\n"
    "Install it first:\n"
    "  npm install -g ui-ux-pro-max-cli\n"
    "  uipro init --ai <your-platform>\n"
    "Or set the UX_PRO_MAX_DIR environment variable to its location.\n"
    "Searched locations:\n"
    "  - $UX_PRO_MAX_DIR\n"
    "  - <project>/.opencode/skills/ui-ux-pro-max/\n"
    "  - <project>/.claude/skills/ui-ux-pro-max/\n"
    "  - <ppt-design-skill sibling>/ui-ux-pro-max/\n"
    "  - ~/.opencode/skills/ui-ux-pro-max/\n"
    "  - ~/.claude/skills/ui-ux-pro-max/\n"
    "  - ~/.config/opencode/skills/ui-ux-pro-max/\n"
)


class UiUxProMaxNotFoundError(ImportError):
    pass


def _search_skill_root() -> Path | None:
    env_dir = os.environ.get("UX_PRO_MAX_DIR")
    if env_dir:
        p = Path(env_dir)
        if (p / "scripts" / "core.py").exists():
            return p

    try:
        this_file = Path(__file__).resolve()
    except Exception:
        this_file = Path.cwd()

    if not this_file.exists():
        this_file = Path.cwd()

    home = Path.home()

    skill_dir_candidates = []

    try:
        project_root = this_file.parents[3]
        skill_dir_candidates.extend([
            project_root / ".opencode" / "skills" / "ui-ux-pro-max",
            project_root / ".claude" / "skills" / "ui-ux-pro-max",
        ])
    except IndexError:
        pass

    for parent in this_file.parents:
        if (parent / "SKILL.md").exists() and (parent / "src").exists():
            sibling = parent.parent / "ui-ux-pro-max"
            if sibling != parent:
                skill_dir_candidates.append(sibling)
            break

    skill_dir_candidates.extend([
        home / ".opencode" / "skills" / "ui-ux-pro-max",
        home / ".claude" / "skills" / "ui-ux-pro-max",
        home / ".config" / "opencode" / "skills" / "ui-ux-pro-max",
    ])

    for c in skill_dir_candidates:
        if (c / "scripts" / "core.py").exists():
            return c

    for p_str in sys.path:
        p = Path(p_str) / "ui-ux-pro-max"
        if (p / "scripts" / "core.py").exists():
            return p

    try:
        import importlib.util
        spec = importlib.util.find_spec("ui_ux_pro_max")
        if spec and spec.origin:
            return Path(spec.origin).parent
    except (ModuleNotFoundError, ValueError):
        pass

    return None


def _try_import() -> None:
    global _UX_AVAILABLE, _UX_DATA_DIR, _UX_SEARCH, _UX_DETECT_DOMAIN, _UX_GEN, _UX_FOUND_PATH
    if _UX_AVAILABLE:
        return

    skill_root = _search_skill_root()
    if skill_root is None:
        _UX_AVAILABLE = False
        return

    try:
        scripts_dir = str(skill_root / "scripts")
        if scripts_dir not in sys.path:
            sys.path.insert(0, scripts_dir)

        from core import search as _search, detect_domain as _detect, DATA_DIR as _data_dir
        from design_system import DesignSystemGenerator as _DSG

        _UX_SEARCH = _search
        _UX_DETECT_DOMAIN = _detect
        _UX_DATA_DIR = Path(_data_dir)
        _UX_GEN = _DSG()
        _UX_FOUND_PATH = str(skill_root)
        _UX_AVAILABLE = True
    except Exception:
        _UX_AVAILABLE = False


_try_import()


def _require() -> None:
    if not _UX_AVAILABLE:
        raise UiUxProMaxNotFoundError(_INSTALL_INSTRUCTIONS)


def is_available() -> bool:
    return _UX_AVAILABLE


def found_path() -> str | None:
    return _UX_FOUND_PATH


def search_design(query: str, domain: str | None = None, max_results: int = 3) -> list[dict[str, Any]]:
    _require()
    try:
        result = _UX_SEARCH(query, domain, max_results)
        return result.get("results", [])
    except Exception:
        return []


def get_design_system(query: str, **kwargs) -> dict[str, Any]:
    _require()
    try:
        ds = _UX_GEN.generate(
            query,
            variance=kwargs.get("variance"),
            motion=kwargs.get("motion"),
            density=kwargs.get("density"),
        )
        return _normalize_design_system(ds)
    except Exception:
        return _normalize_design_system({
            "project_name": query,
            "category": "General",
            "colors": {},
            "typography": {},
            "style": {},
            "pattern": {},
            "key_effects": "",
            "anti_patterns": "",
            "decision_rules": {},
            "severity": "MEDIUM",
            "dials": {},
            "motion_snippet": {},
            "spacing_scale": None,
        })


def search_style(query: str, max_results: int = 3) -> list[dict[str, Any]]:
    _require()
    try:
        result = _UX_SEARCH(query, "style", max_results)
        return result.get("results", [])
    except Exception:
        return []


def search_color(query: str, max_results: int = 2) -> list[dict[str, Any]]:
    _require()
    try:
        result = _UX_SEARCH(query, "color", max_results)
        return result.get("results", [])
    except Exception:
        return []


def search_typography(query: str, max_results: int = 2) -> list[dict[str, Any]]:
    _require()
    try:
        result = _UX_SEARCH(query, "typography", max_results)
        return result.get("results", [])
    except Exception:
        return []


def search_landing(query: str, max_results: int = 2) -> list[dict[str, Any]]:
    _require()
    try:
        result = _UX_SEARCH(query, "landing", max_results)
        return result.get("results", [])
    except Exception:
        return []


def search_reasoning(category: str) -> dict[str, Any]:
    _require()
    try:
        return _UX_GEN._apply_reasoning(category, {})
    except Exception:
        return {}


def load_csv(filename: str) -> list[dict[str, str]]:
    _require()
    filepath = _UX_DATA_DIR / filename
    if not filepath.exists():
        return []
    with open(filepath, "r", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _normalize_design_system(ds: dict[str, Any]) -> dict[str, Any]:
    colors = ds.get("colors", {})
    normalized_colors = {}
    for key, val in colors.items():
        nk = key.replace(" ", "-").replace("/", "-").lower()
        normalized_colors[nk] = val

    if "primary" not in normalized_colors and "primary" in colors:
        normalized_colors["primary"] = colors["primary"]

    ds["colors"] = normalized_colors

    style = ds.get("style", {})
    if style:
        ds["style_name"] = style.get("name", "")
        ds["style_effects"] = style.get("effects", "")
        ds["style_keywords"] = style.get("keywords", "")
        ds["style_best_for"] = style.get("best_for", "")
        ds["style_dark_mode"] = style.get("dark_mode", "")
        ds["style_light_mode"] = style.get("light_mode", "")

    pattern = ds.get("pattern", {})
    if pattern:
        ds["pattern_name"] = pattern.get("name", "")
        ds["pattern_sections"] = pattern.get("sections", "")
        ds["pattern_cta_placement"] = pattern.get("cta_placement", "")
        ds["pattern_color_strategy"] = pattern.get("color_strategy", "")
        ds["pattern_conversion"] = pattern.get("conversion", "")

    return ds
