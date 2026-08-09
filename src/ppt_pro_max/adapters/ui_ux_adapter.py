"""Adapter for design knowledge search — now fully integrated.

All design search functionality (BM25 search, design system generation,
reasoning) is provided by the bundled design_search module which uses
CSV data from ppt_pro_max/data/. No external dependency required.
"""

from __future__ import annotations

from typing import Any

from ppt_pro_max.adapters.design_search import (
    get_design_system as _get_design_system,
    is_available as _is_available,
    search as _search,
)
from ppt_pro_max.adapters.design_search import (
    DesignSystemGenerator,
)

_gen = DesignSystemGenerator()


def is_available() -> bool:
    return _is_available()


def search_design(query: str, domain: str | None = None, max_results: int = 3) -> list[dict[str, Any]]:
    try:
        return _search(query, domain, max_results)
    except Exception:
        return []


def get_design_system(query: str, **kwargs: Any) -> dict[str, Any]:
    try:
        return _get_design_system(query, **kwargs)
    except Exception:
        from ppt_pro_max.adapters.design_search import _normalize_design_system
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
    try:
        return _search(query, "style", max_results)
    except Exception:
        return []


def search_color(query: str, max_results: int = 2) -> list[dict[str, Any]]:
    try:
        return _search(query, "color", max_results)
    except Exception:
        return []


def search_typography(query: str, max_results: int = 2) -> list[dict[str, Any]]:
    try:
        return _search(query, "typography", max_results)
    except Exception:
        return []


def search_landing(query: str, max_results: int = 2) -> list[dict[str, Any]]:
    try:
        return _search(query, "landing", max_results)
    except Exception:
        return []


def search_reasoning(category: str) -> dict[str, Any]:
    try:
        return _gen.apply_reasoning(category)
    except Exception:
        return {}
