"""Adapter for ui-ux-pro-max search engine and design knowledge."""

from __future__ import annotations

import csv
import os
from pathlib import Path
from typing import Any


DATA_DIR: Path | None = None
_SEARCH_AVAILABLE = False

try:
    from ui_ux_pro_max.scripts.core import search as _ux_search, detect_domain, DATA_DIR as _UX_DATA_DIR
    from ui_ux_pro_max.scripts.design_system import generate_design_system as _ux_generate_ds

    DATA_DIR = Path(_UX_DATA_DIR)
    _SEARCH_AVAILABLE = True
except ImportError:
    pass


def is_search_available() -> bool:
    return _SEARCH_AVAILABLE


def search_design(query: str, domain: str | None = None, max_results: int = 3) -> list[dict[str, Any]]:
    if not _SEARCH_AVAILABLE:
        return _fallback_search(query, domain, max_results)
    try:
        return _ux_search(query, domain, max_results)
    except TypeError:
        return _ux_search(query, domain)


def get_design_system(query: str, **kwargs) -> dict[str, Any]:
    if not _SEARCH_AVAILABLE:
        return _fallback_design_system(query)
    return _ux_generate_ds(query, **kwargs)


def load_csv(filename: str) -> list[dict[str, str]]:
    if DATA_DIR is None:
        return []
    filepath = DATA_DIR / filename
    if not filepath.exists():
        return []
    with open(filepath, "r", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _fallback_search(query: str, domain: str | None = None, max_results: int = 3) -> list[dict]:
    return []


def _fallback_design_system(query: str) -> dict[str, Any]:
    return {
        "colors": {
            "primary": "#2563EB",
            "on-primary": "#FFFFFF",
            "secondary": "#64748B",
            "accent": "#F97316",
            "background": "#FFFFFF",
            "foreground": "#1E293B",
            "muted": "#F1F5F9",
            "muted-foreground": "#94A3B8",
            "border": "#E2E8F0",
            "destructive": "#EF4444",
        },
        "typography": {
            "heading": "Inter",
            "body": "Inter",
        },
        "spacing": {"unit": 4},
    }
