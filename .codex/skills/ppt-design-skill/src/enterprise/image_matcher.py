"""ImageMatcher — auto-assign images from pool to slides."""

from __future__ import annotations

import os
import re
from typing import Any


def match_images(
    image_pool: list[str],
    page_designs: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if not image_pool:
        return page_designs

    pool_basenames: dict[str, str] = {}
    for path in image_pool:
        base = os.path.splitext(os.path.basename(path))[0].lower()
        pool_basenames[base] = path

    used: set[str] = set()
    for design in page_designs:
        if design.get("image"):
            continue

        goal = (design.get("goal") or "").lower()
        title = (design.get("title") or "").lower()

        matched = _find_best_match(goal, title, pool_basenames, used)
        if matched:
            design["image"] = matched
            used.add(os.path.splitext(os.path.basename(matched))[0].lower())

    return page_designs


_GOAL_KEYWORDS: dict[str, list[str]] = {
    "hook": ["hero", "cover", "splash", "banner", "title"],
    "problem": ["pain", "problem", "issue", "challenge"],
    "solution": ["solution", "fix", "resolve", "answer"],
    "product": ["product", "dashboard", "app", "platform", "screenshot", "demo"],
    "features": ["feature", "capability", "function"],
    "team": ["team", "people", "group", "staff", "founder"],
    "market": ["market", "chart", "growth", "trend", "size"],
    "business": ["business", "model", "revenue", "pricing"],
    "competition": ["competitor", "versus", "compare", "landscape"],
    "traction": ["traction", "metric", "number", "stat", "kpi"],
    "financials": ["financial", "finance", "money", "projection", "revenue"],
    "ask": ["ask", "invest", "fund", "cta", "closing"],
    "process": ["process", "flow", "workflow", "step"],
    "education": ["education", "learn", "course", "lesson"],
    "training": ["training", "workshop", "exercise"],
    "report": ["report", "summary", "overview", "annual"],
}


def _tokenize(name: str) -> list[str]:
    return re.split(r"[_\-\s]+", name.lower())


def _find_best_match(
    goal: str,
    title: str,
    pool_basenames: dict[str, str],
    used: set[str],
) -> str | None:
    keywords = _GOAL_KEYWORDS.get(goal, [goal])

    for kw in keywords:
        for base, path in pool_basenames.items():
            if base in used:
                continue
            tokens = _tokenize(base)
            if kw in tokens:
                return path

    title_words = re.findall(r"[a-zA-Z]{4,}", title)
    for word in title_words:
        w = word.lower()
        for base, path in pool_basenames.items():
            if base in used:
                continue
            tokens = _tokenize(base)
            if w in tokens:
                return path

    for base, path in pool_basenames.items():
        if base not in used:
            return path

    return None
