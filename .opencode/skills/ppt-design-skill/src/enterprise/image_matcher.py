"""ImageMatcher — auto-assign images from pool to slides."""

from __future__ import annotations

import os
import re
from typing import Any

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = None


_GOAL_SIZE_PREF: dict[str, list[str]] = {
    "hook": ["background", "scene", "icon"],
    "cta": ["background", "scene", "icon"],
    "problem": ["scene", "background", "icon"],
    "solution": ["scene", "background", "icon"],
    "features": ["scene", "icon", "background"],
    "data": ["scene", "icon", "background"],
    "content": ["scene", "background", "icon"],
    "exercise": ["icon", "scene", "background"],
    "code": ["icon", "scene", "background"],
    "overview": ["scene", "background", "icon"],
}

_GOAL_IMAGE_PROMPTS: dict[str, str] = {
    "hook": "professional presentation cover, {title}, minimalist, high quality, no text overlay",
    "problem": "business challenge visualization, dark tones, cinematic, {title}",
    "solution": "technology solution dashboard, clean, modern, {title}",
    "features": "product features grid, bright, professional, {title}",
    "data": "technical architecture diagram, clean lines, professional, {title}",
    "market": "market growth chart, data visualization, blue tones, {title}",
    "cta": "professional contact slide, clean, inviting, {title}",
    "content": "professional business presentation, clean composition, {title}",
    "exercise": "workshop exercise, interactive, engaging, {title}",
    "code": "code editor screenshot, dark theme, professional, {title}",
    "overview": "business overview, clean layout, professional, {title}",
}


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


def classify_image_size(image_path: str) -> str:
    if not os.path.isfile(image_path):
        return "unknown"
    if PILImage is None:
        return "unknown"
    try:
        with PILImage.open(image_path) as img:
            width = img.width
    except Exception:
        return "unknown"
    if width > 1500:
        return "background"
    if width > 800:
        return "scene"
    return "icon"


def assign_images_by_size(
    image_pool: list[str],
    page_designs: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if not image_pool:
        return page_designs

    classified: dict[str, list[str]] = {"background": [], "scene": [], "icon": []}
    for path in image_pool:
        size_cat = classify_image_size(path)
        if size_cat in classified:
            classified[size_cat].append(path)

    used: set[str] = set()

    for design in page_designs:
        if design.get("image"):
            continue

        goal = (design.get("goal") or "content").lower()
        prefs = _GOAL_SIZE_PREF.get(goal, ["scene", "background", "icon"])

        assigned = False
        for pref in prefs:
            for img_path in classified.get(pref, []):
                if img_path not in used:
                    design["image"] = img_path
                    used.add(img_path)
                    assigned = True
                    break
            if assigned:
                break

    remaining = [p for cat in classified.values() for p in cat if p not in used]
    for design in page_designs:
        if design.get("image"):
            continue
        if not remaining:
            break
        design["image"] = remaining.pop(0)

    for img_path in classified.get("icon", []):
        if img_path in used:
            continue
        for design in page_designs:
            cards = design.get("cards")
            if not cards:
                continue
            for card in cards:
                if not card.get("image"):
                    card["image"] = img_path
                    used.add(img_path)
                    break
            if img_path in used:
                break

    return page_designs


def auto_generate_image_prompts(
    page_designs: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    for design in page_designs:
        if design.get("image"):
            continue
        if design.get("image_prompt"):
            continue

        goal = (design.get("goal") or "content").lower()
        title = design.get("title", "")
        template = _GOAL_IMAGE_PROMPTS.get(goal, _GOAL_IMAGE_PROMPTS["content"])
        design["image_prompt"] = template.format(title=title)

    return page_designs
