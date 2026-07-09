"""Adapter for slide_search_core context-aware decision functions."""

from __future__ import annotations

from typing import Any

_SEARCH_AVAILABLE = False

try:
    from slide_search_core import (
        search_with_context,
        get_layout_for_goal,
        get_typography_for_slide,
        get_color_for_emotion,
        get_background_config,
        should_use_full_bleed,
        calculate_pattern_break,
    )

    _SEARCH_AVAILABLE = True
except ImportError:
    pass


def is_available() -> bool:
    return _SEARCH_AVAILABLE


def layout_for_goal(goal: str) -> str:
    if _SEARCH_AVAILABLE:
        try:
            return get_layout_for_goal(goal)
        except Exception:
            pass
    return _fallback_layout(goal)


def typography_for_slide(slide_type: str) -> dict[str, Any]:
    if _SEARCH_AVAILABLE:
        try:
            return get_typography_for_slide(slide_type)
        except Exception:
            pass
    return _fallback_typography(slide_type)


def color_for_emotion(emotion: str) -> dict[str, Any]:
    if _SEARCH_AVAILABLE:
        try:
            return get_color_for_emotion(emotion)
        except Exception:
            pass
    return _fallback_color(emotion)


def background_config(slide_type: str) -> dict[str, Any]:
    if _SEARCH_AVAILABLE:
        try:
            return get_background_config(slide_type)
        except Exception:
            pass
    return {"image_category": "abstract", "overlay_style": "none"}


def full_bleed(position: int, emotion: str) -> bool:
    if _SEARCH_AVAILABLE:
        try:
            return should_use_full_bleed(position, emotion)
        except Exception:
            pass
    return False


def pattern_break(position: int, total: int) -> bool:
    if _SEARCH_AVAILABLE:
        try:
            return calculate_pattern_break(position, total)
        except Exception:
            pass
    return position in (total // 3, 2 * total // 3)


_LAYOUT_FALLBACK = {
    "hook": "title-slide",
    "problem": "content-with-title",
    "solution": "content-with-title",
    "features": "three-column-cards",
    "traction": "four-metrics",
    "metrics": "four-metrics",
    "comparison": "two-column",
    "testimonial": "quote",
    "cta": "cta-closing",
    "section": "section-header",
    "data": "chart-focus",
    "demo": "image-plus-text",
    "agitation": "big-number",
}


def _fallback_layout(goal: str) -> str:
    return _LAYOUT_FALLBACK.get(goal, "content-with-title")


def _fallback_typography(slide_type: str) -> dict[str, Any]:
    return {
        "primary_size": 28,
        "secondary_size": 16,
        "weight_contrast": "600-400",
    }


def _fallback_color(emotion: str) -> dict[str, Any]:
    return {
        "background": "default",
        "text_color": "foreground",
        "accent_usage": "moderate",
    }
