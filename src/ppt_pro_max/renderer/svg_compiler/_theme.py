"""SVG theme bridge — C dict → SVG defaults, mood→gradient presets.

Bridges the build_helpers C color dictionary into SVG-compatible defaults
and provides mood-aware gradient presets for SVGCompiler.

Usage::

    from ppt_pro_max.renderer.svg_compiler._theme import svg_defaults, mood_gradient

    # Get SVG default fill/stroke from C dict
    defaults = svg_defaults(C)
    # → {"fill": "#1D78FA", "stroke": "#0A1E3D", "stroke-width": "1", ...}

    # Get a mood-appropriate gradient SVG snippet
    grad = mood_gradient("tech", C)
    # → '<defs><linearGradient id="mood-grad" ...>...</linearGradient></defs>'
"""
from __future__ import annotations

from ._compiler import _resolve_svg_color
from ._paint import GradientDef


def svg_defaults(C: dict | None = None) -> dict[str, str]:
    """Derive SVG default attribute values from a C (context) dict.

    Returns a dict of SVG attribute defaults that can be applied to the
    root <svg> element or used as fallbacks when elements omit fill/stroke.

    Mapping:
      fill        → C["primary"] (or "#000000")
      stroke      → C["foreground"] (or C["text_dark"], or "#000000")
      stroke-width→ "1"
      font-family → C["font_body"] (or "Arial")
      font-size   → "14"
      color       → C["text_dark"] (or C["foreground"], or "#000000")
      opacity     → "1"
    """
    C = C or {}
    return {
        "fill": _safe_resolve(C.get("primary", "#000000"), C),
        "stroke": _safe_resolve(C.get("foreground", C.get("text_dark", "#000000")), C),
        "stroke-width": "1",
        "font-family": C.get("font_body", "Arial"),
        "font-size": "14",
        "color": _safe_resolve(C.get("text_dark", C.get("foreground", "#000000")), C),
        "opacity": "1",
    }


def _safe_resolve(val: str, C: dict) -> str:
    try:
        r = _resolve_svg_color(val, C, "#000000")
        return r if r is not None else val
    except (ValueError, KeyError, TypeError):
        return val


_MOOD_GRADIENT_PRESETS: dict[str, dict] = {
    "tech": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "100%",
        "stops": [
            (0, "primary", 1.0),
            (0.5, "accent", 0.9),
            (1, "secondary", 0.8),
        ],
    },
    "dark": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "0%", "y2": "100%",
        "stops": [
            (0, "foreground", 1.0),
            (1, "muted", 0.6),
        ],
    },
    "warm": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 1.0),
            (0.6, "accent", 1.0),
            (1, "secondary", 0.9),
        ],
    },
    "elegant": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "100%",
        "stops": [
            (0, "primary", 1.0),
            (1, "accent", 0.85),
        ],
    },
    "luxury": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "100%",
        "stops": [
            (0, "primary", 1.0),
            (0.5, "accent", 0.95),
            (1, "primary", 0.7),
        ],
    },
    "vibrant": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "accent", 1.0),
            (0.5, "primary", 1.0),
            (1, "secondary", 1.0),
        ],
    },
    "startup": {
        "type": "linear",
        "x1": "0%", "y1": "100%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 1.0),
            (1, "accent", 1.0),
        ],
    },
    "nature": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "0%", "y2": "100%",
        "stops": [
            (0, "primary", 1.0),
            (1, "muted", 0.8),
        ],
    },
    "calm": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 0.9),
            (1, "muted", 0.7),
        ],
    },
    "minimal": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 0.8),
            (1, "border", 0.5),
        ],
    },
    "bold": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "100%",
        "stops": [
            (0, "primary", 1.0),
            (1, "destructive", 0.9),
        ],
    },
    "fintech": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 1.0),
            (0.7, "accent", 0.9),
            (1, "secondary", 0.8),
        ],
    },
    "neon": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "accent", 1.0),
            (0.5, "primary", 1.0),
            (1, "destructive", 0.9),
        ],
    },
    "ink-wash": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "0%", "y2": "100%",
        "stops": [
            (0, "foreground", 0.9),
            (1, "muted", 0.4),
        ],
    },
    "professional": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 1.0),
            (1, "secondary", 0.7),
        ],
    },
    "creative": {
        "type": "linear",
        "x1": "0%", "y1": "100%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "accent", 1.0),
            (0.5, "primary", 0.9),
            (1, "secondary", 0.8),
        ],
    },
    "mckinsey": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 1.0),
            (1, "accent", 0.85),
        ],
    },
    "consulting": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 1.0),
            (1, "accent", 0.85),
        ],
    },
    "retro": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "100%",
        "stops": [
            (0, "primary", 0.9),
            (0.5, "accent", 0.85),
            (1, "secondary", 0.7),
        ],
    },
    "pastel": {
        "type": "linear",
        "x1": "0%", "y1": "0%", "x2": "100%", "y2": "0%",
        "stops": [
            (0, "primary", 0.7),
            (1, "accent", 0.6),
        ],
    },
}


def mood_gradient(mood: str, C: dict | None = None, grad_id: str = "mood-grad") -> str:
    """Generate an SVG gradient <defs> snippet for a given mood.

    Returns a string like:
      '<defs><linearGradient id="mood-grad" x1="0%" y1="0%" ...>
        <stop offset="0" stop-color="#1D78FA" stop-opacity="1"/>
        ...
      </linearGradient></defs>'

    If the mood is not found, falls back to "professional".
    """
    C = C or {}
    preset = _MOOD_GRADIENT_PRESETS.get(mood, _MOOD_GRADIENT_PRESETS["professional"])
    gtype = preset["type"]

    if gtype == "radial":
        return _build_radial_svg(preset, C, grad_id)
    return _build_linear_svg(preset, C, grad_id)


def mood_gradient_def(mood: str, C: dict | None = None) -> GradientDef:
    """Return a GradientDef for a given mood (for programmatic use).

    This is the non-SVG-string version — returns a GradientDef that
    SVGCompiler can use directly via apply_gradient().
    """
    C = C or {}
    preset = _MOOD_GRADIENT_PRESETS.get(mood, _MOOD_GRADIENT_PRESETS["professional"])

    stops: list[tuple[float, str, float]] = []
    for pos, color_key, opacity in preset["stops"]:
        resolved = _safe_resolve(C.get(color_key, "#000000"), C)
        stops.append((pos, resolved, opacity))

    if preset["type"] == "radial":
        return GradientDef(
            stops=stops,
            cx=0.5, cy=0.5, r=0.5,
            gradient_type="radial",
        )

    x1 = _parse_pct(preset.get("x1", "0%"))
    y1 = _parse_pct(preset.get("y1", "0%"))
    x2 = _parse_pct(preset.get("x2", "100%"))
    y2 = _parse_pct(preset.get("y2", "0%"))
    return GradientDef(
        stops=stops,
        x1=x1, y1=y1, x2=x2, y2=y2,
        gradient_type="linear",
    )


def _parse_pct(v: str) -> float:
    if v.endswith("%"):
        return float(v.rstrip("%")) / 100.0
    return float(v)


def _build_linear_svg(preset: dict, C: dict, grad_id: str) -> str:
    attrs = [
        f'id="{grad_id}"',
        f'x1="{preset.get("x1", "0%")}"',
        f'y1="{preset.get("y1", "0%")}"',
        f'x2="{preset.get("x2", "100%")}"',
        f'y2="{preset.get("y2", "0%")}"',
    ]
    stops_xml = ""
    for pos, color_key, opacity in preset["stops"]:
        resolved = _safe_resolve(C.get(color_key, "#000000"), C)
        stops_xml += (
            f'<stop offset="{pos}" stop-color="{resolved}"'
            f' stop-opacity="{opacity}"/>\n'
        )
    return (
        f'<defs><linearGradient {" ".join(attrs)}>\n'
        f'{stops_xml}</linearGradient></defs>'
    )


def _build_radial_svg(preset: dict, C: dict, grad_id: str) -> str:
    attrs = [
        f'id="{grad_id}"',
        f'cx="{preset.get("cx", "50%")}"',
        f'cy="{preset.get("cy", "50%")}"',
        f'r="{preset.get("r", "50%")}"',
    ]
    stops_xml = ""
    for pos, color_key, opacity in preset["stops"]:
        resolved = _safe_resolve(C.get(color_key, "#000000"), C)
        stops_xml += (
            f'<stop offset="{pos}" stop-color="{resolved}"'
            f' stop-opacity="{opacity}"/>\n'
        )
    return (
        f'<defs><radialGradient {" ".join(attrs)}>\n'
        f'{stops_xml}</radialGradient></defs>'
    )


def c_to_svg_style(C: dict | None = None) -> str:
    """Generate an inline CSS style string from C dict for SVG <style> injection.

    Maps C keys to CSS custom properties so SVG can use var(--primary), etc.
    Example output: ':root{--primary:#1D78FA;--accent:#FF5500;--text-dark:#0A1E3D;}'
    """
    C = C or {}
    _KEY_MAP = {
        "primary": "primary",
        "on-primary": "on-primary",
        "secondary": "secondary",
        "accent": "accent",
        "background": "background",
        "foreground": "foreground",
        "muted": "muted",
        "muted-foreground": "muted-foreground",
        "border": "border",
        "destructive": "destructive",
        "text_dark": "text-dark",
        "text_body": "text-body",
        "text_muted": "text-muted",
        "card_bg": "card-bg",
        "card_line": "card-line",
        "divider": "divider",
        "white": "white",
        "light": "light",
        "bg_tint": "bg-tint",
    }
    parts = []
    for c_key, css_name in _KEY_MAP.items():
        val = C.get(c_key)
        if val:
            resolved = _safe_resolve(val, C)
            parts.append(f"--{css_name}:{resolved}")
    if not parts:
        return ""
    return f":root{{{';'.join(parts)};}}"


def available_mood_gradients() -> list[str]:
    return list(_MOOD_GRADIENT_PRESETS.keys())
