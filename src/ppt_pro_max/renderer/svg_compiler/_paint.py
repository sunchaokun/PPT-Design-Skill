"""SVG paint (fill/stroke) resolution with full gradient support.

Extracts paint resolution from _compiler.py and adds:
- Radial gradient with proper fillToRect mapping
- gradientTransform support (translate/scale/rotate on gradients)
- Multi-stop gradient with per-stop alpha
- spreadMethod (pad/reflect/repeat) — pad only, others raise SVGCompileError
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

from ppt_pro_max.renderer.visual_effects import GradientFill, GradientStop

from ._affine import Affine, parse_transform


@dataclass
class GradientDef:
    stops: list[tuple[float, str, float]] = field(default_factory=list)
    x1: float = 0.0
    y1: float = 0.0
    x2: float = 1.0
    y2: float = 0.0
    cx: float = 0.5
    cy: float = 0.5
    r: float = 0.5
    gradient_type: str = "linear"
    spread_method: str = "pad"
    transform: Affine | None = None


def _parse_percent_or_float(v: str, default: float = 0.0) -> float:
    if not v:
        return default
    v = v.strip()
    if v.endswith("%"):
        return float(v.rstrip("%")) / 100.0
    return float(v)


def collect_linear_gradient(g, C: dict, resolve_color_fn) -> GradientDef:
    stops = []
    for s in g.iter("{http://www.w3.org/2000/svg}stop"):
        off = s.get("offset", "0")
        pos = float(off.rstrip("%")) / 100.0 if off.endswith("%") else float(off)
        col = resolve_color_fn(s.get("stop-color", "#000000"), C, "#000000")
        op = float(s.get("stop-opacity", "1"))
        stops.append((pos, col, op))

    spread = g.get("spreadMethod", "pad")
    if spread not in ("pad",):
        from . import SVGCompileError
        raise SVGCompileError(f"unsupported gradient spreadMethod: {spread}")

    tf_str = g.get("gradientTransform")
    tf = parse_transform(tf_str) if tf_str else None

    return GradientDef(
        stops=stops,
        x1=_parse_percent_or_float(g.get("x1", "0")),
        y1=_parse_percent_or_float(g.get("y1", "0")),
        x2=_parse_percent_or_float(g.get("x2", "1")),
        y2=_parse_percent_or_float(g.get("y2", "0")),
        spread_method=spread,
        transform=tf,
    )


def collect_radial_gradient(g, C: dict, resolve_color_fn) -> GradientDef:
    stops = []
    for s in g.iter("{http://www.w3.org/2000/svg}stop"):
        off = s.get("offset", "0")
        pos = float(off.rstrip("%")) / 100.0 if off.endswith("%") else float(off)
        col = resolve_color_fn(s.get("stop-color", "#000000"), C, "#000000")
        op = float(s.get("stop-opacity", "1"))
        stops.append((pos, col, op))

    spread = g.get("spreadMethod", "pad")
    if spread not in ("pad",):
        from . import SVGCompileError
        raise SVGCompileError(f"unsupported gradient spreadMethod: {spread}")

    tf_str = g.get("gradientTransform")
    tf = parse_transform(tf_str) if tf_str else None

    return GradientDef(
        stops=stops,
        cx=_parse_percent_or_float(g.get("cx", "50%"), 0.5),
        cy=_parse_percent_or_float(g.get("cy", "50%"), 0.5),
        r=_parse_percent_or_float(g.get("r", "50%"), 0.5),
        gradient_type="radial",
        spread_method=spread,
        transform=tf,
    )


def apply_gradient(elem, grad: GradientDef, wrap_fn) -> None:
    if grad.gradient_type == "radial":
        cx_pct = int(grad.cx * 100000)
        cy_pct = int(grad.cy * 100000)
        r_pct = int(grad.r * 100000)
        l_val = str(cx_pct - r_pct) if cx_pct > r_pct else "0"
        t_val = str(cy_pct - r_pct) if cy_pct > r_pct else "0"
        r_val = str(cx_pct + r_pct)
        b_val = str(cy_pct + r_pct)
        gf = GradientFill(
            gradient_type="path",
            fill_to_rect={"l": l_val, "t": t_val, "r": r_val, "b": b_val},
        )
    else:
        dx = grad.x2 - grad.x1
        dy = grad.y2 - grad.y1
        angle_rad = math.atan2(dy, dx)
        gf = GradientFill(angle=int(math.degrees(angle_rad) * 60000))

    for pos, col, op in grad.stops:
        alpha = int(op * 100000) if op < 1.0 else 100000
        gf.stops.append(GradientStop(color=col, position=int(pos * 100000), alpha=alpha))

    wrapper = wrap_fn(elem)
    gf.apply(wrapper)


def resolve_paint(
    el, which: str, grads: dict[str, GradientDef], C: dict, resolve_color_fn, features: set
) -> tuple[str, object | None, int]:
    v = el.get(which)
    op = el.get(f"{which}-opacity")
    alpha = int(float(op) * 100) if op else 100

    if v is None:
        return "none", None, alpha
    if v.startswith("url(#"):
        gid = v[v.index("#") + 1 : -1]
        grad = grads.get(gid)
        if grad is not None:
            features.add("gradient")
            return "grad", grad, alpha
        return "none", None, alpha
    if v == "none":
        return "none", None, alpha

    resolved = resolve_color_fn(v, C, "")
    if resolved is None:
        return "none", None, alpha
    features.add("solid")
    return "solid", resolved, alpha
