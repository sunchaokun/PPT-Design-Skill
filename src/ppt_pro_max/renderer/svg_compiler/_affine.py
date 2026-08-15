"""SVG affine transform helpers (translate / scale / rotate / matrix).

Pure-Python, no external dependencies. Translated from probe with type annotations.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass


@dataclass
class Affine:
    """2D affine matrix | a c e |
                        | b d f |
                        | 0 0 1 |
    """

    a: float = 1.0
    b: float = 0.0
    c: float = 0.0
    d: float = 1.0
    e: float = 0.0
    f: float = 0.0

    def compose(self, o: Affine) -> Affine:
        return Affine(
            self.a * o.a + self.b * o.c,
            self.a * o.b + self.b * o.d,
            self.c * o.a + self.d * o.c,
            self.c * o.b + self.d * o.d,
            self.e * o.a + self.f * o.c + o.e,
            self.e * o.b + self.f * o.d + o.f,
        )

    def apply(self, x: float, y: float) -> tuple[float, float]:
        return (self.a * x + self.c * y + self.e, self.b * x + self.d * y + self.f)


_TRANSFORM_RE = re.compile(r"([a-zA-Z]+)\(([^)]*)\)")


def parse_transform(s: str | None) -> Affine:
    """Parse an SVG `transform` attribute string into an ``Affine``.

    Supports ``translate``, ``scale``, ``rotate``, ``matrix``, ``skewX``, ``skewY``.
    Unknown ops are silently treated as identity (best-effort).
    """
    if not s:
        return Affine()
    out = Affine()
    for m in _TRANSFORM_RE.finditer(s):
        op = m.group(1)
        raw = m.group(2).replace(",", " ").split()
        args = [float(v) for v in raw if v != ""]
        if op == "translate":
            tx = args[0] if len(args) > 0 else 0.0
            ty = args[1] if len(args) > 1 else 0.0
            t = Affine(1, 0, 0, 1, tx, ty)
        elif op == "scale":
            sx = args[0] if len(args) > 0 else 1.0
            sy = args[1] if len(args) > 1 else sx
            t = Affine(sx, 0, 0, sy, 0, 0)
        elif op == "rotate":
            ang = math.radians(args[0])
            cx = args[1] if len(args) > 1 else 0.0
            cy = args[2] if len(args) > 2 else 0.0
            t = (
                Affine(1, 0, 0, 1, cx, cy)
                .compose(Affine(math.cos(ang), math.sin(ang), -math.sin(ang), math.cos(ang), 0, 0))
                .compose(Affine(1, 0, 0, 1, -cx, -cy))
            )
        elif op == "matrix":
            t = Affine(*args[:6])
        elif op == "skewX":
            t = Affine(1, 0, math.tan(math.radians(args[0])), 1, 0, 0)
        elif op == "skewY":
            t = Affine(1, math.tan(math.radians(args[0])), 0, 1, 0, 0)
        else:
            t = Affine()
        out = out.compose(t)
    return out
