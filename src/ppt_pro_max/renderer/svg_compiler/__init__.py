"""SVGCompiler — compile SVG subset to native editable PPTX shapes.

Usage::

    from ppt_pro_max.renderer.svg_compiler import SVGCompiler, SVGCompileError, SVGResult
    result = SVGCompiler(C=context).compile(svg_text, slide, rect)
"""
from ._affine import Affine, parse_transform
from ._compiler import SVGCompileError, SVGCompiler, SVGResult
from ._dash import StrokeStyle, apply_stroke_style, parse_stroke_style
from ._paint import (
    GradientDef,
    apply_gradient,
    collect_linear_gradient,
    collect_radial_gradient,
    resolve_paint,
)
from ._path import arc_to_cubics, parse_path, to_beziers
from ._sanitizer import sanitize
from ._text import render_svg_text

__all__ = [
    "Affine",
    "GradientDef",
    "SVGCompileError",
    "SVGCompiler",
    "SVGResult",
    "StrokeStyle",
    "apply_gradient",
    "apply_stroke_style",
    "arc_to_cubics",
    "collect_linear_gradient",
    "collect_radial_gradient",
    "parse_path",
    "parse_stroke_style",
    "parse_transform",
    "render_svg_text",
    "resolve_paint",
    "sanitize",
    "to_beziers",
]
