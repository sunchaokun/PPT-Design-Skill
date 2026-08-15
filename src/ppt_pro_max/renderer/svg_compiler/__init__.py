"""SVGCompiler — compile SVG subset to native editable PPTX shapes.

Usage::

    from ppt_pro_max.renderer.svg_compiler import SVGCompiler, SVGCompileError, SVGResult
    result = SVGCompiler(C=context).compile(svg_text, slide, rect)
"""
from ._affine import Affine, parse_transform
from ._compiler import SVGCompileError, SVGCompiler, SVGResult
from ._path import arc_to_cubics, parse_path, to_beziers
from ._sanitizer import sanitize

__all__ = [
    "Affine",
    "SVGCompileError",
    "SVGCompiler",
    "SVGResult",
    "arc_to_cubics",
    "parse_path",
    "parse_transform",
    "sanitize",
    "to_beziers",
]
