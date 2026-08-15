"""SVG text rendering with baseline-aware positioning.

Extracts text rendering from _compiler.py and adds:
- Pillow-based text measurement (with graceful fallback to estimate_text_size)
- SVG baseline (dominant-baseline, alignment-baseline) → PPT vertical anchor mapping
- Multi-line <tspan> support
- font-family → PPT font mapping
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from ppt_pro_max.renderer.text_measurer import estimate_text_size

from ._affine import Affine

_FONT_MAP: dict[str, str] = {
    "arial": "Arial",
    "helvetica": "Arial",
    "sans-serif": "Arial",
    "times new roman": "Times New Roman",
    "times": "Times New Roman",
    "serif": "Times New Roman",
    "georgia": "Georgia",
    "courier": "Courier New",
    "courier new": "Courier New",
    "monospace": "Courier New",
    "consolas": "Consolas",
    "verdana": "Verdana",
    "tahoma": "Tahoma",
    "calibri": "Calibri",
    "impact": "Impact",
    "comic sans ms": "Comic Sans MS",
}

_BASELINE_MAP: dict[str, MSO_ANCHOR] = {
    "auto": MSO_ANCHOR.MIDDLE,
    "alphabetic": MSO_ANCHOR.BOTTOM,
    "text-after-edge": MSO_ANCHOR.BOTTOM,
    "text-before-edge": MSO_ANCHOR.TOP,
    "central": MSO_ANCHOR.MIDDLE,
    "middle": MSO_ANCHOR.MIDDLE,
    "hanging": MSO_ANCHOR.TOP,
    "mathematical": MSO_ANCHOR.MIDDLE,
}

_ANCHOR_MAP: dict[str, PP_ALIGN] = {
    "middle": PP_ALIGN.CENTER,
    "end": PP_ALIGN.RIGHT,
    "start": PP_ALIGN.LEFT,
}


@dataclass
class TextMetrics:
    width_inches: float
    height_inches: float
    ascent_ratio: float = 0.8
    descent_ratio: float = 0.2


def _resolve_font_family(raw: str | None) -> str:
    if not raw:
        return "Calibri"
    families = re.split(r"[,\s]+", raw.strip("'\""))
    for f in families:
        key = f.strip().lower()
        if key in _FONT_MAP:
            return _FONT_MAP[key]
    return families[0].strip().strip("'\"")


def _resolve_baseline(el) -> MSO_ANCHOR:
    db = el.get("dominant-baseline", el.get("alignment-baseline", "auto"))
    return _BASELINE_MAP.get(db.lower(), MSO_ANCHOR.MIDDLE)


def _measure_text(
    content: str, font_size_pt: float, font_family: str, max_width_inches: float
) -> TextMetrics:
    try:
        from PIL import ImageFont

        try:
            font = ImageFont.truetype(font_family + ".ttf", int(font_size_pt))
        except OSError:
            font = ImageFont.load_default()

        bbox = font.getbbox(content)
        w_px = bbox[2] - bbox[0]
        h_px = bbox[3] - bbox[1]
        ascent = font.getmetrics()[0]
        total_h = font.getmetrics()[0] + font.getmetrics()[1]

        px_per_inch = 96.0
        w_in = w_px / px_per_inch
        h_in = h_px / px_per_inch

        return TextMetrics(
            width_inches=max(w_in, 0.5),
            height_inches=max(h_in, 0.3),
            ascent_ratio=ascent / total_h if total_h > 0 else 0.8,
            descent_ratio=1.0 - (ascent / total_h if total_h > 0 else 0.8),
        )
    except ImportError:
        _, h_est = estimate_text_size(content, max(8, int(font_size_pt)), max_width_inches)
        return TextMetrics(
            width_inches=max_width_inches,
            height_inches=h_est,
            ascent_ratio=0.8,
            descent_ratio=0.2,
        )


def render_svg_text(
    el,
    tf: Affine,
    to_inches_fn,
    slide,
    C: dict,
    resolve_color_fn,
    features: set,
) -> None:
    features.add("text")

    x = float(el.get("x", 0))
    y = float(el.get("y", 0))
    ix, iy = to_inches_fn(*tf.apply(x, y))

    content = "".join(el.itertext()).strip()
    if not content:
        return

    fs = float(re.sub(r"[^\d.]", "", el.get("font-size", "14")))
    anchor = el.get("text-anchor", "start")
    font_family = _resolve_font_family(el.get("font-family"))
    v_anchor = _resolve_baseline(el)

    fval = el.get("fill", "#000000")
    if fval and fval.startswith("url(#"):
        fval = "#000000"
    elif fval and fval != "none":
        resolved = resolve_color_fn(fval, C, "")
        if resolved:
            fval = resolved
        else:
            fval = "#000000"
    else:
        fval = "#000000"

    metrics = _measure_text(content, fs, font_family, 8.0)

    if anchor == "middle":
        left = ix - metrics.width_inches / 2
    elif anchor == "end":
        left = ix - metrics.width_inches
    else:
        left = ix - 0.1

    if v_anchor == MSO_ANCHOR.TOP:
        top = iy
    elif v_anchor == MSO_ANCHOR.BOTTOM:
        top = iy - metrics.height_inches
    else:
        top = iy - metrics.height_inches * metrics.ascent_ratio

    width = metrics.width_inches + 0.2
    height = metrics.height_inches * 1.5

    tb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf_el = tb.text_frame
    tf_el.word_wrap = False
    tf_el.vertical_anchor = v_anchor
    p = tf_el.paragraphs[0]
    p.alignment = _ANCHOR_MAP.get(anchor, PP_ALIGN.LEFT)
    run = p.add_run()
    run.text = content
    run.font.size = Pt(fs)
    run.font.color.rgb = RGBColor.from_string(fval.lstrip("#"))
    run.font.name = font_family

    bold = el.get("font-weight")
    if bold and bold not in ("normal", "100", "200", "300"):
        run.font.bold = True

    italic = el.get("font-style")
    if italic == "italic":
        run.font.italic = True
