"""Louvre Abu Dhabi — hero cover v2.

One photographic anchor, one typographic anchor, and a receding editable SVG
field. No title card and no competing numeric callout.
"""
from pathlib import Path

from pptx_designer import Presentation, svg_chart
from pptx_designer.compiler import SVGCompileError
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.shapes import rect
from pptx_designer.tools.text import text

ROOT = Path(__file__).resolve().parent
IMAGE = ROOT / "assets" / "images" / "louvre_abudhabi_night_metalocus.jpg"
OUT = ROOT / "output" / "louvre_abudhabi_cover_hero_v3.pptx"

C = {
    "night": "#11242E", "white": "#F6F6F1", "silver": "#AEBCC0",
    "gold": "#B18C58", "faint": "#49636C", "faint_gold": "#756545",
}

STAR_FIELD = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 780 420">
  <g stroke="#49636C" stroke-width="1.2">
    <line x1="380" y1="24" x2="535" y2="88"/><line x1="535" y1="88" x2="600" y2="240"/>
    <line x1="600" y1="240" x2="535" y2="392"/><line x1="535" y1="392" x2="380" y2="456"/>
    <line x1="380" y1="456" x2="225" y2="392"/><line x1="225" y1="392" x2="160" y2="240"/>
    <line x1="160" y1="240" x2="225" y2="88"/><line x1="225" y1="88" x2="380" y2="24"/>
    <line x1="380" y1="78" x2="498" y2="127"/><line x1="498" y1="127" x2="547" y2="240"/>
    <line x1="547" y1="240" x2="498" y2="353"/><line x1="498" y1="353" x2="380" y2="402"/>
    <line x1="380" y1="402" x2="262" y2="353"/><line x1="262" y1="353" x2="213" y2="240"/>
    <line x1="213" y1="240" x2="262" y2="127"/><line x1="262" y1="127" x2="380" y2="78"/>
    <line x1="380" y1="24" x2="380" y2="456"/><line x1="160" y1="240" x2="600" y2="240"/>
  </g>
  <g stroke="#756545" stroke-width="1"><line x1="380" y1="0" x2="380" y2="480"/><line x1="110" y1="240" x2="650" y2="240"/></g>
  <g fill="#756545"><circle cx="380" cy="24" r="4"/><circle cx="600" cy="240" r="3.5"/><circle cx="380" cy="456" r="4"/></g>
</svg>'''


def tx(slide, x, y, w, h, value, size, color, font, bold=False, align=None):
    kw = dict(font_size=size, color=color, font_name=font, bold=bold, C=C)
    if align:
        kw["align"] = align
    return text(slide, x, y, w, h, value, **kw)


def main():
    if not IMAGE.exists():
        raise FileNotFoundError(IMAGE)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["night"], C=C)
    # A deliberately short, wide image field crops away most of the water and
    # enlarges the dome into the cover's sole photographic focal point.
    cover_image(slide, 0, 2.42, 13.333, 4.18, str(IMAGE))
    rect(slide, 0.66, 0.55, 0.72, 0.024, fill=C["gold"], C=C)
    tx(slide, 0.66, 0.82, 3.35, 0.12, "ARCHITECTURAL DOSSIER  /  2026", 7.0, "silver", "Arial Narrow", True)
    tx(slide, 0.66, 1.05, 5.30, 1.14, "LOUVRE\nABU DHABI", 40.0, "white", "Georgia")
    tx(slide, 5.12, 2.15, 2.10, 0.11, "JEAN NOUVEL  /  2017", 6.8, "silver", "Arial Narrow", True)
    try:
        result = svg_chart(slide, STAR_FIELD, x=8.12, y=0.28, w=4.22, h=2.04, C=C)
    except SVGCompileError as exc:
        raise RuntimeError(f"Star field compilation failed: {exc}") from exc
    rect(slide, 0.66, 6.83, 12.01, 0.022, fill=C["gold"], C=C)
    tx(slide, 0.70, 7.05, 3.45, 0.12, "SAADIYAT ISLAND  /  ABU DHABI, UAE", 6.7, "silver", "Arial Narrow", True)
    tx(slide, 11.98, 6.94, 0.55, 0.32, "01", 20.0, "silver", "Georgia", False, "right")
    prs.save(str(OUT))
    print(OUT)
    print(f"svg_shapes={result.shape_count}; svg_warnings={result.warnings}")


if __name__ == "__main__":
    main()
