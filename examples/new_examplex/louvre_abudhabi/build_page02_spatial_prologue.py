"""Louvre Abu Dhabi — P2, the spatial prologue.

Editorial architecture layout: narrative on a restrained paper column;
the unaltered interior photograph carries the experiential proof.
"""
from pathlib import Path

from pptx_designer import Presentation, svg_chart
from pptx_designer.compiler import SVGCompileError
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.shapes import rect
from pptx_designer.tools.text import text


ROOT = Path(__file__).resolve().parent
IMAGE = ROOT / "assets" / "images" / "louvre_abudhabi_rain_of_light.jpg"
OUT = ROOT / "output" / "louvre_abudhabi_p2_rain_of_light_v1.pptx"

C = {
    "paper": "#F4F3ED",
    "ink": "#15232D",
    "muted": "#596973",
    "pale": "#B7C1C2",
    "gold": "#B18C58",
    "white": "#F6F6F1",
}

# A delicate, editable trace of daylight travelling through the dome. It is a
# page component, not a second illustration: intentionally low contrast.
LIGHT_TRACE = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 180">
  <g stroke="#9DA9AB" stroke-width="1.1">
    <line x1="18" y1="24" x2="210" y2="24"/>
    <line x1="42" y1="50" x2="188" y2="50"/>
    <line x1="66" y1="76" x2="166" y2="76"/>
    <line x1="90" y1="102" x2="142" y2="102"/>
    <line x1="18" y1="24" x2="90" y2="158"/>
    <line x1="210" y1="24" x2="142" y2="158"/>
  </g>
  <g fill="#B18C58">
    <circle cx="18" cy="24" r="3.4"/><circle cx="210" cy="24" r="3.4"/>
  </g>
  <g fill="#596973">
    <circle cx="90" cy="158" r="2.7"/><circle cx="142" cy="158" r="2.7"/>
  </g>
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

    # The image is deliberately a clean photographic field, without darkened
    # overlays or synthetic treatment. cover_image preserves aspect ratio.
    cover_image(slide, 3.24, 0, 10.093, 7.5, str(IMAGE))
    rect(slide, 0, 0, 3.24, 7.5, fill=C["paper"], C=C)
    rect(slide, 3.17, 0, 0.024, 7.5, fill=C["gold"], C=C)

    tx(slide, 0.63, 0.66, 1.95, 0.14, "01  /  SPATIAL PROLOGUE", 7.0,
       "muted", "Arial Narrow", True)
    rect(slide, 0.63, 1.24, 0.72, 0.023, fill=C["gold"], C=C)
    tx(slide, 0.63, 1.62, 2.16, 1.20, "RAIN\nOF LIGHT", 30.0,
       "ink", "Georgia")
    tx(slide, 0.66, 3.20, 2.12, 0.58,
       "BENEATH A POROUS DOME,\nDAYLIGHT BECOMES A\nCULTURAL MEDIUM.",
       9.0, "ink", "Arial Narrow", True)
    tx(slide, 0.66, 4.08, 2.12, 0.42,
       "The roof does not simply shelter the museum;\nit composes its atmosphere.",
       7.6, "muted", "Arial")

    try:
        result = svg_chart(slide, LIGHT_TRACE, x=0.62, y=4.86, w=2.02, h=1.52, C=C)
    except SVGCompileError as exc:
        raise RuntimeError(f"Light trace SVG failed to compile: {exc}") from exc

    tx(slide, 0.63, 6.63, 1.9, 0.12, "LOUVRE ABU DHABI", 6.6,
       "muted", "Arial Narrow", True)
    tx(slide, 0.63, 6.89, 1.9, 0.12, "JEAN NOUVEL  /  2017", 6.6,
       "muted", "Arial Narrow", True)
    tx(slide, 2.25, 6.51, 0.52, 0.45, "02", 28.0,
       "pale", "Georgia", False, "right")

    prs.save(str(OUT))
    print(OUT)
    print(f"svg_shapes={result.shape_count}; svg_warnings={result.warnings}")


if __name__ == "__main__":
    main()
