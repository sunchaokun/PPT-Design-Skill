"""Louvre Abu Dhabi — flagship portfolio P3–P5.

P3 is a dark technical reading; P4 a panoramic aerial sequence; P5 a quiet
waterfront coda. All diagrams compile into editable PowerPoint shapes.
"""
from pathlib import Path

from pptx_designer import Presentation, svg_chart
from pptx_designer.compiler import SVGCompileError
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.shapes import rect
from pptx_designer.tools.text import text


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets" / "images"
OUT = ROOT / "output" / "louvre_abudhabi_p3_p5_v1.pptx"

C = {
    "night": "#10212B",
    "night2": "#17303B",
    "paper": "#F4F3ED",
    "white": "#F6F6F1",
    "silver": "#AEBCC0",
    "muted": "#576870",
    "gold": "#B18C58",
    "dark": "#16252D",
}

DOME_SECTION = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 980 560">
  <g stroke="#6F8991" stroke-width="1.3">
    <line x1="600" y1="42" x2="760" y2="108"/><line x1="760" y1="108" x2="826" y2="268"/>
    <line x1="826" y1="268" x2="760" y2="428"/><line x1="760" y1="428" x2="600" y2="494"/>
    <line x1="600" y1="494" x2="440" y2="428"/><line x1="440" y1="428" x2="374" y2="268"/>
    <line x1="374" y1="268" x2="440" y2="108"/><line x1="440" y1="108" x2="600" y2="42"/>
    <line x1="600" y1="92" x2="724" y2="144"/><line x1="724" y1="144" x2="774" y2="268"/>
    <line x1="774" y1="268" x2="724" y2="392"/><line x1="724" y1="392" x2="600" y2="444"/>
    <line x1="600" y1="444" x2="476" y2="392"/><line x1="476" y1="392" x2="426" y2="268"/>
    <line x1="426" y1="268" x2="476" y2="144"/><line x1="476" y1="144" x2="600" y2="92"/>
    <line x1="600" y1="42" x2="600" y2="494"/><line x1="374" y1="268" x2="826" y2="268"/>
    <line x1="440" y1="108" x2="760" y2="428"/><line x1="760" y1="108" x2="440" y2="428"/>
  </g>
  <g stroke="#B18C58" stroke-width="1.0">
    <line x1="600" y1="12" x2="600" y2="530"/>
    <line x1="340" y1="268" x2="860" y2="268"/>
  </g>
  <g fill="#B18C58">
    <circle cx="600" cy="42" r="5"/><circle cx="760" cy="108" r="4"/>
    <circle cx="826" cy="268" r="4"/><circle cx="600" cy="494" r="5"/>
  </g>
  <g fill="#AEBCC0">
    <circle cx="374" cy="268" r="4"/><circle cx="440" cy="428" r="4"/>
    <circle cx="600" cy="268" r="4"/><circle cx="724" cy="392" r="3"/>
  </g>
</svg>'''

WATER_TRACE = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 500 180">
  <g stroke="#AEBCC0" stroke-width="1.1">
    <line x1="18" y1="26" x2="482" y2="26"/><line x1="72" y1="74" x2="428" y2="74"/>
    <line x1="128" y1="122" x2="372" y2="122"/>
  </g>
  <g stroke="#B18C58" stroke-width="1.1">
    <line x1="18" y1="26" x2="128" y2="164"/><line x1="482" y1="26" x2="372" y2="164"/>
  </g>
  <g fill="#B18C58"><circle cx="18" cy="26" r="4"/><circle cx="482" cy="26" r="4"/></g>
  <g fill="#576870"><circle cx="128" cy="164" r="3"/><circle cx="372" cy="164" r="3"/></g>
</svg>'''


def tx(slide, x, y, w, h, value, size, color, font, bold=False, align=None):
    kw = dict(font_size=size, color=color, font_name=font, bold=bold, C=C)
    if align:
        kw["align"] = align
    return text(slide, x, y, w, h, value, **kw)


def svg(slide, markup, x, y, w, h):
    try:
        return svg_chart(slide, markup, x=x, y=y, w=w, h=h, C=C)
    except SVGCompileError as exc:
        raise RuntimeError(f"SVG compilation failed: {exc}") from exc


def slide_p3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["night"], C=C)
    rect(slide, 0.66, 0.63, 0.74, 0.025, fill=C["gold"], C=C)
    tx(slide, 0.66, 0.83, 2.8, 0.14, "02  /  GEOMETRY AS CLIMATE", 7.0,
       "silver", "Arial Narrow", True)
    tx(slide, 0.66, 1.42, 4.5, 1.28, "THE DOME\nAS INSTRUMENT", 31.0,
       "white", "Georgia")
    tx(slide, 0.70, 3.22, 3.65, 0.52,
       "Four outer layers of stainless steel\nmeet four inner layers of aluminium.",
       9.0, "white", "Arial Narrow", True)
    tx(slide, 0.70, 4.12, 3.35, 0.58,
       "A geometric canopy filters daylight, shades\nthe plaza and generates a microclimate below.",
       7.8, "silver", "Arial")
    tx(slide, 0.70, 6.42, 4.2, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.2,
       "silver", "Arial Narrow", True)
    tx(slide, 0.70, 6.75, 1.25, 0.45, "03", 28.0, "#45616B", "Georgia")
    result = svg(slide, DOME_SECTION, 5.08, 0.63, 7.55, 4.76)
    tx(slide, 5.33, 5.54, 2.65, 0.54, "7,850", 37.0, "white", "Arial Narrow")
    tx(slide, 5.38, 6.14, 2.32, 0.12, "STARS IN THE PATTERN", 6.8, "gold", "Arial Narrow", True)
    tx(slide, 9.42, 5.54, 2.70, 0.54, "8", 37.0, "white", "Arial Narrow")
    tx(slide, 9.46, 6.14, 2.32, 0.12, "OVERLAPPING LAYERS", 6.8, "gold", "Arial Narrow", True)
    tx(slide, 5.33, 6.69, 2.65, 0.42, "180 M", 20.0, "silver", "Arial Narrow")
    tx(slide, 7.14, 6.80, 2.90, 0.12, "DIAMETER OF THE DOME", 6.8, "silver", "Arial Narrow", True)
    return result.shape_count


def slide_p4(prs):
    image = ASSETS / "louvre_abudhabi_aerial_expedia.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["paper"], C=C)
    tx(slide, 0.66, 0.45, 1.95, 0.14, "03  /  URBAN FIGURE", 7.0,
       "muted", "Arial Narrow", True)
    tx(slide, 0.66, 0.79, 7.72, 0.50, "A MUSEUM CITY  /  IN THE SEA", 26.0,
       "dark", "Georgia")
    tx(slide, 10.73, 0.89, 1.87, 0.12, "ABU DHABI · UAE", 6.6,
       "muted", "Arial Narrow", True, "right")
    rect(slide, 0.66, 1.42, 12.01, 0.022, fill=C["gold"], C=C)
    # Exact 2.4:1 photographic ratio; no image stretching or crop.
    cover_image(slide, 0, 1.66, 13.333, 5.555, str(image))
    rect(slide, 9.03, 5.76, 3.65, 1.15, fill=C["night"], C=C)
    tx(slide, 9.34, 5.98, 3.03, 0.30, "55 BUILDINGS", 17.0,
       "white", "Arial Narrow")
    tx(slide, 9.37, 6.44, 2.82, 0.12, "23 ARE GALLERIES", 6.8,
       "gold", "Arial Narrow", True)
    tx(slide, 0.67, 7.26, 0.55, 0.14, "04", 7.0, "muted", "Arial Narrow", True)


def slide_p5(prs):
    image = ASSETS / "louvre_abudhabi_water_gallery_mediaoffice.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["paper"], C=C)
    cover_image(slide, 0.66, 0.56, 4.38, 6.38, str(image))
    rect(slide, 5.33, 0.56, 0.024, 6.38, fill=C["gold"], C=C)
    tx(slide, 5.80, 0.69, 2.0, 0.12, "04  /  CIVIC INTERIOR", 7.0,
       "muted", "Arial Narrow", True)
    tx(slide, 5.80, 1.29, 6.18, 1.42, "WATER\nAS PUBLIC ROOM", 34.0,
       "dark", "Georgia")
    tx(slide, 5.84, 3.16, 5.26, 0.48,
       "Paths cross the water beneath the dome, turning the\nmuseum into a walkable, climate-softened urban room.",
       9.1, "dark", "Arial Narrow", True)
    tx(slide, 5.84, 4.14, 4.78, 0.35,
       "Architecture is not an object beside the sea.\nIt becomes a small city within it.",
       7.8, "muted", "Arial")
    result = svg(slide, WATER_TRACE, 5.80, 5.10, 5.84, 1.05)
    tx(slide, 5.83, 6.60, 2.0, 0.12, "LAND + SEA", 6.6,
       "gold", "Arial Narrow", True)
    tx(slide, 10.80, 6.36, 1.10, 0.46, "05", 28.0,
       "#B7C1C2", "Georgia", False, "right")
    tx(slide, 5.83, 6.92, 3.65, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0,
       "muted", "Arial Narrow", True)
    return result.shape_count


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    p3_shapes = slide_p3(prs)
    slide_p4(prs)
    p5_shapes = slide_p5(prs)
    prs.save(str(OUT))
    print(OUT)
    print(f"p3_svg_shapes={p3_shapes}; p5_svg_shapes={p5_shapes}")


if __name__ == "__main__":
    main()
