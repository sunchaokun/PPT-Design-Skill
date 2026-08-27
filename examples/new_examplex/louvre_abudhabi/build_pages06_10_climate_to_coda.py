"""Louvre Abu Dhabi — P6–P10 flagship continuation.

Five intentionally distinct editorial architectures. The SVG studies compile
as native editable PowerPoint shapes; photography is always cover-fitted.
"""
from pathlib import Path

from pptx_designer import Presentation, svg_chart
from pptx_designer.compiler import SVGCompileError
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.shapes import rect
from pptx_designer.tools.text import text


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets" / "images"
OUT = ROOT / "output" / "louvre_abudhabi_p6_p10_v1.pptx"

C = {
    "night": "#11242E", "night2": "#1E3843", "paper": "#F4F3ED",
    "white": "#F6F6F1", "silver": "#B3C1C4", "muted": "#586A73",
    "gold": "#B18C58", "dark": "#16252D", "pale": "#CCD4D3",
}

CLIMATE_FIELD = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 720 560">
  <g stroke="#7D9499" stroke-width="1.2">
    <line x1="360" y1="55" x2="540" y2="130"/><line x1="540" y1="130" x2="615" y2="280"/>
    <line x1="615" y1="280" x2="540" y2="430"/><line x1="540" y1="430" x2="360" y2="505"/>
    <line x1="360" y1="505" x2="180" y2="430"/><line x1="180" y1="430" x2="105" y2="280"/>
    <line x1="105" y1="280" x2="180" y2="130"/><line x1="180" y1="130" x2="360" y2="55"/>
    <line x1="360" y1="105" x2="505" y2="165"/><line x1="505" y1="165" x2="565" y2="280"/>
    <line x1="565" y1="280" x2="505" y2="395"/><line x1="505" y1="395" x2="360" y2="455"/>
    <line x1="360" y1="455" x2="215" y2="395"/><line x1="215" y1="395" x2="155" y2="280"/>
    <line x1="155" y1="280" x2="215" y2="165"/><line x1="215" y1="165" x2="360" y2="105"/>
    <line x1="360" y1="55" x2="360" y2="505"/><line x1="105" y1="280" x2="615" y2="280"/>
  </g>
  <g stroke="#B18C58" stroke-width="1.15">
    <line x1="360" y1="2" x2="360" y2="540"/><line x1="64" y1="280" x2="656" y2="280"/>
  </g>
  <g fill="#B18C58"><circle cx="360" cy="55" r="5"/><circle cx="615" cy="280" r="4"/><circle cx="360" cy="505" r="5"/></g>
  <g fill="#7D9499"><circle cx="105" cy="280" r="4"/><circle cx="360" cy="280" r="4"/><circle cx="505" cy="165" r="3"/></g>
</svg>'''

LIGHT_ROUTE = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 760 230">
  <g stroke="#6D858D" stroke-width="1.15">
    <line x1="20" y1="28" x2="740" y2="28"/><line x1="110" y1="102" x2="650" y2="102"/>
    <line x1="200" y1="176" x2="560" y2="176"/>
  </g>
  <g stroke="#B18C58" stroke-width="1.1"><line x1="20" y1="28" x2="200" y2="214"/><line x1="740" y1="28" x2="560" y2="214"/></g>
  <g fill="#B18C58"><circle cx="20" cy="28" r="4"/><circle cx="740" cy="28" r="4"/></g>
  <g fill="#6D858D"><circle cx="200" cy="214" r="3"/><circle cx="560" cy="214" r="3"/></g>
</svg>'''

STAR_FIELD = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 560">
  <g stroke="#38535D" stroke-width="1">
    <line x1="400" y1="28" x2="545" y2="88"/><line x1="545" y1="88" x2="605" y2="230"/>
    <line x1="605" y1="230" x2="545" y2="372"/><line x1="545" y1="372" x2="400" y2="432"/>
    <line x1="400" y1="432" x2="255" y2="372"/><line x1="255" y1="372" x2="195" y2="230"/>
    <line x1="195" y1="230" x2="255" y2="88"/><line x1="255" y1="88" x2="400" y2="28"/>
    <line x1="680" y1="128" x2="800" y2="178"/><line x1="800" y1="178" x2="850" y2="298"/>
    <line x1="850" y1="298" x2="800" y2="418"/><line x1="800" y1="418" x2="680" y2="468"/>
    <line x1="680" y1="468" x2="560" y2="418"/><line x1="560" y1="418" x2="510" y2="298"/>
    <line x1="510" y1="298" x2="560" y2="178"/><line x1="560" y1="178" x2="680" y2="128"/>
    <line x1="400" y1="28" x2="400" y2="432"/><line x1="195" y1="230" x2="605" y2="230"/>
  </g>
  <g fill="#536D75"><circle cx="400" cy="28" r="3"/><circle cx="605" cy="230" r="3"/><circle cx="680" cy="128" r="3"/><circle cx="850" cy="298" r="3"/></g>
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


def p6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["paper"], C=C)
    tx(slide, 0.66, 0.58, 2.55, 0.12, "05  /  THERMAL FIELD", 7.0, "muted", "Arial Narrow", True)
    rect(slide, 0.66, 1.03, 0.74, 0.023, fill=C["gold"], C=C)
    tx(slide, 0.66, 1.45, 4.08, 1.46, "SHADE\nBECOMES\nINFRASTRUCTURE", 30.0, "dark", "Georgia")
    tx(slide, 0.70, 3.55, 3.65, 0.54, "The canopy filters daylight, reduces heat\nand protects the public realm below.", 9.0, "dark", "Arial Narrow", True)
    tx(slide, 0.70, 4.46, 3.44, 0.43, "Natural cooling is not an add-on.\nIt is the spatial idea of the project.", 7.8, "muted", "Arial")
    result = svg(slide, CLIMATE_FIELD, 5.15, 0.62, 6.96, 4.96)
    rect(slide, 5.24, 5.91, 6.88, 0.023, fill=C["gold"], C=C)
    tx(slide, 5.25, 6.13, 1.88, 0.46, "36 M", 24.0, "dark", "Arial Narrow")
    tx(slide, 7.05, 6.31, 1.60, 0.12, "DOME ABOVE GROUND", 6.5, "muted", "Arial Narrow", True)
    tx(slide, 8.95, 6.13, 1.88, 0.46, "7,500 T", 24.0, "dark", "Arial Narrow")
    tx(slide, 10.78, 6.31, 1.16, 0.12, "DOME WEIGHT", 6.5, "muted", "Arial Narrow", True)
    tx(slide, 0.66, 6.91, 3.52, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0, "muted", "Arial Narrow", True)
    tx(slide, 12.00, 6.79, 0.50, 0.34, "06", 20.0, "pale", "Georgia", False, "right")
    return result.shape_count


def p7(prs):
    images = [
        ASSETS / "louvre_abudhabi_rain_of_light.jpg",
        ASSETS / "louvre_abudhabi_aerial_expedia.jpg",
        ASSETS / "louvre_abudhabi_water_gallery_mediaoffice.jpg",
    ]
    for image in images:
        if not image.exists():
            raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["night"], C=C)
    tx(slide, 0.66, 0.57, 2.4, 0.12, "06  /  SPATIAL SEQUENCE", 7.0, "silver", "Arial Narrow", True)
    tx(slide, 0.66, 1.09, 6.15, 0.55, "WALK  /  PAUSE  /  RETURN", 27.0, "white", "Georgia")
    tx(slide, 9.73, 1.27, 2.75, 0.13, "A MUSEUM DESIGNED AS A JOURNEY", 6.6, "gold", "Arial Narrow", True, "right")
    rect(slide, 0.66, 1.91, 12.00, 0.022, fill=C["gold"], C=C)
    cover_image(slide, 0.66, 2.30, 3.60, 3.74, str(images[0]))
    cover_image(slide, 4.50, 2.30, 3.60, 3.74, str(images[1]))
    cover_image(slide, 8.34, 2.30, 3.60, 3.74, str(images[2]))
    tx(slide, 0.70, 6.33, 2.65, 0.12, "01  /  FILTERED DAYLIGHT", 6.8, "silver", "Arial Narrow", True)
    tx(slide, 4.54, 6.33, 2.55, 0.12, "02  /  CITY + SEA", 6.8, "silver", "Arial Narrow", True)
    tx(slide, 8.38, 6.33, 2.75, 0.12, "03  /  CIVIC WATERROOM", 6.8, "silver", "Arial Narrow", True)
    tx(slide, 0.66, 6.93, 4.02, 0.12, "THREE ATMOSPHERES, ONE CONTINUOUS ROOF.", 6.5, "gold", "Arial Narrow", True)
    tx(slide, 12.01, 6.76, 0.50, 0.34, "07", 20.0, "#48626B", "Georgia", False, "right")


def p8(prs):
    image = ASSETS / "louvre_abudhabi_rain_of_light.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["paper"], C=C)
    cover_image(slide, 0, 0, 13.333, 4.37, str(image))
    rect(slide, 0, 4.37, 13.333, 3.13, fill=C["paper"], C=C)
    tx(slide, 0.66, 4.73, 2.6, 0.12, "07  /  MATERIAL LOGIC", 7.0, "muted", "Arial Narrow", True)
    tx(slide, 0.66, 5.23, 6.48, 0.54, "MATERIAL IS A FILTER", 29.0, "dark", "Georgia")
    tx(slide, 0.70, 6.09, 5.20, 0.40, "Steel, aluminium and patterned voids work as a single\noptical instrument between sky and ground.", 8.7, "dark", "Arial Narrow", True)
    rect(slide, 8.04, 4.73, 0.024, 1.98, fill=C["gold"], C=C)
    tx(slide, 8.36, 4.87, 3.50, 0.34, "8 LAYERS", 20.0, "dark", "Arial Narrow")
    tx(slide, 8.39, 5.34, 3.37, 0.12, "4 STAINLESS STEEL / 4 ALUMINIUM", 6.8, "muted", "Arial Narrow", True)
    tx(slide, 8.36, 5.90, 3.86, 0.28, "The roof permits daylight\nwithout excessive heat or wind.", 7.5, "muted", "Arial")
    tx(slide, 0.66, 7.08, 3.52, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0, "muted", "Arial Narrow", True)
    tx(slide, 12.01, 6.99, 0.50, 0.34, "08", 20.0, "pale", "Georgia", False, "right")


def p9(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["night"], C=C)
    result = svg(slide, STAR_FIELD, 4.68, 0.48, 7.84, 5.70)
    tx(slide, 0.66, 0.62, 2.65, 0.12, "08  /  ARCHITECTURAL POSITION", 7.0, "silver", "Arial Narrow", True)
    rect(slide, 0.66, 1.14, 0.74, 0.023, fill=C["gold"], C=C)
    tx(slide, 0.66, 1.67, 6.32, 1.90, "A WELCOMING WORLD\nOF LIGHT, SHADOW,\nREFLECTION AND CALM.", 31.0, "white", "Georgia")
    tx(slide, 0.70, 4.32, 3.7, 0.36, "A museum-city that belongs to its geography\nwithout becoming a literal interpretation of it.", 8.6, "silver", "Arial")
    tx(slide, 0.70, 5.68, 2.95, 0.12, "JEAN NOUVEL  /  ARCHITECT", 6.8, "gold", "Arial Narrow", True)
    tx(slide, 0.70, 6.93, 3.85, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0, "silver", "Arial Narrow", True)
    tx(slide, 12.00, 6.79, 0.50, 0.34, "09", 20.0, "#46616A", "Georgia", False, "right")
    return result.shape_count


def p10(prs):
    image = ASSETS / "louvre_abudhabi_night_metalocus.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C["night"], C=C)
    rect(slide, 0.66, 0.60, 0.74, 0.023, fill=C["gold"], C=C)
    tx(slide, 0.66, 0.88, 2.5, 0.12, "09  /  CODA", 7.0, "silver", "Arial Narrow", True)
    cover_image(slide, 0.66, 1.49, 7.45, 4.20, str(image))
    rect(slide, 0.66, 5.98, 7.45, 0.022, fill=C["gold"], C=C)
    tx(slide, 0.70, 6.25, 3.26, 0.12, "LOUVRE ABU DHABI  /  SAADIYAT ISLAND", 6.8, "silver", "Arial Narrow", True)
    tx(slide, 8.70, 1.58, 3.65, 1.60, "LIGHT.\nWATER.\nCITY.", 36.0, "white", "Georgia")
    tx(slide, 8.74, 4.01, 3.07, 0.38, "A common sky turns architecture\ninto a public cultural landscape.", 8.9, "silver", "Arial Narrow", True)
    result = svg(slide, LIGHT_ROUTE, 8.46, 5.02, 3.82, 1.15)
    tx(slide, 8.74, 6.63, 2.08, 0.12, "END OF STUDY", 6.7, "gold", "Arial Narrow", True)
    tx(slide, 12.01, 6.76, 0.50, 0.34, "10", 20.0, "#48626B", "Georgia", False, "right")
    return result.shape_count


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    p6_shapes = p6(prs)
    p7(prs)
    p8(prs)
    p9_shapes = p9(prs)
    p10_shapes = p10(prs)
    prs.save(str(OUT))
    print(OUT)
    print(f"p6_svg_shapes={p6_shapes}; p9_svg_shapes={p9_shapes}; p10_svg_shapes={p10_shapes}")


if __name__ == "__main__":
    main()
