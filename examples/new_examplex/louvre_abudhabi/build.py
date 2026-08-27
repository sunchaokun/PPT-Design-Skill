"""Build the complete editable Louvre Abu Dhabi P1–P10 case study."""
from pathlib import Path

from pptx_designer import Presentation, svg_chart
from pptx_designer.compiler import SVGCompileError
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.shapes import rect
from pptx_designer.tools.text import text

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets" / "images"
OUT = ROOT / "output" / "louvre_abudhabi_complete.pptx"

cover_C = {
    "night": "#11242E", "white": "#F6F6F1", "silver": "#AEBCC0",
    "gold": "#B18C58", "faint": "#49636C", "faint_gold": "#756545",
}

cover_STAR_FIELD = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 780 420">
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


def cover_tx(slide, x, y, w, h, value, size, color, font, bold=False, align=None):
    kw = dict(font_size=size, color=color, font_name=font, bold=bold, C=cover_C)
    if align:
        kw["align"] = align
    return text(slide, x, y, w, h, value, **kw)



p2_C = {
    "paper": "#F4F3ED",
    "ink": "#15232D",
    "muted": "#596973",
    "pale": "#B7C1C2",
    "gold": "#B18C58",
    "white": "#F6F6F1",
}

# A delicate, editable trace of daylight travelling through the dome. It is a
# page component, not a second illustration: intentionally low contrast.
p2_LIGHT_TRACE = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 180">
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


def p2_tx(slide, x, y, w, h, value, size, color, font, bold=False, align=None):
    kw = dict(font_size=size, color=color, font_name=font, bold=bold, C=p2_C)
    if align:
        kw["align"] = align
    return text(slide, x, y, w, h, value, **kw)



p3_C = {
    "night": "#10212B",
    "night2": "#17303B",
    "paper": "#F4F3ED",
    "white": "#F6F6F1",
    "silver": "#AEBCC0",
    "muted": "#576870",
    "gold": "#B18C58",
    "dark": "#16252D",
}

p3_DOME_SECTION = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 980 560">
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

p3_WATER_TRACE = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 500 180">
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


def p3_tx(slide, x, y, w, h, value, size, color, font, bold=False, align=None):
    kw = dict(font_size=size, color=color, font_name=font, bold=bold, C=p3_C)
    if align:
        kw["align"] = align
    return text(slide, x, y, w, h, value, **kw)


def p3_svg(slide, markup, x, y, w, h):
    try:
        return svg_chart(slide, markup, x=x, y=y, w=w, h=h, C=p3_C)
    except SVGCompileError as exc:
        raise RuntimeError(f"SVG compilation failed: {exc}") from exc


def slide_p3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=p3_C["night"], C=p3_C)
    rect(slide, 0.66, 0.63, 0.74, 0.025, fill=p3_C["gold"], C=p3_C)
    p3_tx(slide, 0.66, 0.83, 2.8, 0.14, "02  /  GEOMETRY AS CLIMATE", 7.0,
       "silver", "Arial Narrow", True)
    p3_tx(slide, 0.66, 1.42, 4.5, 1.28, "THE DOME\nAS INSTRUMENT", 31.0,
       "white", "Georgia")
    p3_tx(slide, 0.70, 3.22, 3.65, 0.52,
       "Four outer layers of stainless steel\nmeet four inner layers of aluminium.",
       9.0, "white", "Arial Narrow", True)
    p3_tx(slide, 0.70, 4.12, 3.35, 0.58,
       "A geometric canopy filters daylight, shades\nthe plaza and generates a microclimate below.",
       7.8, "silver", "Arial")
    p3_tx(slide, 0.70, 6.42, 4.2, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.2,
       "silver", "Arial Narrow", True)
    p3_tx(slide, 0.70, 6.75, 1.25, 0.45, "03", 28.0, "#45616B", "Georgia")
    result = p3_svg(slide, p3_DOME_SECTION, 5.08, 0.63, 7.55, 4.76)
    p3_tx(slide, 5.33, 5.54, 2.65, 0.54, "7,850", 37.0, "white", "Arial Narrow")
    p3_tx(slide, 5.38, 6.14, 2.32, 0.12, "STARS IN THE PATTERN", 6.8, "gold", "Arial Narrow", True)
    p3_tx(slide, 9.42, 5.54, 2.70, 0.54, "8", 37.0, "white", "Arial Narrow")
    p3_tx(slide, 9.46, 6.14, 2.32, 0.12, "OVERLAPPING LAYERS", 6.8, "gold", "Arial Narrow", True)
    p3_tx(slide, 5.33, 6.69, 2.65, 0.42, "180 M", 20.0, "silver", "Arial Narrow")
    p3_tx(slide, 7.14, 6.80, 2.90, 0.12, "DIAMETER OF THE DOME", 6.8, "silver", "Arial Narrow", True)
    return result.shape_count


def slide_p4(prs):
    image = ASSETS / "louvre_abudhabi_aerial_expedia.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=p3_C["paper"], C=p3_C)
    p3_tx(slide, 0.66, 0.45, 1.95, 0.14, "03  /  URBAN FIGURE", 7.0,
       "muted", "Arial Narrow", True)
    p3_tx(slide, 0.66, 0.79, 7.72, 0.50, "A MUSEUM CITY  /  IN THE SEA", 26.0,
       "dark", "Georgia")
    p3_tx(slide, 10.73, 0.89, 1.87, 0.12, "ABU DHABI · UAE", 6.6,
       "muted", "Arial Narrow", True, "right")
    rect(slide, 0.66, 1.42, 12.01, 0.022, fill=p3_C["gold"], C=p3_C)
    # Exact 2.4:1 photographic ratio; no image stretching or crop.
    cover_image(slide, 0, 1.66, 13.333, 5.555, str(image))
    rect(slide, 9.03, 5.76, 3.65, 1.15, fill=p3_C["night"], C=p3_C)
    p3_tx(slide, 9.34, 5.98, 3.03, 0.30, "55 BUILDINGS", 17.0,
       "white", "Arial Narrow")
    p3_tx(slide, 9.37, 6.44, 2.82, 0.12, "23 ARE GALLERIES", 6.8,
       "gold", "Arial Narrow", True)
    p3_tx(slide, 0.67, 7.26, 0.55, 0.14, "04", 7.0, "muted", "Arial Narrow", True)


def slide_p5(prs):
    image = ASSETS / "louvre_abudhabi_water_gallery_mediaoffice.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=p3_C["paper"], C=p3_C)
    cover_image(slide, 0.66, 0.56, 4.38, 6.38, str(image))
    rect(slide, 5.33, 0.56, 0.024, 6.38, fill=p3_C["gold"], C=p3_C)
    p3_tx(slide, 5.80, 0.69, 2.0, 0.12, "04  /  CIVIC INTERIOR", 7.0,
       "muted", "Arial Narrow", True)
    p3_tx(slide, 5.80, 1.29, 6.18, 1.42, "WATER\nAS PUBLIC ROOM", 34.0,
       "dark", "Georgia")
    p3_tx(slide, 5.84, 3.16, 5.26, 0.48,
       "Paths cross the water beneath the dome, turning the\nmuseum into a walkable, climate-softened urban room.",
       9.1, "dark", "Arial Narrow", True)
    p3_tx(slide, 5.84, 4.14, 4.78, 0.35,
       "Architecture is not an object beside the sea.\nIt becomes a small city within it.",
       7.8, "muted", "Arial")
    result = p3_svg(slide, p3_WATER_TRACE, 5.80, 5.10, 5.84, 1.05)
    p3_tx(slide, 5.83, 6.60, 2.0, 0.12, "LAND + SEA", 6.6,
       "gold", "Arial Narrow", True)
    p3_tx(slide, 10.80, 6.36, 1.10, 0.46, "05", 28.0,
       "#B7C1C2", "Georgia", False, "right")
    p3_tx(slide, 5.83, 6.92, 3.65, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0,
       "muted", "Arial Narrow", True)
    return result.shape_count



p6_C = {
    "night": "#11242E", "night2": "#1E3843", "paper": "#F4F3ED",
    "white": "#F6F6F1", "silver": "#B3C1C4", "muted": "#586A73",
    "gold": "#B18C58", "dark": "#16252D", "pale": "#CCD4D3",
}

p6_CLIMATE_FIELD = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 720 560">
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

p6_LIGHT_ROUTE = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 760 230">
  <g stroke="#6D858D" stroke-width="1.15">
    <line x1="20" y1="28" x2="740" y2="28"/><line x1="110" y1="102" x2="650" y2="102"/>
    <line x1="200" y1="176" x2="560" y2="176"/>
  </g>
  <g stroke="#B18C58" stroke-width="1.1"><line x1="20" y1="28" x2="200" y2="214"/><line x1="740" y1="28" x2="560" y2="214"/></g>
  <g fill="#B18C58"><circle cx="20" cy="28" r="4"/><circle cx="740" cy="28" r="4"/></g>
  <g fill="#6D858D"><circle cx="200" cy="214" r="3"/><circle cx="560" cy="214" r="3"/></g>
</svg>'''

p6_STAR_FIELD = r'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 560">
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


def p6_tx(slide, x, y, w, h, value, size, color, font, bold=False, align=None):
    kw = dict(font_size=size, color=color, font_name=font, bold=bold, C=p6_C)
    if align:
        kw["align"] = align
    return text(slide, x, y, w, h, value, **kw)


def p6_svg(slide, markup, x, y, w, h):
    try:
        return svg_chart(slide, markup, x=x, y=y, w=w, h=h, C=p6_C)
    except SVGCompileError as exc:
        raise RuntimeError(f"SVG compilation failed: {exc}") from exc


def p6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=p6_C["paper"], C=p6_C)
    p6_tx(slide, 0.66, 0.58, 2.55, 0.12, "05  /  THERMAL FIELD", 7.0, "muted", "Arial Narrow", True)
    rect(slide, 0.66, 1.03, 0.74, 0.023, fill=p6_C["gold"], C=p6_C)
    p6_tx(slide, 0.66, 1.45, 4.08, 1.46, "SHADE\nBECOMES\nINFRASTRUCTURE", 30.0, "dark", "Georgia")
    p6_tx(slide, 0.70, 3.55, 3.65, 0.54, "The canopy filters daylight, reduces heat\nand protects the public realm below.", 9.0, "dark", "Arial Narrow", True)
    p6_tx(slide, 0.70, 4.46, 3.44, 0.43, "Natural cooling is not an add-on.\nIt is the spatial idea of the project.", 7.8, "muted", "Arial")
    result = p6_svg(slide, p6_CLIMATE_FIELD, 5.15, 0.62, 6.96, 4.96)
    rect(slide, 5.24, 5.91, 6.88, 0.023, fill=p6_C["gold"], C=p6_C)
    p6_tx(slide, 5.25, 6.13, 1.88, 0.46, "36 M", 24.0, "dark", "Arial Narrow")
    p6_tx(slide, 7.05, 6.31, 1.60, 0.12, "DOME ABOVE GROUND", 6.5, "muted", "Arial Narrow", True)
    p6_tx(slide, 8.95, 6.13, 1.88, 0.46, "7,500 T", 24.0, "dark", "Arial Narrow")
    p6_tx(slide, 10.78, 6.31, 1.16, 0.12, "DOME WEIGHT", 6.5, "muted", "Arial Narrow", True)
    p6_tx(slide, 0.66, 6.91, 3.52, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0, "muted", "Arial Narrow", True)
    p6_tx(slide, 12.00, 6.79, 0.50, 0.34, "06", 20.0, "pale", "Georgia", False, "right")
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
    rect(slide, 0, 0, 13.333, 7.5, fill=p6_C["night"], C=p6_C)
    p6_tx(slide, 0.66, 0.57, 2.4, 0.12, "06  /  SPATIAL SEQUENCE", 7.0, "silver", "Arial Narrow", True)
    p6_tx(slide, 0.66, 1.09, 6.15, 0.55, "WALK  /  PAUSE  /  RETURN", 27.0, "white", "Georgia")
    p6_tx(slide, 9.73, 1.27, 2.75, 0.13, "A MUSEUM DESIGNED AS A JOURNEY", 6.6, "gold", "Arial Narrow", True, "right")
    rect(slide, 0.66, 1.91, 12.00, 0.022, fill=p6_C["gold"], C=p6_C)
    cover_image(slide, 0.66, 2.30, 3.60, 3.74, str(images[0]))
    cover_image(slide, 4.50, 2.30, 3.60, 3.74, str(images[1]))
    cover_image(slide, 8.34, 2.30, 3.60, 3.74, str(images[2]))
    p6_tx(slide, 0.70, 6.33, 2.65, 0.12, "01  /  FILTERED DAYLIGHT", 6.8, "silver", "Arial Narrow", True)
    p6_tx(slide, 4.54, 6.33, 2.55, 0.12, "02  /  CITY + SEA", 6.8, "silver", "Arial Narrow", True)
    p6_tx(slide, 8.38, 6.33, 2.75, 0.12, "03  /  CIVIC WATERROOM", 6.8, "silver", "Arial Narrow", True)
    p6_tx(slide, 0.66, 6.93, 4.02, 0.12, "THREE ATMOSPHERES, ONE CONTINUOUS ROOF.", 6.5, "gold", "Arial Narrow", True)
    p6_tx(slide, 12.01, 6.76, 0.50, 0.34, "07", 20.0, "#48626B", "Georgia", False, "right")


def p8(prs):
    image = ASSETS / "louvre_abudhabi_rain_of_light.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=p6_C["paper"], C=p6_C)
    cover_image(slide, 0, 0, 13.333, 4.37, str(image))
    rect(slide, 0, 4.37, 13.333, 3.13, fill=p6_C["paper"], C=p6_C)
    p6_tx(slide, 0.66, 4.73, 2.6, 0.12, "07  /  MATERIAL LOGIC", 7.0, "muted", "Arial Narrow", True)
    p6_tx(slide, 0.66, 5.23, 6.48, 0.54, "MATERIAL IS A FILTER", 29.0, "dark", "Georgia")
    p6_tx(slide, 0.70, 6.09, 5.20, 0.40, "Steel, aluminium and patterned voids work as a single\noptical instrument between sky and ground.", 8.7, "dark", "Arial Narrow", True)
    rect(slide, 8.04, 4.73, 0.024, 1.98, fill=p6_C["gold"], C=p6_C)
    p6_tx(slide, 8.36, 4.87, 3.50, 0.34, "8 LAYERS", 20.0, "dark", "Arial Narrow")
    p6_tx(slide, 8.39, 5.34, 3.37, 0.12, "4 STAINLESS STEEL / 4 ALUMINIUM", 6.8, "muted", "Arial Narrow", True)
    p6_tx(slide, 8.36, 5.90, 3.86, 0.28, "The roof permits daylight\nwithout excessive heat or wind.", 7.5, "muted", "Arial")
    p6_tx(slide, 0.66, 7.08, 3.52, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0, "muted", "Arial Narrow", True)
    p6_tx(slide, 12.01, 6.99, 0.50, 0.34, "08", 20.0, "pale", "Georgia", False, "right")


def p9(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=p6_C["night"], C=p6_C)
    result = p6_svg(slide, p6_STAR_FIELD, 4.68, 0.48, 7.84, 5.70)
    p6_tx(slide, 0.66, 0.62, 2.65, 0.12, "08  /  ARCHITECTURAL POSITION", 7.0, "silver", "Arial Narrow", True)
    rect(slide, 0.66, 1.14, 0.74, 0.023, fill=p6_C["gold"], C=p6_C)
    p6_tx(slide, 0.66, 1.67, 6.32, 1.90, "A WELCOMING WORLD\nOF LIGHT, SHADOW,\nREFLECTION AND CALM.", 31.0, "white", "Georgia")
    p6_tx(slide, 0.70, 4.32, 3.7, 0.36, "A museum-city that belongs to its geography\nwithout becoming a literal interpretation of it.", 8.6, "silver", "Arial")
    p6_tx(slide, 0.70, 5.68, 2.95, 0.12, "JEAN NOUVEL  /  ARCHITECT", 6.8, "gold", "Arial Narrow", True)
    p6_tx(slide, 0.70, 6.93, 3.85, 0.12, "SOURCE  /  LOUVRE ABU DHABI — ARCHITECTURE", 6.0, "silver", "Arial Narrow", True)
    p6_tx(slide, 12.00, 6.79, 0.50, 0.34, "09", 20.0, "#46616A", "Georgia", False, "right")
    return result.shape_count


def p10(prs):
    image = ASSETS / "louvre_abudhabi_night_metalocus.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=p6_C["night"], C=p6_C)
    rect(slide, 0.66, 0.60, 0.74, 0.023, fill=p6_C["gold"], C=p6_C)
    p6_tx(slide, 0.66, 0.88, 2.5, 0.12, "09  /  CODA", 7.0, "silver", "Arial Narrow", True)
    cover_image(slide, 0.66, 1.49, 7.45, 4.20, str(image))
    rect(slide, 0.66, 5.98, 7.45, 0.022, fill=p6_C["gold"], C=p6_C)
    p6_tx(slide, 0.70, 6.25, 3.26, 0.12, "LOUVRE ABU DHABI  /  SAADIYAT ISLAND", 6.8, "silver", "Arial Narrow", True)
    p6_tx(slide, 8.70, 1.58, 3.65, 1.60, "LIGHT.\nWATER.\nCITY.", 36.0, "white", "Georgia")
    p6_tx(slide, 8.74, 4.01, 3.07, 0.38, "A common sky turns architecture\ninto a public cultural landscape.", 8.9, "silver", "Arial Narrow", True)
    result = p6_svg(slide, p6_LIGHT_ROUTE, 8.46, 5.02, 3.82, 1.15)
    p6_tx(slide, 8.74, 6.63, 2.08, 0.12, "END OF STUDY", 6.7, "gold", "Arial Narrow", True)
    p6_tx(slide, 12.01, 6.76, 0.50, 0.34, "10", 20.0, "#48626B", "Georgia", False, "right")
    return result.shape_count


def add_cover(prs):
    image = ASSETS / "louvre_abudhabi_night_metalocus.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    C = cover_C
    rect(slide, 0, 0, 13.333, 7.5, fill=C["night"], C=C)
    cover_image(slide, 0, 2.42, 13.333, 4.18, str(image))
    rect(slide, 0.66, 0.55, 0.72, 0.024, fill=C["gold"], C=C)
    cover_tx(slide, 0.66, 0.82, 3.35, 0.12, "ARCHITECTURAL DOSSIER  /  2026", 7.0, "silver", "Arial Narrow", True)
    cover_tx(slide, 0.66, 1.05, 5.30, 1.14, "LOUVRE\nABU DHABI", 40.0, "white", "Georgia")
    cover_tx(slide, 5.12, 2.15, 2.10, 0.11, "JEAN NOUVEL  /  2017", 6.8, "silver", "Arial Narrow", True)
    result = svg_chart(slide, cover_STAR_FIELD, x=8.12, y=0.28, w=4.22, h=2.04, C=C)
    rect(slide, 0.66, 6.83, 12.01, 0.022, fill=C["gold"], C=C)
    cover_tx(slide, 0.70, 7.05, 3.45, 0.12, "SAADIYAT ISLAND  /  ABU DHABI, UAE", 6.7, "silver", "Arial Narrow", True)
    cover_tx(slide, 11.98, 6.94, 0.55, 0.32, "01", 20.0, "silver", "Georgia", False, "right")
    return result.shape_count


def add_p2(prs):
    image = ASSETS / "louvre_abudhabi_rain_of_light.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    C = p2_C
    cover_image(slide, 3.24, 0, 10.093, 7.5, str(image))
    rect(slide, 0, 0, 3.24, 7.5, fill=C["paper"], C=C)
    rect(slide, 3.17, 0, 0.024, 7.5, fill=C["gold"], C=C)
    p2_tx(slide, 0.63, 0.66, 1.95, 0.14, "01  /  SPATIAL PROLOGUE", 7.0, "muted", "Arial Narrow", True)
    rect(slide, 0.63, 1.24, 0.72, 0.023, fill=C["gold"], C=C)
    p2_tx(slide, 0.63, 1.62, 2.16, 1.20, "RAIN\nOF LIGHT", 30.0, "ink", "Georgia")
    p2_tx(slide, 0.66, 3.20, 2.12, 0.58, "BENEATH A POROUS DOME,\nDAYLIGHT BECOMES A\nCULTURAL MEDIUM.", 9.0, "ink", "Arial Narrow", True)
    p2_tx(slide, 0.66, 4.08, 2.12, 0.42, "The roof does not simply shelter the museum;\nit composes its atmosphere.", 7.6, "muted", "Arial")
    result = svg_chart(slide, p2_LIGHT_TRACE, x=0.62, y=4.86, w=2.02, h=1.52, C=C)
    p2_tx(slide, 0.63, 6.63, 1.9, 0.12, "LOUVRE ABU DHABI", 6.6, "muted", "Arial Narrow", True)
    p2_tx(slide, 0.63, 6.89, 1.9, 0.12, "JEAN NOUVEL  /  2017", 6.6, "muted", "Arial Narrow", True)
    p2_tx(slide, 2.25, 6.51, 0.52, 0.45, "02", 28.0, "pale", "Georgia", False, "right")
    return result.shape_count


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    add_cover(prs)
    add_p2(prs)
    slide_p3(prs)
    slide_p4(prs)
    slide_p5(prs)
    p6(prs)
    p7(prs)
    p8(prs)
    p9(prs)
    p10(prs)
    prs.save(str(OUT))
    print(OUT)
    print(f"slides={len(prs.slides)}")


if __name__ == "__main__":
    main()
