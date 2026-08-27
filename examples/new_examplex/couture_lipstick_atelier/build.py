"""Build the editable COUTURE COLOR flagship portfolio case."""
from pathlib import Path

from pptx_designer import Presentation, svg_chart
from pptx_designer.compiler import SVGCompileError
from pptx_designer.tools.images import cover_image, gradient_mask_image
from pptx_designer.tools.shapes import rect, rrect
from pptx_designer.tools.text import text


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output" / "couture_color_objects_of_desire.pptx"
ASSETS = ROOT / "assets" / "images"

C = {
    "ink": "#160B10", "aubergine": "#261017", "oxblood": "#5C0A21",
    "cherry": "#8E1633", "rose": "#C98C88", "bone": "#F3EADF",
    "champagne": "#C8A76C", "smoke": "#6D5559", "black": "#080607",
    "white": "#FFF9F2", "pale": "#DECBC1",
}

W, H = 13.333, 7.5

SWEEP = """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 800 240'>
<path d='M-20 192 C150 205 251 9 419 58 C556 98 606 224 824 21'
fill='none' stroke='#C8A76C' stroke-width='2' opacity='.65'/>
<path d='M-20 211 C161 226 268 36 428 83 C558 121 617 243 824 47'
fill='none' stroke='#C8A76C' stroke-width='1' opacity='.40'/></svg>"""


def tx(slide, x, y, w, h, value, size, color="bone", font="Arial", bold=False, align="left"):
    return text(slide, left=x, top=y, width=w, height=h, txt=value, font_size=size,
                color=color, font_name=font, bold=bold, align=align, C=C)


def base(slide, page, dark=True):
    bg = "ink" if dark else "bone"
    rect(slide, 0, 0, W, H, fill=bg, C=C)
    rule = "champagne" if dark else "oxblood"
    rect(slide, .56, .38, 12.21, .018, fill=rule, C=C)
    tx(slide, .56, 7.05, 2.9, .13, "COUTURE COLOR  /  CONCEPT COLLECTION", 6.4, "pale" if dark else "smoke", "Arial Narrow", True)
    tx(slide, 12.13, 6.93, .62, .25, f"{page:02d}", 15, "champagne" if dark else "oxblood", "Georgia", False, "right")


def svg(slide, art, x, y, w, h):
    try:
        result = svg_chart(slide, art, x=x, y=y, w=w, h=h, C=C)
        if result.warnings:
            print("SVG warnings:", result.warnings)
        return result.shape_count
    except SVGCompileError as exc:
        raise RuntimeError(f"SVG compilation failed: {exc}") from exc


def lipstick(slide, x, y, scale=1.0, body="oxblood", bullet="cherry", metal="champagne"):
    """Native editable product silhouette."""
    rrect(slide, x, y + .62*scale, .58*scale, 1.42*scale, fill=body, line=metal, C=C)
    rect(slide, x + .035*scale, y + .59*scale, .51*scale, .16*scale, fill=metal, C=C)
    rrect(slide, x + .09*scale, y + .20*scale, .40*scale, .47*scale, fill=metal, line=metal, C=C)
    # angled-looking bullet, made from an editable rectangle + small cap
    rrect(slide, x + .14*scale, y, .30*scale, .30*scale, fill=bullet, line=bullet, C=C)
    rect(slide, x + .14*scale, y + .22*scale, .30*scale, .10*scale, fill=bullet, C=C)


def s1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cover_image(slide, 0, 0, W, H, str(ASSETS / "couture_color_model_cover_v2.png"))
    gradient_mask_image(slide, 0, 0, W, H, bg_color=C["ink"], direction="right", alpha_start=76, alpha_end=0)
    tx(slide, .66, .72, 3.0, .14, "PORTFOLIO CASE  /  2026", 7, "champagne", "Arial Narrow", True)
    tx(slide, .66, 1.22, 5.1, 1.68, "COUTURE\nCOLOR", 48, "bone", "Georgia")
    tx(slide, .70, 3.35, 4.6, .42, "OBJECTS OF DESIRE", 14, "rose", "Arial Narrow", True)
    tx(slide, .70, 4.14, 3.95, .55, "A couture lipstick collection\nwith a face, a gesture, and a trace.", 11.2, "pale", "Arial")
    rect(slide, .70, 6.31, 4.95, .02, fill="champagne", C=C)
    tx(slide, .70, 6.55, 4.6, .14, "CONCEPTUAL BEAUTY EDITORIAL  /  ENGLISH", 6.4, "pale", "Arial Narrow", True)
    tx(slide, 12.13, 6.93, .62, .25, "01", 15, "bone", "Georgia", False, "right")


def s2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 2, False)
    cover_image(slide, 6.50, 0, 6.833, 7.5, str(ASSETS / "couture_color_model_application_same_subject_v3.png"))
    rect(slide, 6.46, 0, .04, 7.5, fill="oxblood", C=C)
    tx(slide, .62, .78, 3.2, .14, "01  /  THE PROPOSITION", 7, "oxblood", "Arial Narrow", True)
    tx(slide, .62, 1.42, 5.0, 1.78, "COLOR IS\nA COUTURE\nGESTURE.", 31, "ink", "Georgia")
    tx(slide, .66, 4.16, 4.82, .54, "Not an accessory to a look — a small,\nprecise object that finishes a point of view.", 10, "smoke", "Arial")
    rect(slide, .66, 5.30, 4.85, .02, fill="oxblood", C=C)
    tx(slide, .66, 5.58, 1.35, .22, "FORM", 8, "oxblood", "Arial Narrow", True)
    tx(slide, 2.22, 5.58, 1.35, .22, "PIGMENT", 8, "oxblood", "Arial Narrow", True)
    tx(slide, 3.90, 5.58, 1.35, .22, "RITUAL", 8, "oxblood", "Arial Narrow", True)
    tx(slide, .66, 5.88, 1.35, .28, "a tailored case", 7.5, "smoke")
    tx(slide, 2.22, 5.88, 1.35, .28, "a held intensity", 7.5, "smoke")
    tx(slide, 3.90, 5.88, 1.35, .28, "a private finale", 7.5, "smoke")


def s3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 3, True)
    tx(slide, .62, .77, 4.0, .14, "02  /  THE CHROMATIC WARDROBE", 7, "champagne", "Arial Narrow", True)
    tx(slide, .62, 1.35, 6.0, .86, "SIX SHADES.\nONE AFTER-DARK SILHOUETTE.", 25, "bone", "Georgia")
    shades = [("01", "NUDE VEIL", "#C79078"), ("02", "ROSE TAILLEUR", "#A95664"), ("03", "GARNET", "#83182E"), ("04", "VIOLET NOIR", "#43202D"), ("05", "OXBLOOD", "#5C0A21"), ("06", "BLACK CHERRY", "#26050D")]
    x0 = .66
    for i, (n, name, hexv) in enumerate(shades):
        x = x0 + i*2.03
        rrect(slide, x, 3.28, 1.62, 1.60, fill=hexv, line=hexv, C=C)
        tx(slide, x+.12, 3.55, 1.30, .18, n, 8, "bone", "Arial Narrow", True)
        tx(slide, x+.12, 4.45, 1.30, .18, name, 6.7, "bone", "Arial Narrow", True)
    tx(slide, .66, 5.54, 5.3, .36, "A wardrobe is not a spectrum. It is a sequence of attitudes.", 10, "pale", "Arial")
    tx(slide, .66, 6.13, 10.8, .18, "THE COLOR SYSTEM IS ENTIRELY EDITABLE: SWATCHES, NAMES, ROLES, AND ORDER.", 6.5, "champagne", "Arial Narrow", True)


def s4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 4, False)
    tx(slide, .62, .77, 3.4, .14, "03  /  OBJECT GRAMMAR", 7, "oxblood", "Arial Narrow", True)
    tx(slide, .62, 1.32, 3.7, .82, "TAILORED\nTO THE HAND.", 26, "ink", "Georgia")
    tx(slide, .66, 3.12, 3.38, .54, "The case is conceived as an architecture\nof proportion, weight, and interruption.", 9.5, "smoke", "Arial")
    rect(slide, 4.68, .80, .018, 5.93, fill="oxblood", C=C)
    lipstick(slide, 8.02, 1.45, 1.95)
    # exploded rings / native callouts
    rect(slide, 8.05, 1.08, 1.06, .04, fill="champagne", C=C)
    rect(slide, 8.05, 3.13, 1.06, .04, fill="champagne", C=C)
    rect(slide, 8.05, 4.58, 1.06, .04, fill="champagne", C=C)
    for y, label, note in [(1.00, "01 / BULLET", "an oblique, held edge"), (3.05, "02 / COLLAR", "a warm metal pause"), (4.50, "03 / CASE", "a lacquered vertical")]:
        tx(slide, 9.65, y, 1.85, .17, label, 7.5, "oxblood", "Arial Narrow", True)
        tx(slide, 9.65, y+.25, 2.15, .24, note, 7.2, "smoke", "Arial")
    tx(slide, 5.28, 5.73, 2.0, .16, "PROPORTION", 7, "oxblood", "Arial Narrow", True)
    tx(slide, 5.28, 6.06, 2.0, .36, "A narrow column\nmeets a broad base.", 9, "ink", "Georgia")


def s5(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 5, True)
    cover_image(slide, .56, .74, 7.12, 5.70, str(ASSETS / "couture_lipstick_macro.png"))
    rect(slide, 7.67, .74, .035, 5.70, fill="champagne", C=C)
    tx(slide, 8.20, .86, 3.7, .14, "04  /  MATERIAL CHOREOGRAPHY", 7, "champagne", "Arial Narrow", True)
    tx(slide, 8.18, 1.46, 3.85, 1.20, "PIGMENT\nHAS A MEMORY.", 27, "bone", "Georgia")
    materials = [("01", "LACQUER", "depth / reflection"), ("02", "METAL", "weight / temperature"), ("03", "PIGMENT", "gesture / trace")]
    for i, (n, name, detail) in enumerate(materials):
        y = 3.32 + i*.72
        tx(slide, 8.20, y, .35, .16, n, 7, "champagne", "Arial Narrow", True)
        tx(slide, 8.68, y, 1.35, .16, name, 8, "bone", "Arial Narrow", True)
        tx(slide, 10.18, y, 1.82, .18, detail, 7.4, "pale", "Arial")
        rect(slide, 8.18, y+.38, 3.82, .012, fill="smoke", C=C)
    tx(slide, 8.20, 5.90, 3.50, .30, "The material story must be readable\nbefore the product is named.", 9, "pale", "Georgia")


def s6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 6, False)
    svg(slide, SWEEP, .48, 1.90, 12.34, 3.70)
    tx(slide, .62, .77, 2.4, .14, "05  /  THE GESTURE", 7, "oxblood", "Arial Narrow", True)
    tx(slide, .62, 1.43, 12.1, 1.15, "A SINGLE, HELD LINE.", 35, "ink", "Georgia")
    tx(slide, .66, 3.25, 3.1, .32, "APPLICATION IS THE LAST\nTAILORING MOVE.", 9, "oxblood", "Arial Narrow", True)
    tx(slide, 8.90, 4.86, 2.95, .43, "A line that does not decorate.\nA line that decides.", 11, "smoke", "Georgia")
    rrect(slide, 1.05, 4.74, 5.72, .20, fill="oxblood", line="oxblood", C=C)
    rrect(slide, 1.05, 4.74, 2.00, .20, fill="cherry", line="cherry", C=C)
    tx(slide, .66, 5.74, 5.30, .18, "EDITABLE SVG SWEEP / NATIVE COLOR TRACE", 6.5, "smoke", "Arial Narrow", True)


def s7(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 7, True)
    tx(slide, .62, .77, 3.2, .14, "06  /  THE COLLECTION", 7, "champagne", "Arial Narrow", True)
    tx(slide, .62, 1.35, 6.0, .75, "THREE TONES.\nTHREE POSTURES.", 26, "bone", "Georgia")
    groups = [("THE NUDE", "quiet structure", "#C79078"), ("THE ROSE", "soft armour", "#A95664"), ("THE RED", "held intensity", "#83182E")]
    for i, (name, note, shade) in enumerate(groups):
        x = .85 + i*4.08
        rect(slide, x, 3.05, 3.34, .02, fill="champagne", C=C)
        lipstick(slide, x+1.32, 3.36, 1.02, "aubergine", shade)
        tx(slide, x, 5.70, 3.34, .20, name, 9, "bone", "Arial Narrow", True, "center")
        tx(slide, x, 6.06, 3.34, .18, note, 8, "pale", "Georgia", False, "center")


def s8(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cover_image(slide, 0, 0, W, H, str(ASSETS / "couture_lipstick_ritual.png"))
    rect(slide, 0, 0, 3.55, H, fill="ink", C=C)
    rect(slide, 3.52, 0, .03, H, fill="champagne", C=C)
    tx(slide, .62, .76, 2.35, .14, "07  /  THE RITUAL", 7, "champagne", "Arial Narrow", True)
    tx(slide, .62, 1.42, 2.38, 1.25, "A PRIVATE\nFINAL LOOK.", 26, "bone", "Georgia")
    tx(slide, .66, 3.28, 2.23, .54, "The collection belongs on a dressing table\nbefore it belongs to a campaign.", 9, "pale", "Arial")
    rect(slide, .66, 4.44, 2.20, .02, fill="champagne", C=C)
    tx(slide, .66, 4.74, 2.24, .16, "RETAIL SETTING", 7, "champagne", "Arial Narrow", True)
    tx(slide, .66, 5.08, 2.25, .40, "Low light. Warm metal.\nA deliberate pause.", 9.5, "bone", "Georgia")
    tx(slide, .66, 6.38, 2.38, .14, "COUTURE COLOR / CONCEPT ONLY", 6.2, "pale", "Arial Narrow", True)
    tx(slide, 12.13, 6.93, .62, .25, "08", 15, "bone", "Georgia", False, "right")


def s9(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 9, False)
    tx(slide, .62, .77, 3.0, .14, "08  /  EDITORIAL RELEASE", 7, "oxblood", "Arial Narrow", True)
    tx(slide, .62, 1.36, 5.6, .76, "REVEAL, HOLD,\nLEAVE A TRACE.", 25, "ink", "Georgia")
    steps = [("I", "THE OBJECT", "a single red form"), ("II", "THE WARDROBE", "six intentional shades"), ("III", "THE GESTURE", "one held application"), ("IV", "THE RITUAL", "a private finale")]
    y = 3.18
    rect(slide, .88, y+.20, 10.96, .025, fill="rose", C=C)
    for i, (roman, title, note) in enumerate(steps):
        x = .82 + i*3.0
        rrect(slide, x, y, .46, .46, fill="oxblood", line="oxblood", C=C)
        tx(slide, x, y+.105, .46, .14, roman, 7.5, "bone", "Arial Narrow", True, "center")
        tx(slide, x, y+.78, 2.34, .18, title, 8.5, "oxblood", "Arial Narrow", True)
        tx(slide, x, y+1.12, 2.34, .28, note, 8, "smoke", "Georgia")
    tx(slide, .66, 5.84, 6.12, .32, "A release system built around attention — not fabricated reach, sales, or performance claims.", 9, "smoke", "Arial")


def s10(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); base(slide, 10, True)
    tx(slide, .62, .80, 2.8, .14, "09  /  CODA", 7, "champagne", "Arial Narrow", True)
    tx(slide, .62, 1.55, 8.7, 1.32, "MAKE COLOR\nA POINT OF VIEW.", 35, "bone", "Georgia")
    tx(slide, .66, 3.46, 4.55, .46, "Couture is not excess. It is the discipline\nof deciding what the object needs to say.", 10, "pale", "Arial")
    lipstick(slide, 10.24, 3.38, 1.65, "aubergine", "cherry")
    rect(slide, .66, 5.34, 8.12, .02, fill="champagne", C=C)
    tx(slide, .66, 5.72, 6.8, .16, "COUTURE COLOR / A FICTIONAL PORTFOLIO CASE", 7, "champagne", "Arial Narrow", True)
    tx(slide, .66, 6.10, 8.0, .28, "All imagery generated for this case. Product language, geometry, and typography remain editable.", 7.2, "pale", "Arial")


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    for add in (s1, s2, s3, s4, s5, s6, s7, s8, s9, s10):
        add(prs)
    prs.save(str(OUT))
    print(OUT)
    print(f"slides={len(prs.slides)}")


if __name__ == "__main__":
    main()
