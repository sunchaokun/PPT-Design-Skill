"""Build the complete editable Louvre Abu Dhabi P1–P10 case study."""
from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.compiler import SVGCompileError

import build_page01_cover as cover
import build_page02_spatial_prologue as p2
import build_pages03_05_dome_to_water as p3_p5
import build_pages06_10_climate_to_coda as p6_p10


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output" / "louvre_abudhabi_complete.pptx"


def add_cover(prs):
    image = ROOT / "assets" / "images" / "louvre_abudhabi_night_metalocus.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    C = cover.C
    cover.rect(slide, 0, 0, 13.333, 7.5, fill=C["night"], C=C)
    cover.cover_image(slide, 0, 2.42, 13.333, 4.18, str(image))
    cover.rect(slide, 0.66, 0.55, 0.72, 0.024, fill=C["gold"], C=C)
    cover.tx(slide, 0.66, 0.82, 3.35, 0.12, "ARCHITECTURAL DOSSIER  /  2026", 7.0, "silver", "Arial Narrow", True)
    cover.tx(slide, 0.66, 1.05, 5.30, 1.14, "LOUVRE\nABU DHABI", 40.0, "white", "Georgia")
    cover.tx(slide, 5.12, 2.15, 2.10, 0.11, "JEAN NOUVEL  /  2017", 6.8, "silver", "Arial Narrow", True)
    try:
        result = cover.svg_chart(slide, cover.STAR_FIELD, x=8.12, y=0.28, w=4.22, h=2.04, C=C)
    except SVGCompileError as exc:
        raise RuntimeError(f"Cover SVG compilation failed: {exc}") from exc
    cover.rect(slide, 0.66, 6.83, 12.01, 0.022, fill=C["gold"], C=C)
    cover.tx(slide, 0.70, 7.05, 3.45, 0.12, "SAADIYAT ISLAND  /  ABU DHABI, UAE", 6.7, "silver", "Arial Narrow", True)
    cover.tx(slide, 11.98, 6.94, 0.55, 0.32, "01", 20.0, "silver", "Georgia", False, "right")
    return result.shape_count


def add_p2(prs):
    image = ROOT / "assets" / "images" / "louvre_abudhabi_rain_of_light.jpg"
    if not image.exists():
        raise FileNotFoundError(image)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    C = p2.C
    p2.cover_image(slide, 3.24, 0, 10.093, 7.5, str(image))
    p2.rect(slide, 0, 0, 3.24, 7.5, fill=C["paper"], C=C)
    p2.rect(slide, 3.17, 0, 0.024, 7.5, fill=C["gold"], C=C)
    p2.tx(slide, 0.63, 0.66, 1.95, 0.14, "01  /  SPATIAL PROLOGUE", 7.0, "muted", "Arial Narrow", True)
    p2.rect(slide, 0.63, 1.24, 0.72, 0.023, fill=C["gold"], C=C)
    p2.tx(slide, 0.63, 1.62, 2.16, 1.20, "RAIN\nOF LIGHT", 30.0, "ink", "Georgia")
    p2.tx(slide, 0.66, 3.20, 2.12, 0.58, "BENEATH A POROUS DOME,\nDAYLIGHT BECOMES A\nCULTURAL MEDIUM.", 9.0, "ink", "Arial Narrow", True)
    p2.tx(slide, 0.66, 4.08, 2.12, 0.42, "The roof does not simply shelter the museum;\nit composes its atmosphere.", 7.6, "muted", "Arial")
    try:
        result = p2.svg_chart(slide, p2.LIGHT_TRACE, x=0.62, y=4.86, w=2.02, h=1.52, C=C)
    except SVGCompileError as exc:
        raise RuntimeError(f"P2 SVG compilation failed: {exc}") from exc
    p2.tx(slide, 0.63, 6.63, 1.9, 0.12, "LOUVRE ABU DHABI", 6.6, "muted", "Arial Narrow", True)
    p2.tx(slide, 0.63, 6.89, 1.9, 0.12, "JEAN NOUVEL  /  2017", 6.6, "muted", "Arial Narrow", True)
    p2.tx(slide, 2.25, 6.51, 0.52, 0.45, "02", 28.0, "pale", "Georgia", False, "right")
    return result.shape_count


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    cover_shapes = add_cover(prs)
    p2_shapes = add_p2(prs)
    p3_p5.slide_p3(prs)
    p3_p5.slide_p4(prs)
    p3_p5.slide_p5(prs)
    p6_p10.p6(prs)
    p6_p10.p7(prs)
    p6_p10.p8(prs)
    p6_p10.p9(prs)
    p6_p10.p10(prs)
    prs.save(str(OUT))
    print(OUT)
    print(f"cover_svg_shapes={cover_shapes}; p2_svg_shapes={p2_shapes}; slides={len(prs.slides)}")


if __name__ == "__main__":
    main()
