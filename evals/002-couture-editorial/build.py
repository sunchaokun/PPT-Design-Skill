"""Evaluation 002: couture editorial deck, Build Mode."""

from pathlib import Path

from pptx_designer import Presentation
from pptx_designer.tools.images import cover_image
from pptx_designer.tools.layout import page_number
from pptx_designer.tools.shapes import rect
from pptx_designer.tools.text import multiline, text


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
OUTPUT = ROOT / "output" / "couture_editorial_eval.pptx"

C = {
    "bone": "#F3EEE7",
    "ink": "#151515",
    "oxblood": "#5B1720",
    "stone": "#B8AEA4",
    "white": "#FFFFFF",
}


def bg(prs, number, color="bone"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, fill=C[color], C=C)
    page_number(slide, number, 6, C=C)
    return slide


def rule(slide, left, top, width, color="ink", height=0.025):
    rect(slide, left, top, width, height, fill=C[color], C=C)


def label(slide, left, top, value, color="oxblood"):
    text(slide, left, top, 2.0, 0.25, value, font_size=10, bold=True, color=color, C=C)


def main():
    prs = Presentation()

    # 1. Manifesto cover: image and type intentionally split.
    slide = bg(prs, 1)
    cover_image(slide, 6.95, 0.0, 6.383, 7.5, str(ASSETS / "couture-editorial.png"))
    label(slide, 0.78, 0.62, "ATELIER NOTES / 01")
    rule(slide, 0.78, 0.98, 2.1, "oxblood")
    text(slide, 0.78, 1.55, 5.4, 0.65, "THE WHITE STUDY", font_size=23, bold=True, color="ink", font_name="Georgia", C=C)
    text(slide, 0.78, 2.45, 4.4, 1.65, "COUTURE\nIN MOTION", font_size=38, bold=True, color="ink", font_name="Georgia", C=C)
    multiline(slide, 0.82, 4.55, 4.2, 0.7, ["A study of volume,", "weight and the hand."], font_size=17, color="ink", C=C, line_spacing=1.1)
    rule(slide, 0.82, 5.75, 1.5, "oxblood")
    text(slide, 0.82, 6.05, 4.8, 0.35, "PARIS / SPRING–SUMMER 2027   PRIVATE EDITORIAL EDITION", font_size=9, color="ink", C=C)
    text(slide, 5.8, 6.65, 0.5, 0.3, "01", font_size=14, bold=True, color="oxblood", C=C, align="right")

    # 2. Body / silhouette: statement opposite a tall crop.
    slide = bg(prs, 2)
    cover_image(slide, 7.25, 0.0, 6.083, 7.5, str(ASSETS / "couture-atelier.png"))
    label(slide, 0.8, 0.7, "02 / SILHOUETTE STUDY")
    rule(slide, 0.8, 1.05, 1.15, "oxblood")
    text(slide, 0.8, 1.65, 5.6, 1.55, "THE BODY\nIS THE PATTERN.", font_size=35, bold=True, color="ink", font_name="Georgia", C=C)
    multiline(slide, 0.82, 3.75, 4.8, 1.0, ["Cut on the bias.", "Volume released at the shoulder."], font_size=19, color="ink", C=C, line_spacing=1.15)
    text(slide, 0.82, 5.45, 3.8, 0.35, "HAND-CUT / HAND-FINISHED", font_size=10, bold=True, color="oxblood", C=C)
    rule(slide, 0.82, 5.9, 3.5, "stone")
    text(slide, 0.82, 6.18, 4.5, 0.4, "A garment begins where the body changes direction.", font_size=13, color="ink", C=C)

    # 3. Volume / movement: image studies with editorial annotation.
    slide = bg(prs, 3)
    label(slide, 0.78, 0.62, "03 / VOLUME STUDY")
    text(slide, 0.78, 1.1, 5.3, 0.7, "CUT\nON THE BIAS", font_size=30, bold=True, color="ink", font_name="Georgia", C=C)
    text(slide, 8.55, 1.2, 3.9, 0.55, "MOTION IS A METHOD.", font_size=16, bold=True, color="oxblood", C=C, align="right")
    cover_image(slide, 0.78, 2.35, 5.55, 3.75, str(ASSETS / "couture-look.png"))
    cover_image(slide, 6.85, 2.0, 5.7, 4.1, str(ASSETS / "couture-editorial.png"))
    label(slide, 0.82, 6.35, "01 / SHOULDER LINE", "ink")
    label(slide, 9.15, 6.55, "02 / RELEASED VOLUME", "ink")
    rule(slide, 0.78, 6.85, 5.55, "oxblood")
    rule(slide, 6.85, 6.85, 5.7, "stone")

    # 4. Material index: specimen logic, not a feature card page.
    slide = bg(prs, 4)
    label(slide, 0.78, 0.62, "04 / MATERIAL INDEX")
    text(slide, 0.78, 1.15, 4.4, 0.8, "LIGHT NEEDS\nTEXTURE.", font_size=30, bold=True, color="ink", font_name="Georgia", C=C)
    multiline(slide, 0.82, 2.65, 3.6, 1.5, ["01  ORGANZA", "AIR / RESISTANCE", "", "02  SATIN", "REFLECTION / WEIGHT"], font_size=16, color="ink", C=C, line_spacing=1.0)
    rule(slide, 0.82, 4.55, 2.25, "oxblood")
    multiline(slide, 0.82, 4.85, 3.6, 1.2, ["03  PEARL", "GRAIN / LIGHT", "", "Each surface changes the gesture."], font_size=16, color="ink", C=C, line_spacing=1.0)
    cover_image(slide, 6.1, 0.0, 7.233, 7.5, str(ASSETS / "couture-material.png"))
    text(slide, 6.4, 6.65, 5.8, 0.35, "MATERIAL / REFLECTION / WEIGHT / GRAIN", font_size=10, bold=True, color="white", C=C)

    # 5. Handwork / fitting: image with sparse annotations and a red line.
    slide = bg(prs, 5, color="ink")
    cover_image(slide, 0.0, 0.0, 7.0, 7.5, str(ASSETS / "couture-atelier.png"))
    label(slide, 7.75, 0.7, "05 / THE HAND", "stone")
    text(slide, 7.75, 1.45, 4.7, 1.35, "NOT A DRESS.\nA GESTURE.", font_size=29, bold=True, color="white", font_name="Georgia", C=C)
    rule(slide, 7.75, 3.45, 1.4, "oxblood")
    multiline(slide, 7.75, 3.85, 4.55, 1.25, ["Fitting is repetition.", "The hand makes the final decision."], font_size=19, color="white", C=C, line_spacing=1.15)
    text(slide, 7.75, 6.45, 4.5, 0.35, "PARIS / 19:30     PRIVATE SALON", font_size=10, bold=True, color="stone", C=C)

    # 6. Close: final fitting, quiet text against a dark field and material crop.
    slide = bg(prs, 6, color="ink")
    cover_image(slide, 7.0, 0.0, 6.333, 7.5, str(ASSETS / "couture-material.png"))
    label(slide, 0.8, 0.7, "UNFINISHED", "stone")
    text(slide, 0.8, 1.5, 5.8, 1.45, "THE FINAL\nFITTING", font_size=37, bold=True, color="white", font_name="Georgia", C=C)
    rule(slide, 0.8, 3.45, 1.4, "oxblood")
    multiline(slide, 0.84, 3.85, 5.5, 1.1, ["A garment is never still.", "It keeps the memory of the hand."], font_size=20, color="white", C=C, line_spacing=1.15)
    text(slide, 0.84, 6.55, 5.8, 0.35, "THE WHITE STUDY — COUTURE / MOTION", font_size=10, bold=True, color="stone", C=C)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUTPUT))
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
