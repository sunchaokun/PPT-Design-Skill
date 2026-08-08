"""Blip Fill — fill shapes with images via OOXML a:blipFill.

Enables images inside circles, hexagons, diamonds, and any MSO_SHAPE.
Uses slide.part.get_or_add_image_part() to register images and obtain rId,
then constructs a:blipFill XML to fill the shape geometry with the image.
"""

from __future__ import annotations

import os

from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches
from lxml import etree


ARTISTIC_EFFECTS: dict[str, tuple[str, dict[str, str]]] = {
    "watercolor_sponge": ("a:artisticWatercolorSponge", {"brushSize": "2", "smoothing": "3"}),
    "pencil_grayscale": ("a:artisticPencilGrayscale", {"pencilSize": "2", "pressure": "3"}),
    "pencil_colored": ("a:artisticPencilColored", {"pencilSize": "2", "pressure": "3"}),
    "mosaic_bubbles": ("a:artisticMosaicBubbles", {"gridSize": "12"}),
    "film_grain": ("a:artisticFilmGrain", {"grainSize": "2", "intensity": "50"}),
    "glow_diffused": ("a:artisticGlowDiffused", {"glowRadius": "5", "intensity": "50"}),
    "blur": ("a:artisticBlur", {"radius": "5"}),
    "cutout": ("a:artisticCutout", {"numberOfShades": "4", "simplify": "3"}),
    "marker": ("a:artisticMarker", {"size": "3", "intensity": "80"}),
    "paint_strokes": ("a:artisticPaintStrokes", {"intensity": "50", "size": "3"}),
    "texturizer": ("a:artisticTexturizer", {"scaling": "50", "textureType": "1"}),
    "light_screen": ("a:artisticLightScreen", {"gridSize": "8", "transparency": "30"}),
    "line_drawing": ("a:artisticLineDrawing", {"pencilSize": "2", "intensity": "50"}),
    "etching": ("a:artisticEtching", {"intensity": "50"}),
    "plastic": ("a:artisticPlastic", {"intensity": "50"}),
    "glass": ("a:artisticGlass", {"scaling": "50", "intensity": "50"}),
    "cement": ("a:artisticCement", {"scaling": "50", "intensity": "50"}),
    "chalk_smokey": ("a:artisticChalkSmokey", {"smudge": "3", "intensity": "50"}),
    "crayon": ("a:artisticCrayon", {"size": "3", "intensity": "50"}),
    "halftone": ("a:artisticHalftone", {"gridSize": "8"}),
    "photocopy": ("a:artisticPhotocopy", {"detail": "3", "intensity": "50"}),
    "stamp": ("a:artisticStamp", {"intensity": "50"}),
}


def _register_image(slide, image_path: str) -> str | None:
    try:
        _image_part, rId = slide.part.get_or_add_image_part(image_path)
        return rId
    except Exception:
        return None


def _remove_shape_fill(spPr: etree._Element) -> None:
    for tag in ("a:solidFill", "a:gradFill", "a:noFill", "a:blipFill", "a:pattFill"):
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)


def _set_no_line(spPr: etree._Element) -> None:
    ln = spPr.find(qn("a:ln"))
    if ln is not None:
        spPr.remove(ln)
    ln = etree.SubElement(spPr, qn("a:ln"))
    etree.SubElement(ln, qn("a:noFill"))


def fill_shape_with_image(shape, slide, image_path: str,
                          crop_mode: str = "stretch",
                          alpha: int = 100) -> str | None:
    if not image_path or not _file_exists(image_path):
        return None
    rId = _register_image(slide, image_path)
    if rId is None:
        return None
    spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return None
    _remove_shape_fill(spPr)
    blipFill = etree.SubElement(spPr, qn("a:blipFill"))
    blip = etree.SubElement(blipFill, qn("a:blip"))
    blip.set(qn("r:embed"), rId)
    if alpha < 100:
        a = etree.SubElement(blip, qn("a:alpha"))
        a.set("val", str(alpha * 1000))
    if crop_mode == "tile":
        etree.SubElement(blipFill, qn("a:tile"))
    else:
        etree.SubElement(blipFill, qn("a:srcRect"))
        stretch = etree.SubElement(blipFill, qn("a:stretch"))
        etree.SubElement(stretch, qn("a:fillRect"))
    _set_no_line(spPr)
    return rId


def add_image_in_shape(slide, shape_type, x: float, y: float,
                       w: float, h: float, image_path: str,
                       crop_mode: str = "stretch", alpha: int = 100,
                       border_hex: str | None = None) -> object:
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    rId = fill_shape_with_image(shape, slide, image_path,
                                crop_mode=crop_mode, alpha=alpha)
    if rId is None:
        return shape
    if border_hex:
        spPr = shape._element.find(qn("p:spPr"))
        ln = spPr.find(qn("a:ln"))
        if ln is not None:
            spPr.remove(ln)
        ln = etree.SubElement(spPr, qn("a:ln"))
        ln.set("w", "12700")
        solidFill = etree.SubElement(ln, qn("a:solidFill"))
        srgb = etree.SubElement(solidFill, qn("a:srgbClr"))
        srgb.set("val", border_hex.lstrip("#"))
    return shape


def add_circle_image(slide, cx: float, cy: float, radius: float,
                     image_path: str, border_hex: str | None = None) -> object:
    x = cx - radius
    y = cy - radius
    size = radius * 2
    return add_image_in_shape(slide, MSO_SHAPE.OVAL, x, y, size, size,
                              image_path, border_hex=border_hex)


def add_hexagon_image(slide, cx: float, cy: float, size: float,
                      image_path: str, border_hex: str | None = None) -> object:
    x = cx - size / 2
    y = cy - size * 0.87 / 2
    return add_image_in_shape(slide, MSO_SHAPE.HEXAGON, x, y, size, size * 0.87,
                              image_path, border_hex=border_hex)


def add_diamond_image(slide, cx: float, cy: float, size: float,
                      image_path: str, border_hex: str | None = None) -> object:
    x = cx - size / 2
    y = cy - size / 2
    return add_image_in_shape(slide, MSO_SHAPE.DIAMOND, x, y, size, size,
                              image_path, border_hex=border_hex)


# ── OOXML blip effects ──


def _get_blip(shape) -> etree._Element | None:
    spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return None
    blipFill = spPr.find(qn("a:blipFill"))
    if blipFill is None:
        return None
    return blipFill.find(qn("a:blip"))


def apply_blip_grayscale(shape) -> None:
    blip = _get_blip(shape)
    if blip is None:
        return
    existing = blip.find(qn("a:grayscl"))
    if existing is not None:
        blip.remove(existing)
    etree.SubElement(blip, qn("a:grayscl"))


def apply_blip_duotone(shape, color1: str, color2: str) -> None:
    blip = _get_blip(shape)
    if blip is None:
        return
    existing = blip.find(qn("a:duotone"))
    if existing is not None:
        blip.remove(existing)
    duotone = etree.SubElement(blip, qn("a:duotone"))
    srgb1 = etree.SubElement(duotone, qn("a:srgbClr"))
    srgb1.set("val", color1.lstrip("#"))
    srgb2 = etree.SubElement(duotone, qn("a:srgbClr"))
    srgb2.set("val", color2.lstrip("#"))


def apply_blip_brightness_contrast(shape, bright_pct: int = 0,
                                   contrast_pct: int = 0) -> None:
    blip = _get_blip(shape)
    if blip is None:
        return
    existing = blip.find(qn("a:lum"))
    if existing is not None:
        blip.remove(existing)
    lum = etree.SubElement(blip, qn("a:lum"))
    lum.set("bright", str(bright_pct * 1000))
    lum.set("contrast", str(contrast_pct * 1000))


def apply_blip_saturation(shape, saturation_pct: int = 100) -> None:
    blip = _get_blip(shape)
    if blip is None:
        return
    existing = blip.find(qn("a:sat"))
    if existing is not None:
        blip.remove(existing)
    sat = etree.SubElement(blip, qn("a:sat"))
    sat.set("val", str(saturation_pct * 1000))


def apply_blip_artistic(shape, effect: str,
                        params: dict[str, str] | None = None) -> None:
    blip = _get_blip(shape)
    if blip is None:
        return
    if effect not in ARTISTIC_EFFECTS:
        raise KeyError(f"Unknown artistic effect: {effect!r}. Available: {list(ARTISTIC_EFFECTS.keys())}")
    tag_name, default_params = ARTISTIC_EFFECTS[effect]
    existing = blip.find(qn(tag_name))
    if existing is not None:
        blip.remove(existing)
    el = etree.SubElement(blip, qn(tag_name))
    merged = {**default_params, **(params or {})}
    for k, v in merged.items():
        el.set(k, v)


def _file_exists(path: str) -> bool:
    return os.path.isfile(path)
