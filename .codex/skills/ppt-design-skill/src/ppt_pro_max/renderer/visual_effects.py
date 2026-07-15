"""Visual Effects — gradient fills, shadows, soft edges, reflections via XML.

This module bridges the gap between python-pptx's limited API and the rich
visual effects available in OOXML. All methods operate on shape elements
directly via lxml, enabling PPT-7 quality output.

Key capabilities:
  - Linear & path (radial) gradient fills with N stops
  - Outer shadows with configurable blur/distance/direction/opacity
  - Soft edges for smooth blending
  - Reflection effects
  - Glow effects
  - Alpha (transparency) on any color
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from pptx.oxml.ns import qn
from lxml import etree


@dataclass
class GradientStop:
    color: str | None = None
    position: int = 0
    alpha: int | None = None
    scheme_color: str | None = None
    lum_mod: int | None = None
    lum_off: int | None = None

    def to_xml(self) -> etree._Element:
        gs = etree.SubElement(etree.Element("dummy"), qn("a:gs"))
        gs.set("pos", str(self.position))
        if self.scheme_color:
            clr = etree.SubElement(gs, qn("a:schemeClr"))
            clr.set("val", self.scheme_color)
            if self.lum_mod is not None:
                lm = etree.SubElement(clr, qn("a:lumMod"))
                lm.set("val", str(self.lum_mod))
            if self.lum_off is not None:
                lo = etree.SubElement(clr, qn("a:lumOff"))
                lo.set("val", str(self.lum_off))
            if self.alpha is not None:
                a = etree.SubElement(clr, qn("a:alpha"))
                a.set("val", str(self.alpha))
        else:
            srgb = etree.SubElement(gs, qn("a:srgbClr"))
            srgb.set("val", self.color.lstrip("#"))
            if self.alpha is not None:
                alpha_el = etree.SubElement(srgb, qn("a:alpha"))
                alpha_el.set("val", str(self.alpha))
        return gs


@dataclass
class GradientFill:
    stops: list[GradientStop] = field(default_factory=list)
    angle: int = 5400000
    gradient_type: str = "linear"
    scaled: bool = True
    fill_to_rect: dict[str, str] | None = None

    def apply(self, shape) -> None:
        spPr = shape._element.find(qn("p:spPr"))
        if spPr is None:
            return
        _remove_existing_fill(spPr)

        gradFill = etree.SubElement(spPr, qn("a:gradFill"))
        gsLst = etree.SubElement(gradFill, qn("a:gsLst"))
        for stop in self.stops:
            gs = etree.SubElement(gsLst, qn("a:gs"))
            gs.set("pos", str(stop.position))
            if stop.scheme_color:
                clr = etree.SubElement(gs, qn("a:schemeClr"))
                clr.set("val", stop.scheme_color)
                if stop.lum_mod is not None:
                    lm = etree.SubElement(clr, qn("a:lumMod"))
                    lm.set("val", str(stop.lum_mod))
                if stop.lum_off is not None:
                    lo = etree.SubElement(clr, qn("a:lumOff"))
                    lo.set("val", str(stop.lum_off))
                if stop.alpha is not None:
                    alpha_el = etree.SubElement(clr, qn("a:alpha"))
                    alpha_el.set("val", str(stop.alpha))
            else:
                srgb = etree.SubElement(gs, qn("a:srgbClr"))
                srgb.set("val", stop.color.lstrip("#"))
                if stop.alpha is not None:
                    alpha_el = etree.SubElement(srgb, qn("a:alpha"))
                    alpha_el.set("val", str(stop.alpha))

        if self.gradient_type == "linear":
            lin = etree.SubElement(gradFill, qn("a:lin"))
            lin.set("ang", str(self.angle))
            lin.set("scaled", "1" if self.scaled else "0")
        elif self.gradient_type == "path":
            path_el = etree.SubElement(gradFill, qn("a:path"))
            path_el.set("path", "circle")
            fillToRect = etree.SubElement(path_el, qn("a:fillToRect"))
            if self.fill_to_rect:
                for k, v in self.fill_to_rect.items():
                    fillToRect.set(k, v)
            else:
                fillToRect.set("l", "50000")
                fillToRect.set("t", "50000")
                fillToRect.set("r", "50000")
                fillToRect.set("b", "50000")


@dataclass
class Shadow:
    blur_pt: float = 6.0
    distance_pt: float = 3.0
    direction_deg: float = 90.0
    color: str = "#000000"
    alpha_pct: int = 25
    align: str = "tl"
    rotate_with_shape: bool = False
    scheme_color: str | None = None
    sx_pct: int | None = None
    sy_pct: int | None = None

    def apply(self, shape) -> None:
        spPr = shape._element.find(qn("p:spPr"))
        if spPr is None:
            return
        effectLst = spPr.find(qn("a:effectLst"))
        if effectLst is not None:
            existing = effectLst.find(qn("a:outerShdw"))
            if existing is not None:
                effectLst.remove(existing)
        else:
            effectLst = etree.SubElement(spPr, qn("a:effectLst"))

        outerShdw = etree.SubElement(effectLst, qn("a:outerShdw"))
        outerShdw.set("blurRad", str(int(self.blur_pt * 12700)))
        outerShdw.set("dist", str(int(self.distance_pt * 12700)))
        outerShdw.set("dir", str(int(self.direction_deg * 60000)))
        outerShdw.set("algn", self.align)
        outerShdw.set("rotWithShape", "1" if self.rotate_with_shape else "0")
        if self.sx_pct is not None:
            outerShdw.set("sx", str(self.sx_pct * 1000))
        if self.sy_pct is not None:
            outerShdw.set("sy", str(self.sy_pct * 1000))

        if self.scheme_color:
            srgbClr = etree.SubElement(outerShdw, qn("a:schemeClr"))
            srgbClr.set("val", self.scheme_color)
        else:
            srgbClr = etree.SubElement(outerShdw, qn("a:srgbClr"))
            srgbClr.set("val", self.color.lstrip("#"))
        alpha_el = etree.SubElement(srgbClr, qn("a:alpha"))
        alpha_el.set("val", str(self.alpha_pct * 1000))


@dataclass
class SoftEdge:
    radius_pt: float = 2.0

    def apply(self, shape) -> None:
        spPr = shape._element.find(qn("p:spPr"))
        if spPr is None:
            return
        effectLst = spPr.find(qn("a:effectLst"))
        if effectLst is None:
            effectLst = etree.SubElement(spPr, qn("a:effectLst"))
        existing = effectLst.find(qn("a:softEdge"))
        if existing is not None:
            effectLst.remove(existing)
        softEdge = etree.SubElement(effectLst, qn("a:softEdge"))
        softEdge.set("rad", str(int(self.radius_pt * 12700)))


@dataclass
class Reflection:
    blur_pt: float = 4.0
    distance_pt: float = 2.0
    direction_deg: float = 90.0
    alpha_pct: int = 20
    end_alpha_pct: int = 0
    end_position: int = 50000

    def apply(self, shape) -> None:
        spPr = shape._element.find(qn("p:spPr"))
        if spPr is None:
            return
        effectLst = spPr.find(qn("a:effectLst"))
        if effectLst is None:
            effectLst = etree.SubElement(spPr, qn("a:effectLst"))
        existing = effectLst.find(qn("a:reflection"))
        if existing is not None:
            effectLst.remove(existing)
        reflection = etree.SubElement(effectLst, qn("a:reflection"))
        reflection.set("blurRad", str(int(self.blur_pt * 12700)))
        reflection.set("dist", str(int(self.distance_pt * 12700)))
        reflection.set("dir", str(int(self.direction_deg * 60000)))
        reflection.set("algn", "bl")
        reflection.set("rotWithShape", "0")
        srgbClr = etree.SubElement(reflection, qn("a:srgbClr"))
        srgbClr.set("val", "000000")
        alpha_el = etree.SubElement(srgbClr, qn("a:alpha"))
        alpha_el.set("val", str(self.alpha_pct * 1000))
        alphaOff = etree.SubElement(reflection, qn("a:alphaOff"))
        alphaOff.set("val", str(self.end_alpha_pct * 1000))


@dataclass
class Glow:
    radius_pt: float = 8.0
    color: str = "#2563EB"
    alpha_pct: int = 40

    def apply(self, shape) -> None:
        spPr = shape._element.find(qn("p:spPr"))
        if spPr is None:
            return
        effectLst = spPr.find(qn("a:effectLst"))
        if effectLst is None:
            effectLst = etree.SubElement(spPr, qn("a:effectLst"))
        existing = effectLst.find(qn("a:glow"))
        if existing is not None:
            effectLst.remove(existing)
        glow = etree.SubElement(effectLst, qn("a:glow"))
        glow.set("rad", str(int(self.radius_pt * 12700)))
        srgbClr = etree.SubElement(glow, qn("a:srgbClr"))
        srgbClr.set("val", self.color.lstrip("#"))
        alpha_el = etree.SubElement(srgbClr, qn("a:alpha"))
        alpha_el.set("val", str(self.alpha_pct * 1000))


def _remove_existing_fill(spPr: etree._Element) -> None:
    for tag in ["a:solidFill", "a:noFill", "a:gradFill", "a:pattFill"]:
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)


def set_solid_fill_with_alpha(shape, color: str, alpha_pct: int) -> None:
    spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return
    _remove_existing_fill(spPr)
    solidFill = etree.SubElement(spPr, qn("a:solidFill"))
    srgbClr = etree.SubElement(solidFill, qn("a:srgbClr"))
    srgbClr.set("val", color.lstrip("#"))
    alpha_el = etree.SubElement(srgbClr, qn("a:alpha"))
    alpha_el.set("val", str(alpha_pct * 1000))


def set_line_gradient(shape, color1: str, color2: str, width_pt: float = 2.0, angle: int = 5400000) -> None:
    spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return
    ln = spPr.find(qn("a:ln"))
    if ln is None:
        ln = etree.SubElement(spPr, qn("a:ln"))
    ln.set("w", str(int(width_pt * 12700)))
    for tag in ["a:solidFill", "a:noFill", "a:gradFill"]:
        el = ln.find(qn(tag))
        if el is not None:
            ln.remove(el)
    gradFill = etree.SubElement(ln, qn("a:gradFill"))
    gsLst = etree.SubElement(gradFill, qn("a:gsLst"))
    gs1 = etree.SubElement(gsLst, qn("a:gs"))
    gs1.set("pos", "0")
    etree.SubElement(gs1, qn("a:srgbClr")).set("val", color1.lstrip("#"))
    gs2 = etree.SubElement(gsLst, qn("a:gs"))
    gs2.set("pos", "100000")
    etree.SubElement(gs2, qn("a:srgbClr")).set("val", color2.lstrip("#"))
    lin = etree.SubElement(gradFill, qn("a:lin"))
    lin.set("ang", str(angle))
    lin.set("scaled", "1")


PRESET_GRADIENTS: dict[str, dict[str, Any]] = {
    "orange-warm": {"c1": "FF5500", "c2": "EE9003", "angle": 5400000},
    "blue-deep": {"c1": "1D78FA", "c2": "0165FF"},
    "red-bold": {"c1": "F13612", "c2": "D11801"},
    "gold-shine": {"c1": "F5AF19", "c2": "FFA300"},
    "sky-light": {"c1": "7BD3FF", "c2": "00BEFF"},
    "cyan-fresh": {"c1": "48E8E4", "c2": "31DCFD"},
    "dark-navy": {"c1": "2D3847", "c2": "121212"},
    "dark-overlay": {"c1": "060B18", "c2": "0A1E3D"},
    "purple-neon": {"c1": "8B5CF6", "c2": "6366F1"},
    "green-calm": {"c1": "4CAF50", "c2": "1B5E20"},
}


def make_gradient(preset: str, gradient_type: str = "linear", angle: int = 5400000) -> GradientFill:
    cfg = PRESET_GRADIENTS.get(preset, PRESET_GRADIENTS["blue-deep"])
    return GradientFill(
        stops=[
            GradientStop(color=cfg["c1"], position=0),
            GradientStop(color=cfg["c2"], position=100000),
        ],
        angle=angle,
        gradient_type=gradient_type,
    )


def apply_preset_gradient(shape, preset: str, gradient_type: str = "linear", angle: int = 5400000) -> None:
    grad = make_gradient(preset, gradient_type, angle)
    grad.apply(shape)


def apply_gradient(shape, color1: str, color2: str, gradient_type: str = "linear", angle: int = 5400000) -> None:
    grad = GradientFill(
        stops=[
            GradientStop(color=color1, position=0),
            GradientStop(color=color2, position=100000),
        ],
        angle=angle,
        gradient_type=gradient_type,
    )
    grad.apply(shape)


def apply_shadow(shape, blur_pt: float = 6.0, distance_pt: float = 3.0, direction_deg: float = 90.0, color: str = "#000000", alpha_pct: int = 25) -> None:
    shadow = Shadow(blur_pt=blur_pt, distance_pt=distance_pt, direction_deg=direction_deg, color=color, alpha_pct=alpha_pct)
    shadow.apply(shape)


def apply_scheme_shadow(shape, scheme_color: str = "accent1", blur_pt: float = 47.0, distance_pt: float = 21.0, direction_deg: float = 90.0, alpha_pct: int = 26, sx_pct: int | None = None, sy_pct: int | None = None) -> None:
    shadow = Shadow(
        blur_pt=blur_pt, distance_pt=distance_pt, direction_deg=direction_deg,
        alpha_pct=alpha_pct, scheme_color=scheme_color, sx_pct=sx_pct, sy_pct=sy_pct,
    )
    shadow.apply(shape)


def apply_luminance_gradient(shape, scheme_color: str = "accent1",
                              pos1: int = 18000, lum_mod1: int = 10000, lum_off1: int = 90000,
                              pos2: int = 73000, lum_mod2: int = 20000, lum_off2: int = 80000,
                              gradient_type: str = "path",
                              fill_to_rect: dict[str, str] | None = None,
                              angle: int = 5400000) -> None:
    grad = GradientFill(
        stops=[
            GradientStop(scheme_color=scheme_color, position=pos1, lum_mod=lum_mod1, lum_off=lum_off1),
            GradientStop(scheme_color=scheme_color, position=pos2, lum_mod=lum_mod2, lum_off=lum_off2),
        ],
        angle=angle,
        gradient_type=gradient_type,
        fill_to_rect=fill_to_rect or {"r": "100000", "b": "100000"},
    )
    grad.apply(shape)


def apply_3stop_gradient(shape, c1: str, c2: str, c3: str,
                          pos1: int = 0, pos2: int = 50000, pos3: int = 100000,
                          gradient_type: str = "linear", angle: int = 5400000) -> None:
    grad = GradientFill(
        stops=[
            GradientStop(color=c1, position=pos1),
            GradientStop(color=c2, position=pos2),
            GradientStop(color=c3, position=pos3),
        ],
        angle=angle,
        gradient_type=gradient_type,
    )
    grad.apply(shape)


def apply_cross_gradient(shape, scheme_accent1: str = "accent1", scheme_accent2: str = "accent2",
                          gradient_type: str = "path",
                          fill_to_rect: dict[str, str] | None = None) -> None:
    grad = GradientFill(
        stops=[
            GradientStop(scheme_color=scheme_accent2, position=0, lum_mod=60000, lum_off=40000),
            GradientStop(scheme_color=scheme_accent1, position=80000),
        ],
        gradient_type=gradient_type,
        fill_to_rect=fill_to_rect or {"r": "100000", "b": "100000"},
    )
    grad.apply(shape)


def apply_bg_gradient(shape, scheme_accent: str = "accent1",
                       fill_to_rect: dict[str, str] | None = None) -> None:
    grad = GradientFill(
        stops=[
            GradientStop(scheme_color=scheme_accent, position=13000, lum_mod=20000, lum_off=80000),
            GradientStop(scheme_color="bg1", position=71000),
            GradientStop(scheme_color="bg1", position=100000),
        ],
        gradient_type="path",
        fill_to_rect=fill_to_rect or {"l": "50000", "t": "130000", "r": "50000", "b": "-30000"},
    )
    grad.apply(shape)


def apply_soft_edge(shape, radius_pt: float = 2.0) -> None:
    se = SoftEdge(radius_pt=radius_pt)
    se.apply(shape)


def apply_reflection(shape, alpha_pct: int = 20) -> None:
    ref = Reflection(alpha_pct=alpha_pct)
    ref.apply(shape)


def apply_glow(shape, radius_pt: float = 8.0, color: str = "#2563EB", alpha_pct: int = 40) -> None:
    glow = Glow(radius_pt=radius_pt, color=color, alpha_pct=alpha_pct)
    glow.apply(shape)


def apply_letter_spacing(run, tracking_em: float, font_size_pt: int) -> None:
    """Set letter spacing on a run via OOXML a:spc element.

    tracking_em: EM ratio, e.g. -0.02 (tight), +0.08 (wide for ALL CAPS)
    font_size_pt: font size in points (OOXML spc val is font-size-dependent)
    val = tracking_em * font_size_pt * 100 (hundredths of a point)
    """
    if tracking_em == 0.0:
        return
    rPr = run._r.get_or_add_rPr()
    spc = rPr.find(qn("a:spc"))
    if spc is None:
        spc = etree.SubElement(rPr, qn("a:spc"))
    spc.set("val", str(int(tracking_em * font_size_pt * 100)))
