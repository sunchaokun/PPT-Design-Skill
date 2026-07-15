"""Effects — shadows, glow, gradients, and other visual effects via XML."""

from __future__ import annotations

from typing import Any

try:
    from pptx.oxml.ns import qn
    from pptx.dml.color import RGBColor

    _PPTX_AVAILABLE = True
except ImportError:
    _PPTX_AVAILABLE = False


class Effects:
    def add_shadow(self, shape, blur_pt: int = 8, offset_pt: int = 4, color: str = "#000000", alpha: int = 40) -> None:
        if not _PPTX_AVAILABLE:
            return
        try:
            spPr = shape._element.find(qn("p:spPr"))
            if spPr is None:
                spPr = shape._element.find(qn("a:spPr"))
            if spPr is None:
                return

            effect_lst = spPr.find(qn("a:effectLst"))
            if effect_lst is None:
                effect_lst = spPr.makeelement(qn("a:effectLst"), {})
                spPr.append(effect_lst)

            outer_shdw = effect_lst.makeelement(qn("a:outerShdw"), {
                "blurRad": str(blur_pt * 12700),
                "dist": str(offset_pt * 12700),
                "dir": "5400000",
                "algn": "tl",
                "rotWithShape": "0",
            })

            srgb = outer_shdw.makeelement(qn("a:srgbClr"), {"val": color.lstrip("#")})
            alpha_elem = srgb.makeelement(qn("a:alpha"), {"val": str(alpha * 1000)})
            srgb.append(alpha_elem)
            outer_shdw.append(srgb)
            effect_lst.append(outer_shdw)
        except Exception:
            pass

    def add_gradient_background(self, slide, color1: str, color2: str, angle: int = 5400000) -> None:
        if not _PPTX_AVAILABLE:
            return
        try:
            bg = slide.background
            fill = bg.fill
            fill.gradient()
            fill.gradient_stops[0].color.rgb = RGBColor.from_string(color1.lstrip("#"))
            fill.gradient_stops[1].color.rgb = RGBColor.from_string(color2.lstrip("#"))
        except Exception:
            pass

    def add_glow(self, shape, radius_pt: int = 8, color: str = "#2563EB", alpha: int = 40) -> None:
        if not _PPTX_AVAILABLE:
            return
        try:
            spPr = shape._element.find(qn("p:spPr"))
            if spPr is None:
                spPr = shape._element.find(qn("a:spPr"))
            if spPr is None:
                return

            effect_lst = spPr.find(qn("a:effectLst"))
            if effect_lst is None:
                effect_lst = spPr.makeelement(qn("a:effectLst"), {})
                spPr.append(effect_lst)

            glow = effect_lst.makeelement(qn("a:glow"), {
                "rad": str(radius_pt * 12700),
            })
            srgb = glow.makeelement(qn("a:srgbClr"), {"val": color.lstrip("#")})
            alpha_elem = srgb.makeelement(qn("a:alpha"), {"val": str(alpha * 1000)})
            srgb.append(alpha_elem)
            glow.append(srgb)
            effect_lst.append(glow)
        except Exception:
            pass
