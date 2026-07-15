"""Color Depth System — OKLCH tint/shade scales, alpha levels, surface colors."""

from __future__ import annotations

import math


def _hex_to_oklch(hex_color: str) -> tuple[float, float, float]:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
    r = r / 12.92 if r <= 0.04045 else ((r + 0.055) / 1.055) ** 2.4
    g = g / 12.92 if g <= 0.04045 else ((g + 0.055) / 1.055) ** 2.4
    b = b / 12.92 if b <= 0.04045 else ((b + 0.055) / 1.055) ** 2.4
    l_ = 0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b
    m_ = 0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b
    s_ = 0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b
    l_ = l_ ** (1/3) if l_ >= 0 else -((-l_) ** (1/3))
    m_ = m_ ** (1/3) if m_ >= 0 else -((-m_) ** (1/3))
    s_ = s_ ** (1/3) if s_ >= 0 else -((-s_) ** (1/3))
    L = 0.2104542553 * l_ + 0.7936177850 * m_ - 0.0040720468 * s_
    a = 1.9779984951 * l_ - 2.4285922050 * m_ + 0.4505937099 * s_
    b_ok = 0.0259040371 * l_ + 0.7827717662 * m_ - 0.8086757660 * s_
    C = (a ** 2 + b_ok ** 2) ** 0.5
    H = math.degrees(math.atan2(b_ok, a)) % 360
    return L, C, H


def _oklch_to_hex(L: float, C: float, H: float) -> str:
    H_rad = H * math.pi / 180
    a = C * math.cos(H_rad)
    b_ok = C * math.sin(H_rad)
    l_ = L + 0.3963377774 * a + 0.2158037573 * b_ok
    m_ = L - 0.1055613458 * a - 0.0638541728 * b_ok
    s_ = L - 0.0894841775 * a - 1.2914855480 * b_ok
    l_ = l_ ** 3 if l_ >= 0 else -((-l_) ** 3)
    m_ = m_ ** 3 if m_ >= 0 else -((-m_) ** 3)
    s_ = s_ ** 3 if s_ >= 0 else -((-s_) ** 3)
    r = +4.0767416621 * l_ - 3.3077115913 * m_ + 0.2309699292 * s_
    g = -1.2684380046 * l_ + 2.6097574011 * m_ - 0.3413193965 * s_
    b = -0.0041960863 * l_ - 0.7034186147 * m_ + 1.7076147010 * s_
    r = max(0, min(1, 12.92 * r if r <= 0.0031308 else 1.055 * r ** (1/2.4) - 0.055))
    g = max(0, min(1, 12.92 * g if g <= 0.0031308 else 1.055 * g ** (1/2.4) - 0.055))
    b = max(0, min(1, 12.92 * b if b <= 0.0031308 else 1.055 * b ** (1/2.4) - 0.055))
    return f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}"


def generate_color_scale(base_hex: str, levels: int = 11) -> dict[str, str]:
    _ = levels
    oklch_l, oklch_c, oklch_h = _hex_to_oklch(base_hex)
    scale = {}
    for level in [50, 100, 200, 300, 400, 500, 600, 700, 800, 900, 950]:
        if level <= 500:
            t = (500 - level) / 500
            new_l = oklch_l + (1.0 - oklch_l) * t * 0.9
            new_c = oklch_c * (1 - t * 0.5)
        else:
            t = (level - 500) / 500
            new_l = oklch_l * (1 - t * 0.8)
            new_c = oklch_c * (1 - t * 0.3)
        scale[str(level)] = _oklch_to_hex(new_l, new_c, oklch_h)
    return scale


def alpha_color(shape, hex_color: str, opacity_pct: int) -> None:
    from ppt_pro_max.renderer.visual_effects import set_solid_fill_with_alpha
    set_solid_fill_with_alpha(shape, hex_color, opacity_pct)


ALPHA_LEVELS = {
    "ghost": 4,
    "subtle": 8,
    "light": 15,
    "medium": 40,
    "strong": 65,
}
