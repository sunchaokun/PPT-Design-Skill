"""TextMeasurer — estimate text size in PPT without rendering."""

from __future__ import annotations

import math


_CJK_RANGE = (
    (0x4E00, 0x9FFF),
    (0x3400, 0x4DBF),
    (0x3000, 0x303F),
    (0xFF00, 0xFFEF),
)


_CHAR_WIDTH_TABLE: dict[int, float] = {
    10: 0.06,
    11: 0.066,
    12: 0.072,
    13: 0.078,
    14: 0.084,
    16: 0.096,
    18: 0.108,
    20: 0.12,
    24: 0.144,
    28: 0.168,
    32: 0.192,
}

_CJK_WIDTH_RATIO = 2.0

_PADDING_V = 0.08
_PADDING_H = 0.15


def _is_cjk(ch: str) -> bool:
    cp = ord(ch)
    for lo, hi in _CJK_RANGE:
        if lo <= cp <= hi:
            return True
    return False


def estimate_text_size(
    text: str,
    font_size_pt: int,
    max_width: float,
    font_family: str = "Calibri",
    line_spacing: float = 1.2,
) -> tuple[float, float]:
    if not text:
        return 0.0, 0.0

    char_width = _CHAR_WIDTH_TABLE.get(font_size_pt, font_size_pt * 0.006)

    font_ratio = 1.0
    fl = font_family.lower()
    if "serif" in fl or "georgia" in fl:
        font_ratio = 1.1
    elif "mono" in fl or "consolas" in fl:
        font_ratio = 1.2

    total_width = 0.0
    for ch in text:
        if _is_cjk(ch):
            total_width += char_width * _CJK_WIDTH_RATIO * font_ratio
        else:
            total_width += char_width * font_ratio

    line_height = font_size_pt / 72 * line_spacing
    if max_width <= 0:
        max_width = 10.0

    if total_width <= max_width:
        lines = 1
    else:
        lines = math.ceil(total_width / max_width)

    actual_width = min(total_width + _PADDING_H * 2, max_width)
    actual_height = lines * line_height + _PADDING_V * 2

    return actual_width, actual_height


def estimate_node_size(
    label: str,
    font_size_pt: int,
    max_width: float,
    padding: float = 0.15,
    font_family: str = "Calibri",
) -> tuple[float, float]:
    text_w, text_h = estimate_text_size(label, font_size_pt, max_width, font_family)
    node_w = text_w + padding * 2
    node_h = text_h + padding * 2
    return node_w, node_h
