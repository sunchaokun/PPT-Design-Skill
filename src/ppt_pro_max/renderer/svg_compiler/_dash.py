"""SVG stroke dash, linecap, and linejoin support.

Adds:
- stroke-dasharray parsing (comma/space separated, "none")
- stroke-linecap → PPT line cap (flat/round/square)
- stroke-linejoin → PPT line join (round/bevel/miter)
- stroke-miterlimit
- Dash pattern application via OOXML XML injection (python-pptx has no dash API)
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from lxml import etree


@dataclass
class StrokeStyle:
    dash_array: list[float] | None = None
    linecap: str = "flat"
    linejoin: str = "miter"
    miterlimit: float = 4.0
    width: float = 1.0


_LINECAP_MAP = {
    "butt": "flat",
    "round": "round",
    "square": "square",
}

_LINEJOIN_MAP = {
    "miter": "miter",
    "round": "round",
    "bevel": "bevel",
}


def parse_stroke_style(el) -> StrokeStyle:
    da = el.get("stroke-dasharray")
    dash_array = None
    if da and da.lower() != "none":
        parts = re.split(r"[,\s]+", da.strip())
        try:
            dash_array = [float(p) for p in parts if p]
        except ValueError:
            dash_array = None

    cap = el.get("stroke-linecap", "butt").lower()
    linecap = _LINECAP_MAP.get(cap, "flat")

    join = el.get("stroke-linejoin", "miter").lower()
    linejoin = _LINEJOIN_MAP.get(join, "round")

    ml = el.get("stroke-miterlimit")
    miterlimit = float(ml) if ml else 4.0

    sw = el.get("stroke-width", "1")
    width = float(re.sub(r"[^\d.]", "", sw)) if sw else 1.0

    return StrokeStyle(
        dash_array=dash_array,
        linecap=linecap,
        linejoin=linejoin,
        miterlimit=miterlimit,
        width=width,
    )


_PPT_DASH_PRESETS: dict[tuple[int, ...], str] = {
    (2, 2): "dash",
    (4, 4): "dash",
    (6, 3): "lgDash",
    (3, 2): "dash",
    (1, 2): "dot",
    (1, 3): "dot",
    (2, 3, 3, 3): "dashDot",
    (4, 2, 1, 2): "dashDot",
    (8, 3, 2, 3): "lgDashDot",
}


def _match_dash_preset(dash_array: list[float]) -> str | None:
    if not dash_array:
        return None
    if len(dash_array) % 2 == 1:
        dash_array = dash_array * 2

    total = sum(dash_array)
    if total == 0:
        return None

    if len(dash_array) == 2:
        d1 = max(dash_array[0], 1)
        d2 = max(dash_array[1], 1)
        if d2 >= d1 * 2 or d1 <= 1.5:
            return "dot"
        if d1 >= 6:
            return "lgDash"
        return "dash"

    normalized = tuple(max(1, round(d / total * 8)) for d in dash_array[:8])
    return _PPT_DASH_PRESETS.get(normalized)


_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def apply_stroke_style(shape, style: StrokeStyle) -> None:
    if not hasattr(shape, "_element"):
        return

    sp = shape._element
    sp_pr = sp.find(f"{{{_NS_P}}}spPr")
    if sp_pr is None:
        sp_pr = sp.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}spPr")
    if sp_pr is None:
        sp_pr = sp.find(".//*[local-name()='spPr']")
    if sp_pr is None:
        return

    ln = sp_pr.find(f"{{{_NS_A}}}ln")
    if ln is None:
        ln = etree.SubElement(sp_pr, f"{{{_NS_A}}}ln")

    if style.linecap != "flat":
        cap_val = style.linecap
        ln.set("cap", cap_val)

    if style.linejoin != "round":
        existing_jn = ln.find(f"{{{_NS_A}}}round")
        if existing_jn is not None:
            ln.remove(existing_jn)
        if style.linejoin == "miter":
            jn = etree.SubElement(ln, f"{{{_NS_A}}}miter")
            jn.set("lim", str(int(style.miterlimit * 1000)))
        elif style.linejoin == "bevel":
            etree.SubElement(ln, f"{{{_NS_A}}}bevel")

    if style.dash_array:
        preset = _match_dash_preset(style.dash_array)
        if preset:
            existing_ds = ln.find(f"{{{_NS_A}}}prstDash")
            if existing_ds is not None:
                ln.remove(existing_ds)
            ds = etree.SubElement(ln, f"{{{_NS_A}}}prstDash")
            ds.set("val", preset)
        else:
            existing_cs = ln.find(f"{{{_NS_A}}}custDash")
            if existing_cs is not None:
                ln.remove(existing_cs)
            cs = etree.SubElement(ln, f"{{{_NS_A}}}custDash")
            arr = style.dash_array if len(style.dash_array) % 2 == 0 else style.dash_array * 2
            for i in range(0, len(arr), 2):
                d = etree.SubElement(cs, f"{{{_NS_A}}}ds")
                d.set("d", str(int(arr[i] * 1000)))
                d.set("sp", str(int(arr[i + 1] * 1000)) if i + 1 < len(arr) else "1000")
