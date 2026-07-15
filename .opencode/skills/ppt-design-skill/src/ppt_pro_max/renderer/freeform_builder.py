"""Freeform Builder — custom geometric paths and decorative connectors.

Creates Freeform shapes (任意多边形) via OOXML custGeom XML, enabling:
  - Curved/straight connector lines between nodes
  - Arrow shapes and decorative pointers
  - Custom geometric backgrounds (chevrons, ribbons, etc.)
  - Bezier curve paths for organic connections

All coordinates are in EMU (914400 per inch) within the shape's local
coordinate system (path w/h), then mapped via xfrm to slide position.
"""

from __future__ import annotations

from typing import Any

from pptx.oxml.ns import qn
from lxml import etree


EMU_PER_INCH = 914400


class FreeformBuilder:

    def __init__(self):
        self._paths: list[list[dict[str, Any]]] = []
        self._current_path: list[dict[str, Any]] = []

    def move_to(self, x: float, y: float) -> FreeformBuilder:
        self._current_path.append({"cmd": "moveTo", "x": x, "y": y})
        return self

    def line_to(self, x: float, y: float) -> FreeformBuilder:
        self._current_path.append({"cmd": "lnTo", "x": x, "y": y})
        return self

    def cubic_bezier_to(self, x1: float, y1: float, x2: float, y2: float, x3: float, y3: float) -> FreeformBuilder:
        self._current_path.append({"cmd": "cubicBezTo", "x1": x1, "y1": y1, "x2": x2, "y2": y2, "x3": x3, "y3": y3})
        return self

    def close(self) -> FreeformBuilder:
        self._current_path.append({"cmd": "close"})
        return self

    def new_path(self) -> FreeformBuilder:
        if self._current_path:
            self._paths.append(self._current_path)
        self._current_path = []
        return self

    def build(
        self, slide, x: float, y: float, w: float, h: float,
        fill_color: str | None = None, line_color: str | None = None,
        line_width_pt: float = 1.5, no_fill: bool = True,
    ) -> object:
        if self._current_path:
            self._paths.append(self._current_path)
            self._current_path = []

        sp_tree = slide.shapes._spTree
        sp = etree.SubElement(sp_tree, qn("p:sp"))

        nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
        cNvPr = etree.SubElement(nvSpPr, qn("p:cNvPr"))
        cNvPr.set("id", str(self._next_id(slide)))
        cNvPr.set("name", "Freeform")
        etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
        etree.SubElement(nvSpPr, qn("p:nvPr"))

        spPr = etree.SubElement(sp, qn("p:spPr"))
        xfrm = etree.SubElement(spPr, qn("a:xfrm"))
        off = etree.SubElement(xfrm, qn("a:off"))
        off.set("x", str(int(x * EMU_PER_INCH)))
        off.set("y", str(int(y * EMU_PER_INCH)))
        ext = etree.SubElement(xfrm, qn("a:ext"))
        ext.set("cx", str(int(w * EMU_PER_INCH)))
        ext.set("cy", str(int(h * EMU_PER_INCH)))

        custGeom = etree.SubElement(spPr, qn("a:custGeom"))
        etree.SubElement(custGeom, qn("a:avLst"))
        etree.SubElement(custGeom, qn("a:gdLst"))

        path_w = int(w * EMU_PER_INCH)
        path_h = int(h * EMU_PER_INCH)
        pathLst = etree.SubElement(custGeom, qn("a:pathLst"))

        for path_cmds in self._paths:
            path_el = etree.SubElement(pathLst, qn("a:path"))
            path_el.set("w", str(path_w))
            path_el.set("h", str(path_h))
            for cmd in path_cmds:
                if cmd["cmd"] == "moveTo":
                    moveTo = etree.SubElement(path_el, qn("a:moveTo"))
                    pt = etree.SubElement(moveTo, qn("a:pt"))
                    pt.set("x", str(int(cmd["x"] * EMU_PER_INCH)))
                    pt.set("y", str(int(cmd["y"] * EMU_PER_INCH)))
                elif cmd["cmd"] == "lnTo":
                    lnTo = etree.SubElement(path_el, qn("a:lnTo"))
                    pt = etree.SubElement(lnTo, qn("a:pt"))
                    pt.set("x", str(int(cmd["x"] * EMU_PER_INCH)))
                    pt.set("y", str(int(cmd["y"] * EMU_PER_INCH)))
                elif cmd["cmd"] == "cubicBezTo":
                    bez = etree.SubElement(path_el, qn("a:cubicBezTo"))
                    for prefix in ["1", "2", "3"]:
                        pt = etree.SubElement(bez, qn("a:pt"))
                        pt.set("x", str(int(cmd[f"x{prefix}"] * EMU_PER_INCH)))
                        pt.set("y", str(int(cmd[f"y{prefix}"] * EMU_PER_INCH)))
                elif cmd["cmd"] == "close":
                    etree.SubElement(path_el, qn("a:close"))

        if no_fill or fill_color is None:
            etree.SubElement(spPr, qn("a:noFill"))
        else:
            solidFill = etree.SubElement(spPr, qn("a:solidFill"))
            etree.SubElement(solidFill, qn("a:srgbClr")).set("val", fill_color.lstrip("#"))

        ln = etree.SubElement(spPr, qn("a:ln"))
        ln.set("w", str(int(line_width_pt * 12700)))
        if line_color:
            solidFill = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(solidFill, qn("a:srgbClr")).set("val", line_color.lstrip("#"))
        else:
            etree.SubElement(ln, qn("a:noFill"))

        self._paths = []
        self._current_path = []
        return sp

    @staticmethod
    def _next_id(slide) -> int:
        max_id = 1
        for shape in slide.shapes:
            try:
                sid = shape.shape_id
                if sid > max_id:
                    max_id = sid
            except Exception:
                pass
        return max_id + 1


def make_arrow_connector(
    slide, x1: float, y1: float, x2: float, y2: float,
    color: str = "#FF5500", width_pt: float = 2.0,
    arrow_start: bool = False, arrow_end: bool = True,
) -> object:
    builder = FreeformBuilder()
    builder.move_to(x1, y1)
    builder.line_to(x2, y2)
    elem = builder.build(slide, min(x1, x2), min(y1, y2),
                         abs(x2 - x1) or 0.01, abs(y2 - y1) or 0.01,
                         no_fill=True, line_color=color, line_width_pt=width_pt)
    if arrow_end:
        _add_arrowhead(elem, "end", color)
    if arrow_start:
        _add_arrowhead(elem, "start", color)
    return elem


def make_curved_connector(
    slide, x1: float, y1: float, x2: float, y2: float,
    color: str = "#FF5500", width_pt: float = 2.0,
    curvature: float = 0.3, arrow_end: bool = True,
) -> object:
    mx = (x1 + x2) / 2
    my = (y1 + y2) / 2
    dx = x2 - x1
    dy = y2 - y1
    cx1 = mx -dy * curvature
    dx * curvature
    cx1 = mx - dy * curvature
    cy1 = my + dx * curvature

    min_x = min(x1, x2, cx1) - 0.1
    min_y = min(y1, y2, cy1) - 0.1
    max_x = max(x1, x2, cx1) + 0.1
    max_y = max(y1, y2, cy1) + 0.1

    builder = FreeformBuilder()
    builder.move_to(x1 - min_x, y1 - min_y)
    builder.cubic_bezier_to(
        cx1 - min_x, cy1 - min_y,
        cx1 - min_x, cy1 - min_y,
        x2 - min_x, y2 - min_y,
    )
    elem = builder.build(slide, min_x, min_y, max_x - min_x, max_y - min_y,
                         no_fill=True, line_color=color, line_width_pt=width_pt)
    if arrow_end:
        _add_arrowhead(elem, "end", color)
    return elem


def make_chevron_shape(
    slide, x: float, y: float, w: float, h: float,
    fill_color: str = "#1D78FA", line_color: str | None = None,
) -> object:
    builder = FreeformBuilder()
    notch = w * 0.15
    builder.move_to(0, 0)
    builder.line_to(w - notch, 0)
    builder.line_to(w, h / 2)
    builder.line_to(w - notch, h)
    builder.line_to(0, h)
    builder.line_to(notch, h / 2)
    builder.close()
    return builder.build(slide, x, y, w, h, fill_color=fill_color, line_color=line_color, no_fill=False)


def make_ribbon_shape(
    slide, x: float, y: float, w: float, h: float,
    fill_color: str = "#FF5500", tab_height: float = 0.3,
) -> object:
    builder = FreeformBuilder()
    builder.move_to(0, 0)
    builder.line_to(w, 0)
    builder.line_to(w, h - tab_height)
    builder.line_to(w * 0.75, h - tab_height)
    builder.line_to(w * 0.65, h)
    builder.line_to(w * 0.55, h - tab_height)
    builder.line_to(0, h - tab_height)
    builder.close()
    return builder.build(slide, x, y, w, h, fill_color=fill_color, no_fill=False)


def _add_arrowhead(elem, end: str, color: str) -> None:
    spPr = elem.find(qn("p:spPr"))
    if spPr is None:
        return
    ln = spPr.find(qn("a:ln"))
    if ln is None:
        return
    if end == "end":
        tailEnd = etree.SubElement(ln, qn("a:tailEnd"))
        tailEnd.set("type", "triangle")
        tailEnd.set("w", "med")
        tailEnd.set("len", "med")
    elif end == "start":
        headEnd = etree.SubElement(ln, qn("a:headEnd"))
        headEnd.set("type", "triangle")
        headEnd.set("w", "med")
        headEnd.set("len", "med")
