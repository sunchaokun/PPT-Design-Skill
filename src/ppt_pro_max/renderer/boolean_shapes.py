"""Boolean Shape Operations — Shapely geometry → OOXML custGeom.

Internal module used by build_helpers. LLM typically uses the
convenience functions in build_helpers (spotlight, bool_donut, etc.)
instead of calling these primitives directly.

Advanced users can compose custom shapes:
    from ppt_pro_max.renderer.boolean_shapes import *
    mask = bool_subtract(poly_rect(0,0,6,4), poly_circle(3,2,1.5))
    bool_shape(mask, slide, 1, 2, 6, 4, fill='#000000', alpha=70)
"""

from __future__ import annotations

import math

from pptx.oxml.ns import qn
from lxml import etree

EMU_PER_INCH = 914400

HAS_SHAPELY = False
try:
    from shapely.geometry import Polygon, MultiPolygon
    from shapely.ops import unary_union
    HAS_SHAPELY = True
except ImportError:
    pass


def poly_rect(x: float, y: float, w: float, h: float):
    if not HAS_SHAPELY:
        return None
    return Polygon([(x, y), (x + w, y), (x + w, y + h), (x, y + h)])


def poly_circle(cx: float, cy: float, r: float, segments: int = 72):
    if not HAS_SHAPELY:
        return None
    return Polygon([
        (cx + r * math.cos(2 * math.pi * i / segments),
         cy + r * math.sin(2 * math.pi * i / segments))
        for i in range(segments)
    ])


def poly_rounded_rect(x: float, y: float, w: float, h: float,
                       radius: float = 0.1, segs_per_corner: int = 8):
    if not HAS_SHAPELY:
        return None
    r = min(radius, w / 2, h / 2)
    pts = []
    corners = [
        (x + r, y, 0, 1),
        (x + w - r, y, 1, 0),
        (x + w, y + r, 0, 1),
        (x + w, y + h - r, -1, 0),
        (x + w - r, y + h, 0, -1),
        (x + r, y + h, -1, 0),
        (x, y + h - r, 0, -1),
        (x, y + r, 1, 0),
    ]
    for cx_c, cy_c, dx, dy in corners:
        for i in range(segs_per_corner + 1):
            angle = (math.pi / 2) * i / segs_per_corner
            pts.append((cx_c + dx * r * math.cos(angle) - dy * r * math.sin(angle),
                        cy_c + dy * r * math.cos(angle) + dx * r * math.sin(angle)))
    return Polygon(pts)


def poly_star(cx: float, cy: float, r: float,
              points: int = 5, inner_ratio: float = 0.4):
    if not HAS_SHAPELY:
        return None
    r_inner = r * inner_ratio
    pts = []
    for i in range(points * 2):
        angle = math.pi * i / points - math.pi / 2
        radius = r if i % 2 == 0 else r_inner
        pts.append((cx + radius * math.cos(angle),
                     cy + radius * math.sin(angle)))
    return Polygon(pts)


def poly_regular(cx: float, cy: float, r: float, sides: int):
    if not HAS_SHAPELY:
        return None
    return Polygon([
        (cx + r * math.cos(2 * math.pi * i / sides - math.pi / 2),
         cy + r * math.sin(2 * math.pi * i / sides - math.pi / 2))
        for i in range(sides)
    ])


def poly_points(coords: list[tuple[float, float]]):
    if not HAS_SHAPELY:
        return None
    return Polygon(coords)


def bool_subtract(a, b):
    if not HAS_SHAPELY or a is None or b is None:
        return None
    return a.difference(b)


def bool_union(*shapes):
    if not HAS_SHAPELY:
        return None
    valid = [s for s in shapes if s is not None]
    if not valid:
        return None
    if len(valid) == 1:
        return valid[0]
    return unary_union(valid)


def bool_intersect(a, b):
    if not HAS_SHAPELY or a is None or b is None:
        return None
    return a.intersection(b)


def bool_symdiff(a, b):
    if not HAS_SHAPELY or a is None or b is None:
        return None
    return a.symmetric_difference(b)


def _polygon_to_path_cmds(poly, scale: float = 1.0):
    coords = list(poly.exterior.coords)
    cmds = []
    cmds.append({"cmd": "moveTo", "x": coords[0][0] * scale, "y": coords[0][1] * scale})
    for x, y in coords[1:]:
        cmds.append({"cmd": "lnTo", "x": x * scale, "y": y * scale})
    cmds.append({"cmd": "close"})
    return cmds


def _polygon_to_paths_with_holes(poly, scale: float = 1.0):
    paths = []
    coords = list(poly.exterior.coords)
    cmds = [{"cmd": "moveTo", "x": coords[0][0] * scale, "y": coords[0][1] * scale}]
    for x, y in coords[1:]:
        cmds.append({"cmd": "lnTo", "x": x * scale, "y": y * scale})
    cmds.append({"cmd": "close"})
    paths.append(cmds)
    for interior in poly.interiors:
        hole_coords = list(interior.coords)
        hole_cmds = [{"cmd": "moveTo", "x": hole_coords[0][0] * scale, "y": hole_coords[0][1] * scale}]
        for x, y in hole_coords[1:]:
            hole_cmds.append({"cmd": "lnTo", "x": x * scale, "y": y * scale})
        hole_cmds.append({"cmd": "close"})
        paths.append(hole_cmds)
    return paths


def _multipolygon_to_paths(geom, scale: float = 1.0):
    if not HAS_SHAPELY or geom is None:
        return []
    paths = []
    if isinstance(geom, Polygon):
        paths.extend(_polygon_to_paths_with_holes(geom, scale))
    elif isinstance(geom, MultiPolygon):
        for poly in geom.geoms:
            paths.extend(_polygon_to_paths_with_holes(poly, scale))
    return paths


def _build_custGeom_shape(slide, paths, x_in, y_in, w_in, h_in,
                          fill_color="#4472C4", line_color=None,
                          alpha=None):
    sp_tree = slide.shapes._spTree
    sp = etree.SubElement(sp_tree, qn("p:sp"))

    nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
    cNvPr = etree.SubElement(nvSpPr, qn("p:cNvPr"))
    max_id = 1
    for sh in slide.shapes:
        try:
            if sh.shape_id > max_id:
                max_id = sh.shape_id
        except Exception:
            pass
    cNvPr.set("id", str(max_id + 1))
    cNvPr.set("name", "BooleanShape")
    etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
    etree.SubElement(nvSpPr, qn("p:nvPr"))

    spPr = etree.SubElement(sp, qn("p:spPr"))
    xfrm = etree.SubElement(spPr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(int(x_in * EMU_PER_INCH)))
    off.set("y", str(int(y_in * EMU_PER_INCH)))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(int(w_in * EMU_PER_INCH)))
    ext.set("cy", str(int(h_in * EMU_PER_INCH)))

    custGeom = etree.SubElement(spPr, qn("a:custGeom"))
    etree.SubElement(custGeom, qn("a:avLst"))
    etree.SubElement(custGeom, qn("a:gdLst"))
    pathLst = etree.SubElement(custGeom, qn("a:pathLst"))

    path_w = int(w_in * EMU_PER_INCH)
    path_h = int(h_in * EMU_PER_INCH)
    off_x_emu = int(x_in * EMU_PER_INCH)
    off_y_emu = int(y_in * EMU_PER_INCH)

    for path_cmds in paths:
        path_el = etree.SubElement(pathLst, qn("a:path"))
        path_el.set("w", str(path_w))
        path_el.set("h", str(path_h))
        for cmd in path_cmds:
            if cmd["cmd"] == "moveTo":
                moveTo = etree.SubElement(path_el, qn("a:moveTo"))
                pt = etree.SubElement(moveTo, qn("a:pt"))
                pt.set("x", str(int(cmd["x"] * EMU_PER_INCH) - off_x_emu))
                pt.set("y", str(int(cmd["y"] * EMU_PER_INCH) - off_y_emu))
            elif cmd["cmd"] == "lnTo":
                lnTo = etree.SubElement(path_el, qn("a:lnTo"))
                pt = etree.SubElement(lnTo, qn("a:pt"))
                pt.set("x", str(int(cmd["x"] * EMU_PER_INCH) - off_x_emu))
                pt.set("y", str(int(cmd["y"] * EMU_PER_INCH) - off_y_emu))
            elif cmd["cmd"] == "close":
                etree.SubElement(path_el, qn("a:close"))

    solidFill = etree.SubElement(spPr, qn("a:solidFill"))
    srgb = etree.SubElement(solidFill, qn("a:srgbClr"))
    srgb.set("val", fill_color.lstrip("#"))
    if alpha is not None and 0 <= alpha <= 100:
        a_elem = etree.SubElement(srgb, qn("a:alpha"))
        a_elem.set("val", str(alpha * 1000))

    ln = etree.SubElement(spPr, qn("a:ln"))
    if line_color:
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", line_color.lstrip("#"))
    else:
        etree.SubElement(ln, qn("a:noFill"))

    return sp


def _boolean_to_slide(slide, geom, x_in, y_in, w_in, h_in,
                       fill_color="#4472C4", line_color=None, alpha=None):
    paths = _multipolygon_to_paths(geom, scale=1.0)
    if not paths:
        return None
    return _build_custGeom_shape(slide, paths, x_in, y_in, w_in, h_in,
                                  fill_color=fill_color, line_color=line_color,
                                  alpha=alpha)


def bool_shape(geometry, slide, x, y, w, h, fill=None, line=None, C=None,
               alpha=None):
    if geometry is None:
        return None
    from ppt_pro_max.build_helpers import _resolve_color
    fill_hex = _resolve_color(fill, C) if fill else '#4472C4'
    line_hex = _resolve_color(line, C) if line else None
    return _boolean_to_slide(slide, geometry, x, y, w, h,
                              fill_color=fill_hex, line_color=line_hex,
                              alpha=alpha)


def bool_image(geometry, slide, x, y, w, h, image_path, border_color=None):
    if geometry is None:
        return None
    paths = _multipolygon_to_paths(geometry, scale=1.0)
    if not paths:
        return None
    sp = _build_custGeom_shape(slide, paths, x, y, w, h,
                               fill_color="#FFFFFF", line_color=border_color)
    if sp is None:
        return None
    from ppt_pro_max.renderer.blip_fill import fill_shape_with_image
    spPr = sp.find(qn("p:spPr"))
    solidFill = spPr.find(qn("a:solidFill"))
    if solidFill is not None:
        spPr.remove(solidFill)
    class _ShapeProxy:
        def __init__(self, element, slide_ref):
            self._element = element
            self._slide = slide_ref
    proxy = _ShapeProxy(sp, slide)
    rId = fill_shape_with_image(proxy, slide, image_path)
    if rId is None:
        return sp
    if border_color:
        ln = spPr.find(qn("a:ln"))
        if ln is not None:
            spPr.remove(ln)
        ln = etree.SubElement(spPr, qn("a:ln"))
        ln.set("w", "12700")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", border_color.lstrip("#"))
    return sp
