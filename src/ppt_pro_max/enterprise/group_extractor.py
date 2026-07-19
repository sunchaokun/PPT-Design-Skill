"""GroupExtractor — recursive extraction of GroupShape content + XML for library.

Extracts the complete <p:grpSp> XML, embedded image blobs, text nodes,
and structural metadata so the GroupShape can be stored in ComponentLibrary
and re-injected into new presentations.

Key design:
- `extract()`: returns full XML + image blobs + metadata for library import
- `extract_all()`: scans all slides in a PPTX for GroupShapes
- XML is stored as-is (preserving all visual properties)
- Image blobs stored as {rId: bytes} for later re-injection
- Text nodes identified by position for data filling
"""

from __future__ import annotations

import os
import zipfile
from typing import Any

from lxml import etree

_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

_CHILD_SHAPE_TAGS = {
    f"{{{_NS['p']}}}sp",
    f"{{{_NS['p']}}}grpSp",
    f"{{{_NS['p']}}}pic",
    f"{{{_NS['p']}}}cxnSp",
    f"{{{_NS['p']}}}graphicFrame",
}


class GroupExtractor:

    def extract(self, group_shape) -> dict[str, Any]:
        grp_elem = group_shape._element

        a_ns = _NS["a"]

        texts: list[dict[str, Any]] = []
        for t_elem in grp_elem.iter(f"{{{a_ns}}}t"):
            if t_elem.text and t_elem.text.strip():
                texts.append({"text": t_elem.text.strip(), "role": "title" if self._is_title_context(t_elem) else "body"})

        child_count = self._count_direct_children(grp_elem)
        depth = self._measure_depth(grp_elem)

        bounds = self._get_bounds(grp_elem)

        category = self._infer_category(grp_elem, texts)

        slide_part = group_shape.part
        image_blobs: dict[str, bytes] = {}
        rId_map: dict[str, str] = {}
        for blip in grp_elem.iter(f"{{{a_ns}}}blip"):
            embed = blip.get(f"{{{_NS['r']}}}embed")
            if embed:
                try:
                    rel = slide_part.rels[embed]
                    image_part = rel.target_part
                    image_blobs[embed] = image_part.blob
                    rId_map[embed] = f"img_{len(rId_map)}"
                except Exception:
                    pass

        orig_bounds = self._get_bounds_emu(grp_elem)
        theme_colors = self._resolve_theme_colors(group_shape)
        resolved_elem = self._resolve_scheme_colors(grp_elem, theme_colors)
        grp_xml = etree.tostring(resolved_elem, xml_declaration=False, encoding="unicode")

        aspect = orig_bounds[2] / orig_bounds[3] if orig_bounds[3] > 0 else 1.0

        import json as _json

        return {
            "type": "group",
            "category": category,
            "variant": "",
            "node_count": len(texts),
            "level_count": depth,
            "texts": texts,
            "bounds": bounds,
            "child_count": child_count,
            "image_blobs": image_blobs,
            "rid_map": rId_map,
            "xml_parts": {
                "group": grp_xml.encode("utf-8"),
                "_meta": _json.dumps({"orig_bounds_emu": list(orig_bounds)}, ensure_ascii=False).encode("utf-8"),
            },
            "meta": {
                "aspect": round(aspect, 3),
                "orig_bounds_emu": list(orig_bounds),
            },
        }

    def extract_all(self, pptx_path: str) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        try:
            with zipfile.ZipFile(pptx_path) as z:
                names = z.namelist()
                slide_files = sorted([
                    n for n in names
                    if n.startswith("ppt/slides/slide") and n.endswith(".xml") and "rels" not in n
                ])

                for sf in slide_files:
                    slide_xml = z.read(sf)
                    slide_root = etree.fromstring(slide_xml)
                    p_ns = _NS["p"]

                    rels_path = sf.replace("slides/", "slides/_rels/") + ".rels"
                    slide_rels = {}
                    if rels_path in names:
                        rels_xml = z.read(rels_path)
                        rels_root = etree.fromstring(rels_xml)
                        for rel in rels_root.findall(f"{{{_NS['rel']}}}Relationship"):
                            slide_rels[rel.get("Id", "")] = rel.get("Target", "")

                    sp_tree = slide_root.find(f".//{{{p_ns}}}spTree")
                    if sp_tree is None:
                        continue

                    for grp_elem in sp_tree.iter(f"{{{p_ns}}}grpSp"):
                        result = self._extract_from_xml_element(grp_elem, z, slide_rels, sf)
                        if result:
                            results.append(result)

        except Exception:
            pass

        return results

    def _extract_from_xml_element(
        self, grp_elem, z: zipfile.ZipFile, slide_rels: dict, slide_path: str
    ) -> dict[str, Any] | None:
        a_ns = _NS["a"]

        grp_xml = etree.tostring(grp_elem, xml_declaration=False, encoding="unicode")

        texts: list[dict[str, Any]] = []
        for t_elem in grp_elem.iter(f"{{{a_ns}}}t"):
            if t_elem.text and t_elem.text.strip():
                texts.append({"text": t_elem.text.strip(), "role": "title" if self._is_title_context(t_elem) else "body"})

        if not texts and not self._has_visual_shapes(grp_elem):
            return None

        child_count = self._count_direct_children(grp_elem)
        depth = self._measure_depth(grp_elem)
        bounds = self._get_bounds(grp_elem)
        category = self._infer_category(grp_elem, texts)

        image_blobs: dict[str, bytes] = {}
        rId_map: dict[str, str] = {}
        for blip in grp_elem.iter(f"{{{a_ns}}}blip"):
            embed = blip.get(f"{{{_NS['r']}}}embed")
            if embed and embed in slide_rels:
                target = slide_rels[embed]
                try:
                    abs_target = os.path.normpath(os.path.join(os.path.dirname(slide_path), target)).replace("\\", "/")
                    if abs_target in z.namelist():
                        image_blobs[embed] = z.read(abs_target)
                        rId_map[embed] = f"img_{len(rId_map)}"
                except Exception:
                    pass

        orig_bounds = self._get_bounds_emu_from_xml(grp_elem)
        theme_colors = self._resolve_theme_colors_from_zip(grp_elem, z, slide_rels, slide_path)
        resolved_elem = self._resolve_scheme_colors(grp_elem, theme_colors)
        grp_xml = etree.tostring(resolved_elem, xml_declaration=False, encoding="unicode")

        aspect = orig_bounds[2] / orig_bounds[3] if orig_bounds[3] > 0 else 1.0

        import json as _json

        return {
            "type": "group",
            "category": category,
            "variant": "",
            "node_count": len(texts),
            "level_count": depth,
            "texts": texts,
            "bounds": bounds,
            "child_count": child_count,
            "image_blobs": image_blobs,
            "rid_map": rId_map,
            "xml_parts": {
                "group": grp_xml.encode("utf-8"),
                "_meta": _json.dumps({"orig_bounds_emu": list(orig_bounds)}, ensure_ascii=False).encode("utf-8"),
            },
            "meta": {
                "aspect": round(aspect, 3),
                "orig_bounds_emu": list(orig_bounds),
            },
        }

    def _is_title_context(self, t_elem) -> bool:
        parent = t_elem.getparent()
        while parent is not None:
            rpr = parent.find(f"{{{_NS['a']}}}rPr")
            if rpr is not None:
                if rpr.get("b") == "1":
                    return True
                sz = rpr.get("sz")
                if sz and int(sz) >= 2400:
                    return True
            parent = parent.getparent()
        return False

    def _count_direct_children(self, grp_elem) -> int:
        count = 0
        for child in grp_elem:
            if child.tag in _CHILD_SHAPE_TAGS:
                count += 1
        return count

    def _measure_depth(self, grp_elem) -> int:
        p_ns = _NS["p"]
        max_depth = 0
        for sub_grp in grp_elem.iter(f"{{{p_ns}}}grpSp"):
            depth = 0
            parent = sub_grp.getparent()
            while parent is not None:
                if parent.tag == f"{{{p_ns}}}grpSp":
                    depth += 1
                parent = parent.getparent()
            max_depth = max(max_depth, depth)
        return max_depth

    def _get_bounds(self, grp_elem) -> tuple[float, float, float, float]:
        a_ns = _NS["a"]
        grpSpPr = grp_elem.find(f"{{{_NS['p']}}}grpSpPr")
        if grpSpPr is not None:
            xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{{{a_ns}}}off")
                ext = xfrm.find(f"{{{a_ns}}}ext")
                if off is not None and ext is not None:
                    try:
                        x = int(off.get("x", "0")) / 914400
                        y = int(off.get("y", "0")) / 914400
                        w = int(ext.get("cx", "0")) / 914400
                        h = int(ext.get("cy", "0")) / 914400
                        return (round(x, 2), round(y, 2), round(w, 2), round(h, 2))
                    except (ValueError, TypeError):
                        pass
        return (0.0, 0.0, 0.0, 0.0)

    def _has_visual_shapes(self, grp_elem) -> bool:
        p_ns = _NS["p"]
        a_ns = _NS["a"]
        for sp in grp_elem.iter(f"{{{p_ns}}}sp"):
            spPr = sp.find(f"{{{p_ns}}}spPr")
            if spPr is not None:
                fill = spPr.find(f"{{{a_ns}}}solidFill")
                prstGeom = spPr.find(f"{{{a_ns}}}prstGeom")
                if fill is not None or prstGeom is not None:
                    return True
        for cxn in grp_elem.iter(f"{{{p_ns}}}cxnSp"):
            return True
        return False

    def _infer_category(self, grp_elem, texts: list[dict]) -> str:
        p_ns = _NS["p"]
        a_ns = _NS["a"]

        direct_shapes = []
        for child in grp_elem:
            if child.tag in _CHILD_SHAPE_TAGS:
                spPr = child.find(f"{{{p_ns}}}spPr")
                if spPr is not None:
                    xfrm = spPr.find(f"{{{a_ns}}}xfrm")
                    if xfrm is not None:
                        off = xfrm.find(f"{{{a_ns}}}off")
                        ext = xfrm.find(f"{{{a_ns}}}ext")
                        if off is not None and ext is not None:
                            try:
                                x = int(off.get("x", "0")) / 914400
                                y = int(off.get("y", "0")) / 914400
                                w = int(ext.get("cx", "0")) / 914400
                                h = int(ext.get("cy", "0")) / 914400
                                direct_shapes.append({"bounds": (round(x, 2), round(y, 2), round(w, 2), round(h, 2))})
                            except (ValueError, TypeError):
                                pass

        if len(direct_shapes) >= 3:
            xs = [s["bounds"][0] for s in direct_shapes]
            ys = [s["bounds"][1] for s in direct_shapes]
            x_range = max(xs) - min(xs) if len(xs) >= 2 else 0
            y_range = max(ys) - min(ys) if len(ys) >= 2 else 0
            if x_range > y_range * 2:
                return "process"
            if y_range > x_range * 2:
                return "hierarchy"

        text_content = " ".join(t.get("text", "") for t in texts).lower()
        category_keywords = {
            "swot": "swot", "strength": "swot", "weakness": "swot",
            "opportunity": "swot", "threat": "swot",
            "pie": "chart", "bar": "chart", "chart": "chart", "graph": "chart",
            "step": "process", "phase": "process", "stage": "process",
            "flow": "process",
            "org": "hierarchy", "team": "hierarchy", "report": "hierarchy",
            "vs": "comparison", "versus": "comparison", "compare": "comparison",
            "timeline": "timeline", "year": "timeline", "month": "timeline",
            "feature": "features", "benefit": "features",
            "kpi": "dashboard", "metric": "dashboard", "score": "dashboard",
        }
        for kw, cat in category_keywords.items():
            if kw in text_content:
                return cat

        return "infographic"

    def _get_bounds_emu(self, grp_elem) -> tuple[int, int, int, int]:
        a_ns = _NS["a"]
        p_ns = _NS["p"]
        grpSpPr = grp_elem.find(f"{{{p_ns}}}grpSpPr")
        if grpSpPr is not None:
            xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{{{a_ns}}}off")
                ext = xfrm.find(f"{{{a_ns}}}ext")
                if off is not None and ext is not None:
                    try:
                        return (
                            int(off.get("x", "0")),
                            int(off.get("y", "0")),
                            int(ext.get("cx", "0")),
                            int(ext.get("cy", "0")),
                        )
                    except (ValueError, TypeError):
                        pass
        return (0, 0, 9144000, 5486400)

    def _get_bounds_emu_from_xml(self, grp_elem) -> tuple[int, int, int, int]:
        return self._get_bounds_emu(grp_elem)

    def _resolve_theme_colors(self, group_shape) -> dict[str, str]:
        try:
            slide_part = group_shape.part
            theme_part = self._find_theme_part(slide_part)
            if theme_part is None:
                return {}
            return self._parse_theme_xml(theme_part.blob)
        except Exception:
            return {}

    def _find_theme_part(self, slide_part):
        for rel in slide_part.rels.values():
            if "theme" in str(getattr(rel, "reltype", "")):
                return rel.target_part
        for rel in slide_part.rels.values():
            if "slideLayout" in str(getattr(rel, "reltype", "")):
                layout_part = rel.target_part
                for lr in layout_part.rels.values():
                    if "theme" in str(getattr(lr, "reltype", "")):
                        return lr.target_part
                for lr in layout_part.rels.values():
                    if "slideMaster" in str(getattr(lr, "reltype", "")):
                        master_part = lr.target_part
                        for mr in master_part.rels.values():
                            if "theme" in str(getattr(mr, "reltype", "")):
                                return mr.target_part
        return None

    def _resolve_theme_colors_from_zip(self, grp_elem, z, slide_rels, slide_path) -> dict[str, str]:
        try:
            rel_ns = _NS["rel"]
            for rId, target in slide_rels.items():
                if "slideLayout" in target:
                    layout_path = os.path.normpath(os.path.join(os.path.dirname(slide_path), target)).replace("\\", "/")
                    if layout_path not in z.namelist():
                        continue
                    layout_dir = os.path.dirname(layout_path)
                    layout_basename = os.path.basename(layout_path)
                    layout_rels_path = os.path.join(layout_dir, "_rels", layout_basename + ".rels").replace("\\", "/")
                    if layout_rels_path not in z.namelist():
                        continue
                    lr_xml = z.read(layout_rels_path)
                    lr_root = etree.fromstring(lr_xml)
                    for lr in lr_root.findall(f"{{{rel_ns}}}Relationship"):
                        lr_target = lr.get("Target", "")
                        if "theme" in lr_target.lower():
                            abs_theme = os.path.normpath(os.path.join(os.path.dirname(layout_path), lr_target)).replace("\\", "/")
                            if abs_theme in z.namelist():
                                return self._parse_theme_xml(z.read(abs_theme))
                        if "slideMaster" in lr_target:
                            master_path = os.path.normpath(os.path.join(os.path.dirname(layout_path), lr_target)).replace("\\", "/")
                            if master_path not in z.namelist():
                                continue
                            master_dir = os.path.dirname(master_path)
                            master_basename = os.path.basename(master_path)
                            master_rels_path = os.path.join(master_dir, "_rels", master_basename + ".rels").replace("\\", "/")
                            if master_rels_path not in z.namelist():
                                continue
                            mr_xml = z.read(master_rels_path)
                            mr_root = etree.fromstring(mr_xml)
                            for mr in mr_root.findall(f"{{{rel_ns}}}Relationship"):
                                mr_target = mr.get("Target", "")
                                if "theme" in mr_target.lower():
                                    abs_theme = os.path.normpath(os.path.join(os.path.dirname(master_path), mr_target)).replace("\\", "/")
                                    if abs_theme in z.namelist():
                                        return self._parse_theme_xml(z.read(abs_theme))
                    break
            return {}
        except Exception:
            return {}

    def _parse_theme_xml(self, theme_blob: bytes) -> dict[str, str]:
        a_ns = _NS["a"]
        try:
            root = etree.fromstring(theme_blob)
            color_map = {}
            theme_elements = root.find(f".//{{{a_ns}}}themeElements")
            if theme_elements is None:
                return {}
            clr_scheme = theme_elements.find(f"{{{a_ns}}}clrScheme")
            if clr_scheme is None:
                return {}
            for child in clr_scheme:
                tag = child.tag.split("}")[-1]
                srgb = child.find(f"{{{a_ns}}}srgbClr")
                sys_clr = child.find(f"{{{a_ns}}}sysClr")
                if srgb is not None:
                    color_map[tag] = f"#{srgb.get('val', '000000')}"
                elif sys_clr is not None:
                    color_map[tag] = f"#{sys_clr.get('lastClr', '000000')}"
            if "lt1" in color_map and "bg1" not in color_map:
                color_map["bg1"] = color_map["lt1"]
            if "dk1" in color_map and "tx1" not in color_map:
                color_map["tx1"] = color_map["dk1"]
            if "lt2" in color_map and "bg2" not in color_map:
                color_map["bg2"] = color_map["lt2"]
            if "dk2" in color_map and "tx2" not in color_map:
                color_map["tx2"] = color_map["dk2"]
            return color_map
        except Exception:
            return {}

    def _resolve_scheme_colors(self, grp_elem, theme_colors: dict[str, str]):
        if not theme_colors:
            return grp_elem
        import copy as _copy
        a_ns = _NS["a"]
        elem = _copy.deepcopy(grp_elem)
        style_elems = set()
        for style in elem.iter(f"{{{a_ns}}}style"):
            style_elems.add(id(style))
        for scheme in list(elem.iter(f"{{{a_ns}}}schemeClr")):
            parent = scheme.getparent()
            if parent is not None and id(parent) in style_elems:
                continue
            in_style = False
            p = scheme.getparent()
            while p is not None:
                if p.tag == f"{{{a_ns}}}style":
                    in_style = True
                    break
                p = p.getparent()
            if in_style:
                continue
            val = scheme.get("val", "")
            if val in theme_colors:
                if parent is not None:
                    idx = list(parent).index(scheme)
                    srgb = etree.SubElement(parent, f"{{{a_ns}}}srgbClr")
                    srgb.set("val", theme_colors[val].lstrip("#"))
                    for child in scheme:
                        srgb.append(_copy.deepcopy(child))
                    parent.remove(scheme)
                    parent.insert(idx, srgb)
        return elem


