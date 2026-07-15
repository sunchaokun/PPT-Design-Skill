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
        normalized_elem = self._normalize_coordinates(grp_elem, orig_bounds)
        grp_xml = etree.tostring(normalized_elem, xml_declaration=False, encoding="unicode")

        aspect = orig_bounds[2] / orig_bounds[3] if orig_bounds[3] > 0 else 1.0

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
            "xml_parts": {"group": grp_xml.encode("utf-8")},
            "meta": {"aspect": round(aspect, 3), "normalized": True},
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
        normalized_elem = self._normalize_coordinates(grp_elem, orig_bounds)
        grp_xml = etree.tostring(normalized_elem, xml_declaration=False, encoding="unicode")

        aspect = orig_bounds[2] / orig_bounds[3] if orig_bounds[3] > 0 else 1.0

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
            "xml_parts": {"group": grp_xml.encode("utf-8")},
            "meta": {"aspect": round(aspect, 3), "normalized": True},
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

    def _normalize_coordinates(self, grp_elem, orig_bounds: tuple[int, int, int, int]) -> Any:
        import copy as _copy
        a_ns = _NS["a"]

        ox, oy, ow, oh = orig_bounds
        if ow <= 0 or oh <= 0:
            return grp_elem

        elem = _copy.deepcopy(grp_elem)

        self._normalize_group_recursive(elem, ox, oy, ow, oh, is_root=True)

        for ln in elem.iter(f"{{{a_ns}}}ln"):
            w = ln.get("w")
            if w:
                try:
                    rel_w = int(w) / oh
                    ln.set("w", _f(rel_w))
                except (ValueError, TypeError):
                    pass

        for rpr in elem.iter(f"{{{a_ns}}}rPr"):
            sz = rpr.get("sz")
            if sz:
                try:
                    rel_sz = int(sz) / oh
                    rpr.set("sz", _f(rel_sz))
                except (ValueError, TypeError):
                    pass

        for end_rpr in elem.iter(f"{{{a_ns}}}endParaRPr"):
            sz = end_rpr.get("sz")
            if sz:
                try:
                    rel_sz = int(sz) / oh
                    end_rpr.set("sz", _f(rel_sz))
                except (ValueError, TypeError):
                    pass

        for bodyPr in elem.iter(f"{{{a_ns}}}bodyPr"):
            for attr in ("lIns", "tIns", "rIns", "bIns"):
                val = bodyPr.get(attr)
                if val:
                    try:
                        rel_val = int(val) / oh
                        bodyPr.set(attr, _f(rel_val))
                    except (ValueError, TypeError):
                        pass

        return elem

    def _normalize_group_recursive(self, grp_elem, parent_chx: int, parent_chy: int, parent_chw: int, parent_chh: int, is_root: bool = False) -> None:
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        grpSpPr = grp_elem.find(f"{{{p_ns}}}grpSpPr")
        if grpSpPr is None:
            return

        xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
        if xfrm is None:
            return

        off = xfrm.find(f"{{{a_ns}}}off")
        ext = xfrm.find(f"{{{a_ns}}}ext")
        chOff = xfrm.find(f"{{{a_ns}}}chOff")
        chExt = xfrm.find(f"{{{a_ns}}}chExt")

        if off is not None:
            try:
                rel_x = (int(off.get("x", "0")) - parent_chx) / parent_chw
                rel_y = (int(off.get("y", "0")) - parent_chy) / parent_chh
                off.set("x", _f(rel_x))
                off.set("y", _f(rel_y))
            except (ValueError, TypeError):
                pass

        if ext is not None:
            try:
                rel_cx = int(ext.get("cx", "0")) / parent_chw
                rel_cy = int(ext.get("cy", "0")) / parent_chh
                ext.set("cx", _f(rel_cx))
                ext.set("cy", _f(rel_cy))
            except (ValueError, TypeError):
                pass

        orig_chx = parent_chx
        orig_chy = parent_chy
        orig_chw = parent_chw
        orig_chh = parent_chh

        if chOff is not None:
            try:
                orig_chx = int(chOff.get("x", "0"))
                orig_chy = int(chOff.get("y", "0"))
                rel_x = (orig_chx - parent_chx) / parent_chw
                rel_y = (orig_chy - parent_chy) / parent_chh
                chOff.set("x", _f(rel_x))
                chOff.set("y", _f(rel_y))
            except (ValueError, TypeError):
                pass

        if chExt is not None:
            try:
                orig_chw = int(chExt.get("cx", "0"))
                orig_chh = int(chExt.get("cy", "0"))
                rel_cx = orig_chw / parent_chw
                rel_cy = orig_chh / parent_chh
                chExt.set("cx", _f(rel_cx))
                chExt.set("cy", _f(rel_cy))
            except (ValueError, TypeError):
                pass

        for child in grp_elem:
            tag = child.tag
            if tag == f"{{{p_ns}}}grpSp":
                self._normalize_group_recursive(child, orig_chx, orig_chy, orig_chw, orig_chh)
            elif tag in (f"{{{p_ns}}}sp", f"{{{p_ns}}}pic", f"{{{p_ns}}}cxnSp", f"{{{p_ns}}}graphicFrame"):
                spPr = child.find(f"{{{p_ns}}}spPr")
                if spPr is None:
                    continue
                child_xfrm = spPr.find(f"{{{a_ns}}}xfrm")
                if child_xfrm is None:
                    continue
                c_off = child_xfrm.find(f"{{{a_ns}}}off")
                c_ext = child_xfrm.find(f"{{{a_ns}}}ext")
                if c_off is not None:
                    try:
                        rel_x = (int(c_off.get("x", "0")) - orig_chx) / orig_chw
                        rel_y = (int(c_off.get("y", "0")) - orig_chy) / orig_chh
                        c_off.set("x", _f(rel_x))
                        c_off.set("y", _f(rel_y))
                    except (ValueError, TypeError):
                        pass
                if c_ext is not None:
                    try:
                        rel_cx = int(c_ext.get("cx", "0")) / orig_chw
                        rel_cy = int(c_ext.get("cy", "0")) / orig_chh
                        c_ext.set("cx", _f(rel_cx))
                        c_ext.set("cy", _f(rel_cy))
                    except (ValueError, TypeError):
                        pass


def _f(v: float) -> str:
    return f"{v:.6f}"
