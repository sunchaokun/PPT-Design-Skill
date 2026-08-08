"""SmartArtExtractor — extract SmartArt text, structure, and XML from .pptx.

Uses lxml low-level XML parsing to extract SmartArt data that python-pptx
high-level API cannot access. Key findings from R&D:

- SmartArt text is in ppt/diagrams/dataN.xml (dgm:pt/a:t nodes)
- Text-bearing pt nodes have type="" (empty), not "node"
- Skip pt types: doc, parTrans, sibTrans, pres
- Layout type from ppt/diagrams/layoutN.xml (dgm:alg/@type)
- Colors from ppt/diagrams/colorsN.xml (schemeClr roles)
- drawing.xml NOT needed (PowerPoint auto-rebuilds)
- colors.xml MUST be stored as-is (131 color mappings, can't generate from scratch)
"""

from __future__ import annotations

import os
import zipfile
from typing import Any

from lxml import etree


_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

_SKIP_PT_TYPES = {"doc", "parTrans", "sibTrans", "pres"}

_ALG_TYPE_MAP = {
    "lin": "process",
    "snake": "process",
    "cycle": "cycle",
    "hier": "hierarchy",
    "pyra": "pyramid",
    "matrix": "matrix",
    "seg": "relationship",
    "pic": "picture",
}


class SmartArtExtractor:

    def extract(self, pptx_path: str, slide_index: int, shape_rId: str) -> dict[str, Any] | None:
        try:
            with zipfile.ZipFile(pptx_path) as z:
                names = z.namelist()
                rels_path = f"ppt/slides/_rels/slide{slide_index + 1}.xml.rels"
                if rels_path not in names:
                    return None

                rels_xml = z.read(rels_path)
                rels_root = etree.fromstring(rels_xml)
                rel_ns = _NS["rel"]

                dgm_refs: dict[str, str] = {}
                for rel in rels_root.findall(f"{{{rel_ns}}}Relationship"):
                    target = rel.get("Target", "")
                    rid = rel.get("Id", "")
                    if "diagram" in target.lower():
                        dgm_refs[rid] = target

                if shape_rId not in dgm_refs:
                    return None

                target = dgm_refs[shape_rId]
                parts_dir = os.path.dirname(target)
                abs_dir = os.path.normpath(f"ppt/slides/{parts_dir}").replace("\\", "/")

                xml_parts: dict[str, bytes] = {}
                for part_name in ("data", "layout", "colors", "quickStyle"):
                    fname = os.path.basename(target)
                    part_fname = self._replace_prefix(fname, part_name)
                    part_path = f"{abs_dir}/{part_fname}"
                    if part_path in names:
                        xml_parts[part_name] = z.read(part_path)

                if "data" not in xml_parts:
                    return None

                nodes = self._parse_data_xml(xml_parts["data"])
                category = "process"
                variant = ""
                if "layout" in xml_parts:
                    layout_info = self._parse_layout_xml(xml_parts["layout"])
                    category = layout_info.get("category", "process")
                    variant = layout_info.get("variant", "")

                color_roles = {}
                if "colors" in xml_parts:
                    color_roles = self._parse_colors_xml(xml_parts["colors"])

                return {
                    "type": "smartart",
                    "category": category,
                    "variant": variant,
                    "nodes": nodes,
                    "color_roles": color_roles,
                    "xml_parts": xml_parts,
                }
        except Exception:
            return None

    def extract_all(self, pptx_path: str) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        try:
            with zipfile.ZipFile(pptx_path) as z:
                names = z.namelist()
                slide_rels = sorted([
                    n for n in names
                    if n.startswith("ppt/slides/_rels/") and n.endswith(".rels")
                ])

                for sr in slide_rels:
                    content = z.read(sr)
                    root = etree.fromstring(content)
                    rel_ns = _NS["rel"]

                    dgm_refs: dict[str, str] = {}
                    for rel in root.findall(f"{{{rel_ns}}}Relationship"):
                        target = rel.get("Target", "")
                        rid = rel.get("Id", "")
                        if "diagram" in target.lower():
                            dgm_refs[rid] = target

                    if not dgm_refs:
                        continue

                    for rid, target in dgm_refs.items():
                        parts_dir = os.path.dirname(target)
                        abs_dir = os.path.normpath(f"ppt/slides/{parts_dir}").replace("\\", "/")
                        fname = os.path.basename(target)

                        if not fname.startswith("data"):
                            continue

                        abs_path = f"{abs_dir}/{fname}"
                        if abs_path not in names:
                            continue

                        xml_parts: dict[str, bytes] = {}
                        for part_name in ("data", "layout", "colors", "quickStyle"):
                            part_fname = self._replace_prefix(fname, part_name)
                            part_path = f"{abs_dir}/{part_fname}"
                            if part_path in names:
                                xml_parts[part_name] = z.read(part_path)

                        if "data" not in xml_parts:
                            continue

                        nodes = self._parse_data_xml(xml_parts["data"])
                        category = "process"
                        variant = ""
                        if "layout" in xml_parts:
                            layout_info = self._parse_layout_xml(xml_parts["layout"])
                            category = layout_info.get("category", "process")
                            variant = layout_info.get("variant", "")

                        color_roles = {}
                        if "colors" in xml_parts:
                            color_roles = self._parse_colors_xml(xml_parts["colors"])

                        results.append({
                            "type": "smartart",
                            "category": category,
                            "variant": variant,
                            "nodes": nodes,
                            "color_roles": color_roles,
                            "xml_parts": xml_parts,
                        })
        except Exception:
            pass

        return results

    def _parse_data_xml(self, data_xml: bytes) -> list[dict]:
        root = etree.fromstring(data_xml)
        a_ns = _NS["a"]
        dgm_ns = _NS["dgm"]
        nodes = []

        for pt in root.iter(f"{{{dgm_ns}}}pt"):
            pt_type = pt.get("type", "")
            if pt_type in _SKIP_PT_TYPES:
                continue
            t_elem = pt.find(f".//{{{a_ns}}}t")
            if t_elem is not None and t_elem.text:
                level = 0
                prlo = pt.find(f"{{{dgm_ns}}}prLo")
                if prlo is not None:
                    try:
                        level = int(prlo.get("val", "0"))
                    except ValueError:
                        pass
                nodes.append({"id": len(nodes), "text": t_elem.text, "level": level})

        return nodes

    def _parse_layout_xml(self, layout_xml: bytes) -> dict:
        root = etree.fromstring(layout_xml)
        dgm_ns = _NS["dgm"]

        category = "process"
        variant = ""

        alg_elems = root.findall(f".//{{{dgm_ns}}}alg")
        for alg in alg_elems:
            alg_type = alg.get("type", "")
            category = _ALG_TYPE_MAP.get(alg_type, category)
            break

        cat_elems = root.findall(f".//{{{dgm_ns}}}cat")
        for cat in cat_elems:
            cat_type = cat.get("type", "")
            if cat_type:
                variant = cat_type.split("/")[-1] if "/" in cat_type else cat_type
                break

        return {"category": category, "variant": variant}

    def _parse_colors_xml(self, colors_xml: bytes) -> dict[str, int]:
        root = etree.fromstring(colors_xml)
        a_ns = _NS["a"]
        roles: dict[str, int] = {}

        for scheme in root.iter(f"{{{a_ns}}}schemeClr"):
            val = scheme.get("val", "")
            if val:
                roles[val] = roles.get(val, 0) + 1

        return roles

    def _replace_prefix(self, fname: str, new_prefix: str) -> str:
        for prefix in ("data", "layout", "colors", "quickStyle", "drawing"):
            if fname.startswith(prefix):
                return new_prefix + fname[len(prefix):]
        return fname
