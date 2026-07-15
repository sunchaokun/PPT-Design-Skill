"""ComponentRenderer — match component → fill data → apply brand → inject to slide.

Bridge between ComponentLibrary and PrecisionRenderer. Handles:
1. SmartArt rendering: match template → fill text → apply brand colors → inject XML
2. GroupShape rendering: match template or rebuild with shapes
3. Fallback to DiagramEngine when no component match

Key constraints from R&D:
- Only modify data XML for text (PowerPoint auto-syncs to rendering)
- Apply brand colors by replacing schemeClr → srgbClr in colors XML
- Clone original colors XML (131 mappings, can't generate from scratch)
- No drawing.xml needed (PowerPoint auto-rebuilds)
- 4 diagram MIME types must be exact
- slide rels need 5 rIds (rId1=slideLayout + rId2-5=dm/lo/qs/cs)
"""

from __future__ import annotations

import copy

from lxml import etree

from ppt_pro_max.enterprise.brand_spec import BrandSpec


_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

_SKIP_PT_TYPES = {"doc", "parTrans", "sibTrans", "pres"}

_CONTENT_TYPE_MAP = {
    "data": "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
    "layout": "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
    "quickStyle": "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
    "colors": "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
}

_REL_TYPE_MAP = {
    "data": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
    "layout": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
    "quickStyle": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
    "colors": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
}

_ALG_CATEGORY_MAP = {
    "process": "process",
    "cycle": "cycle",
    "hierarchy": "tree",
    "pyramid": "pyramid",
    "matrix": "swot",
}


class ComponentRenderer:

    def render(self, slide, complex_element: dict, brand_spec: BrandSpec | None = None,
               component_lib=None) -> bool:
        elem_type = complex_element.get("type", "")

        if elem_type == "smartart":
            return self.render_smartart(slide, complex_element, brand_spec, component_lib)
        elif elem_type == "group":
            return self.render_group(slide, complex_element, brand_spec, component_lib)

        return False

    def render_smartart(self, slide, element: dict, brand_spec: BrandSpec | None = None,
                        component_lib=None) -> bool:
        if component_lib is not None:
            component = component_lib.match(element)
            if component is not None:
                xml_parts = component_lib.load_xml(component["id"])
                if xml_parts:
                    filled = self._fill_data(xml_parts, element.get("nodes", []))
                    styled = self._apply_brand_colors(filled, brand_spec)
                    self._inject_to_slide(slide, styled, element.get("bounds", (1, 1, 10, 5)))
                    return True

        return self._fallback_diagram(slide, element, brand_spec)

    def _fill_data(self, xml_parts: dict, nodes: list[dict]) -> dict:
        if "data" not in xml_parts:
            return xml_parts

        data_xml = xml_parts["data"]
        root = etree.fromstring(data_xml)
        a_ns = _NS["a"]
        dgm_ns = _NS["dgm"]

        node_idx = 0
        for pt in root.iter(f"{{{dgm_ns}}}pt"):
            pt_type = pt.get("type", "")
            if pt_type in _SKIP_PT_TYPES:
                continue
            t_elem = pt.find(f".//{{{a_ns}}}t")
            if t_elem is not None and t_elem.text is not None:
                if node_idx < len(nodes):
                    new_text = nodes[node_idx].get("text", nodes[node_idx]) if isinstance(nodes[node_idx], dict) else str(nodes[node_idx])
                    t_elem.text = new_text
                node_idx += 1

        result = dict(xml_parts)
        result["data"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        return result

    def _apply_brand_colors(self, xml_parts: dict, brand_spec: BrandSpec | None) -> dict:
        if brand_spec is None or brand_spec.colors is None or "colors" not in xml_parts:
            return xml_parts

        brand_colors = {}
        for role, hex_val in brand_spec.colors.items():
            if not hex_val.startswith("#"):
                hex_val = f"#{hex_val}"
            brand_colors[role] = hex_val.lstrip("#")

        colors_xml = xml_parts["colors"]
        root = etree.fromstring(colors_xml)
        a_ns = _NS["a"]

        for scheme in list(root.iter(f"{{{a_ns}}}schemeClr")):
            val = scheme.get("val", "")
            if val in brand_colors:
                parent = scheme.getparent()
                if parent is not None:
                    srgb = etree.SubElement(parent, f"{{{a_ns}}}srgbClr")
                    srgb.set("val", brand_colors[val])
                    for child in scheme:
                        srgb.append(copy.deepcopy(child))
                    parent.remove(scheme)

        result = dict(xml_parts)
        result["colors"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        return result

    def _inject_to_slide(self, slide, xml_parts: dict, bounds: tuple) -> None:
        self._inject_via_python_pptx(slide, xml_parts, bounds)

    def _inject_via_zip(self, pptx_path: str, slide, xml_parts: dict, bounds: tuple) -> None:
        pass

    def _inject_via_python_pptx(self, slide, xml_parts: dict, bounds: tuple) -> None:
        try:
            ns_p = _NS["p"]
            ns_a = _NS["a"]
            ns_dgm = _NS["dgm"]
            ns_r = _NS["r"]

            sp_tree = slide._element.find(f".//{{{ns_p}}}spTree")
            if sp_tree is None:
                return

            x_emu = int(bounds[0] * 914400) if bounds else 914400
            y_emu = int(bounds[1] * 914400) if bounds else 914400
            cx_emu = int(bounds[2] * 914400) if bounds else 9144000
            cy_emu = int(bounds[3] * 914400) if bounds else 5486400

            next_id = 2
            for cNv in sp_tree.iter(f"{{{ns_p}}}cNvPr"):
                try:
                    next_id = max(next_id, int(cNv.get("id", "1")) + 1)
                except (ValueError, TypeError):
                    pass

            graphic_frame_xml = f'''<p:graphicFrame xmlns:p="{ns_p}" xmlns:a="{ns_a}" xmlns:r="{ns_r}" xmlns:dgm="{ns_dgm}">
  <p:nvGraphicFramePr>
    <p:cNvPr id="{next_id}" name="SmartArt"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="{x_emu}" y="{y_emu}"/>
    <a:ext cx="{cx_emu}" cy="{cy_emu}"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
      <dgm:relIds r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>'''

            gf_elem = etree.fromstring(graphic_frame_xml)
            sp_tree.append(gf_elem)

            slide_part = slide.part

            diagram_idx = 1
            try:
                for part in slide_part.package.iter_parts():
                    pname = str(getattr(part, "partname", ""))
                    if "/ppt/diagrams/" in pname:
                        import re
                        m = re.search(r"(\d+)\.xml$", pname)
                        if m:
                            diagram_idx = max(diagram_idx, int(m.group(1)) + 1)
            except Exception:
                pass

            for part_name, part_xml in xml_parts.items():
                if part_name not in _REL_TYPE_MAP:
                    continue
                content_type = _CONTENT_TYPE_MAP.get(part_name, "application/xml")
                blob = part_xml if isinstance(part_xml, bytes) else part_xml.encode("utf-8")

                from pptx.opc.package import Part
                from pptx.opc.packuri import PackURI
                part_obj = Part(
                    partname=PackURI(f"/ppt/diagrams/{part_name}{diagram_idx}.xml"),
                    content_type=content_type,
                    package=slide_part.package,
                    blob=blob,
                )

                slide_part.relate_to(part_obj, _REL_TYPE_MAP[part_name])

        except Exception as exc:
            import logging
            logging.getLogger(__name__).warning("SmartArt injection failed: %s", exc)

    def _fallback_diagram(self, slide, element: dict, brand_spec: BrandSpec | None) -> bool:
        category = element.get("category", "process")
        diagram_type = _ALG_CATEGORY_MAP.get(category, "process")
        nodes = element.get("nodes", [])

        if not nodes:
            return False

        try:
            from ppt_pro_max.renderer.diagram_engine import DiagramEngine
            from ppt_pro_max.renderer.diagram.diagram_style import DiagramStyle
            from ppt_pro_max.renderer.diagram.layout_engine import Region

            style = DiagramStyle.from_brand_spec(brand_spec) if brand_spec else DiagramStyle()
            region = Region(left=0.9, top=1.5, width=7.0, height=5.0)

            diagram_data = {
                "items": [{"label": n.get("text", str(n)) if isinstance(n, dict) else str(n)} for n in nodes],
            }

            engine = DiagramEngine()
            engine.render(slide, diagram_type, diagram_data, style, region)
            return True
        except Exception:
            return False

    def render_group(self, slide, element: dict, brand_spec: BrandSpec | None = None,
                     component_lib=None) -> bool:
        if component_lib is not None:
            component = component_lib.match(element)
            if component is not None:
                xml_parts = component_lib.load_xml(component["id"])
                if xml_parts and "group" in xml_parts:
                    filled = self._fill_group_data(xml_parts, element.get("texts", []))
                    styled = self._apply_brand_colors(filled, brand_spec)
                    self._inject_group_to_slide(slide, styled, element.get("bounds", (1, 1, 10, 5)))
                    return True

        return self._fallback_group(slide, element, brand_spec)

    def _fill_group_data(self, xml_parts: dict, new_texts: list) -> dict:
        if "group" not in xml_parts or not new_texts:
            return xml_parts

        group_xml = xml_parts["group"]
        root = etree.fromstring(group_xml)
        a_ns = _NS["a"]

        text_idx = 0
        for t_elem in root.iter(f"{{{a_ns}}}t"):
            if t_elem.text and t_elem.text.strip():
                if text_idx < len(new_texts):
                    new_text = new_texts[text_idx]
                    if isinstance(new_text, dict):
                        new_text = new_text.get("text", "")
                    t_elem.text = str(new_text)
                    text_idx += 1

        result = dict(xml_parts)
        result["group"] = etree.tostring(root, xml_declaration=False, encoding="UTF-8")
        return result

    def _inject_group_to_slide(self, slide, xml_parts: dict, bounds: tuple) -> None:
        if "group" not in xml_parts:
            return

        try:
            ns_p = _NS["p"]
            ns_a = _NS["a"]
            ns_r = _NS["r"]

            sp_tree = slide._element.find(f".//{{{ns_p}}}spTree")
            if sp_tree is None:
                return

            group_xml = xml_parts["group"]
            if isinstance(group_xml, str):
                group_xml = group_xml.encode("utf-8")
            grp_elem = etree.fromstring(group_xml)

            next_id = 2
            for sp in sp_tree:
                for cNv in sp.iter(f"{{{ns_p}}}cNvPr"):
                    try:
                        next_id = max(next_id, int(cNv.get("id", "1")) + 1)
                    except (ValueError, TypeError):
                        pass

            for cNv in grp_elem.iter(f"{{{ns_p}}}cNvPr"):
                cNv.set("id", str(next_id))
                next_id += 1

            for cNv in grp_elem.iter(f"{{{ns_a}}}cNvPr"):
                if cNv.getparent() is not None and cNv.getparent().tag != f"{{{ns_p}}}nvGrpSpPr":
                    cNv.set("id", str(next_id))
                    next_id += 1

            target_left = bounds[0] if bounds and len(bounds) > 0 else 0.9
            target_top = bounds[1] if bounds and len(bounds) > 1 else 1.5
            target_w = bounds[2] if bounds and len(bounds) > 2 else 11.5
            target_h = bounds[3] if bounds and len(bounds) > 3 else 5.2

            self._denormalize_coordinates(grp_elem, target_left, target_top, target_w, target_h)

            image_keys = [k for k in xml_parts if k.startswith("img_")]
            if image_keys:
                slide_part = slide.part
                rid_remap: dict[str, str] = {}
                for img_key in image_keys:
                    img_blob = xml_parts[img_key]
                    original_rId = img_key[4:]

                    from pptx.opc.package import Part
                    try:
                        from pptx.opc.spec import content_type as _guess_ct  # noqa: F401
                    except ImportError:
                        pass

                    content_type = "image/png"
                    if img_blob[:3] == b"\xff\xd8\xff":
                        content_type = "image/jpeg"
                    elif img_blob[:8] == b"\x89PNG\r\n\x1a\n":
                        content_type = "image/png"
                    elif img_blob[:4] == b"RIFF":
                        content_type = "image/webp"
                    elif img_blob[:4] == b"GIF8":
                        content_type = "image/gif"

                    img_part = Part(
                        partname=f"/ppt/media/library_img_{original_rId}.png",
                        content_type=content_type,
                        blob=img_blob,
                        package=slide_part.package,
                    )
                    new_rId = slide_part.relate_to(img_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                    rid_remap[original_rId] = new_rId

                for blip in grp_elem.iter(f"{{{ns_a}}}blip"):
                    old_embed = blip.get(f"{{{ns_r}}}embed")
                    if old_embed and old_embed in rid_remap:
                        blip.set(f"{{{ns_r}}}embed", rid_remap[old_embed])

            sp_tree.append(grp_elem)

        except Exception as exc:
            import logging
            logging.getLogger(__name__).warning("GroupShape injection failed: %s", exc)

    def _denormalize_coordinates(self, grp_elem, left: float, top: float, width: float, height: float) -> None:
        a_ns = _NS["a"]

        left_emu = int(left * 914400)
        top_emu = int(top * 914400)
        w_emu = int(width * 914400)
        h_emu = int(height * 914400)

        for xfrm in grp_elem.iter(f"{{{a_ns}}}xfrm"):
            off = xfrm.find(f"{{{a_ns}}}off")
            ext = xfrm.find(f"{{{a_ns}}}ext")
            chOff = xfrm.find(f"{{{a_ns}}}chOff")
            chExt = xfrm.find(f"{{{a_ns}}}chExt")

            if off is not None:
                try:
                    rx = float(off.get("x", "0"))
                    ry = float(off.get("y", "0"))
                    off.set("x", str(int(rx * w_emu) + left_emu))
                    off.set("y", str(int(ry * h_emu) + top_emu))
                except (ValueError, TypeError):
                    pass

            if ext is not None:
                try:
                    rcx = float(ext.get("cx", "0"))
                    rcy = float(ext.get("cy", "0"))
                    ext.set("cx", str(int(rcx * w_emu)))
                    ext.set("cy", str(int(rcy * h_emu)))
                except (ValueError, TypeError):
                    pass

            if chOff is not None:
                try:
                    rx = float(chOff.get("x", "0"))
                    ry = float(chOff.get("y", "0"))
                    chOff.set("x", str(int(rx * w_emu) + left_emu))
                    chOff.set("y", str(int(ry * h_emu) + top_emu))
                except (ValueError, TypeError):
                    pass

            if chExt is not None:
                try:
                    rcx = float(chExt.get("cx", "0"))
                    rcy = float(chExt.get("cy", "0"))
                    chExt.set("cx", str(int(rcx * w_emu)))
                    chExt.set("cy", str(int(rcy * h_emu)))
                except (ValueError, TypeError):
                    pass

        for ln in grp_elem.iter(f"{{{a_ns}}}ln"):
            w = ln.get("w")
            if w:
                try:
                    ln.set("w", str(int(float(w) * h_emu)))
                except (ValueError, TypeError):
                    pass

        for rpr in grp_elem.iter(f"{{{a_ns}}}rPr"):
            sz = rpr.get("sz")
            if sz:
                try:
                    sz_hundredths = int(float(sz) * h_emu)
                    clamped = max(800, min(7200, sz_hundredths))  # 8pt-72pt
                    rpr.set("sz", str(clamped))
                except (ValueError, TypeError):
                    pass

        for end_rpr in grp_elem.iter(f"{{{a_ns}}}endParaRPr"):
            sz = end_rpr.get("sz")
            if sz:
                try:
                    sz_hundredths = int(float(sz) * h_emu)
                    clamped = max(800, min(7200, sz_hundredths))  # 8pt-72pt
                    end_rpr.set("sz", str(clamped))
                except (ValueError, TypeError):
                    pass

        for bodyPr in grp_elem.iter(f"{{{a_ns}}}bodyPr"):
            for attr in ("lIns", "tIns", "rIns", "bIns"):
                val = bodyPr.get(attr)
                if val:
                    try:
                        bodyPr.set(attr, str(int(float(val) * h_emu)))
                    except (ValueError, TypeError):
                        pass

    def _fallback_group(self, slide, element: dict, brand_spec: BrandSpec | None) -> bool:
        texts = element.get("texts", [])
        if not texts:
            return False

        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer

        renderer = PrecisionRenderer(brand_spec=brand_spec)
        bullet_lines = [f"•  {t}" for t in texts[:8]]
        renderer.add_multiline(slide, bullet_lines, 0.9, 1.6, 7, 4.5,
                               size=12, color_role="foreground", spacing=6)
        return True
