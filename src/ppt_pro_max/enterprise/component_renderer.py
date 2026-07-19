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


_LIPSUM_WORDS = frozenset(
    "aenean commodo ligula eget dolor massa cum sociis natoque penatibus "
    "et magnis dis parturient montes nascetur ridiculus mus donec quam "
    "felis ultricies pellentesque pretium nullam quis risus eget urna "
    "mollis ornare vel eu leo pellentesque".split()
)

_SINGLE_LETTERS = frozenset("ABCDEFGHIJKLMNOPQRSTUVWXYZ")


def _is_placeholder(text: str | None) -> bool:
    if not text or not text.strip():
        return False
    t = text.strip()
    if t in _SINGLE_LETTERS:
        return True
    if t.lower() in _LIPSUM_WORDS:
        return True
    words = t.split()
    if len(words) >= 2 and all(w.lower() in _LIPSUM_WORDS for w in words if len(w) > 2):
        return True
    if t.endswith(".") and any(w.lower() in _LIPSUM_WORDS for w in words):
        return True
    if t.lower().startswith("step.") or t.lower().startswith("dummy"):
        return True
    return False


def _classify_text_slots(t_elems: list) -> list[int]:
    primary = []
    for i, t in enumerate(t_elems):
        text = t.text.strip() if t.text else ""
        if not _is_placeholder(text):
            continue
        if text in _SINGLE_LETTERS:
            continue
        words = text.split()
        if len(words) <= 3 and not text.endswith("."):
            primary.append(i)
    if not primary:
        seen = set()
        for i, t in enumerate(t_elems):
            text = t.text.strip() if t.text else ""
            if text and text[0] not in seen:
                seen.add(text[0])
                primary.append(i)
    return primary


def _shape_primary_slots(root, a_ns: str, p_ns: str) -> list[int]:
    all_t = list(root.iter(f"{{{a_ns}}}t"))
    nonempty = [t for t in all_t if t.text and t.text.strip()]
    t_to_idx = {id(t): i for i, t in enumerate(nonempty)}

    title_slots = []
    body_slots = []
    for sp in root.iter(f"{{{p_ns}}}sp"):
        shape_texts = [t for t in sp.iter(f"{{{a_ns}}}t") if t.text and t.text.strip()]
        if not shape_texts:
            continue
        full_text = " ".join(t.text.strip() for t in shape_texts)
        first_idx = t_to_idx.get(id(shape_texts[0]))
        if first_idx is None:
            continue
        if _is_title_text(full_text):
            title_slots.append(first_idx)
        else:
            body_slots.append(first_idx)

    return title_slots + body_slots


def _is_title_text(text: str) -> bool:
    t = text.strip()
    if len(t) <= 20 and not t.endswith("."):
        if t.lower().startswith("step"):
            return True
        if t.lower().startswith("dummy"):
            return True
        words = t.split()
        if len(words) <= 3 and not any(w.lower() in _LIPSUM_WORDS for w in words if len(w) > 2):
            return True
    return False


def _clear_shape_secondary(root, a_ns: str, p_ns: str, filled_indices: set, all_t: list) -> None:
    t_to_idx = {id(t): i for i, t in enumerate(all_t)}

    for sp in root.iter(f"{{{p_ns}}}sp"):
        shape_texts = [t for t in sp.iter(f"{{{a_ns}}}t") if t.text and t.text.strip()]
        if not shape_texts:
            continue
        first_idx = t_to_idx.get(id(shape_texts[0]))
        if first_idx is not None and first_idx in filled_indices:
            for t in shape_texts[1:]:
                if _is_placeholder(t.text):
                    t.text = ""


def _clear_unfilled_placeholders(root, a_ns: str, p_ns: str, filled_indices: set, all_t: list) -> None:
    t_to_idx = {id(t): i for i, t in enumerate(all_t)}

    for sp in root.iter(f"{{{p_ns}}}sp"):
        shape_texts = [t for t in sp.iter(f"{{{a_ns}}}t") if t.text and t.text.strip()]
        if not shape_texts:
            continue

        shape_has_filled = any(t_to_idx.get(id(t)) in filled_indices for t in shape_texts)
        if shape_has_filled:
            for t in shape_texts:
                idx = t_to_idx.get(id(t))
                if idx is not None and idx not in filled_indices:
                    t.text = ""
            continue

        full_text = " ".join(t.text.strip() for t in shape_texts)
        if _is_shape_placeholder(full_text):
            for t in shape_texts:
                t.text = ""


def _is_shape_placeholder(text: str) -> bool:
    t = text.strip()
    if not t:
        return False
    if t in _SINGLE_LETTERS:
        return True
    words = t.split()
    if len(words) >= 3:
        lipsum_count = sum(1 for w in words if w.lower().rstrip(".,;:!?") in _LIPSUM_WORDS)
        if lipsum_count / max(len(words), 1) > 0.3:
            return True
    if t.lower().startswith("step.") or t.lower().startswith("dummy"):
        return True
    return False


def _color_brightness(hex_color: str) -> float:
    h = hex_color.lstrip("#")
    if len(h) < 6:
        h = h[0]*2 + h[1]*2 + h[2]*2 if len(h) == 3 else h + "0"*(6-len(h))
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _hex_to_hsl(hex_val: str):
    import colorsys
    h = hex_val.lstrip("#")
    if len(h) < 6:
        h = h[0]*2 + h[1]*2 + h[2]*2 if len(h) == 3 else h + "0"*(6-len(h))
    r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
    hue, light, sat = colorsys.rgb_to_hls(r, g, b)
    return hue * 360, sat, light


def _generate_data_palette(primary_hex: str, count: int = 6, brand_spec: "BrandSpec | None" = None) -> list[str]:
    import colorsys
    if brand_spec and hasattr(brand_spec, "_dna_actual_colors") and brand_spec._dna_actual_colors:
        actual = brand_spec._dna_actual_colors
        sorted_by_bri = sorted(actual.keys(), key=lambda c: _color_brightness(c))
        mid_bri = sum(_color_brightness(c) for c in sorted_by_bri) / len(sorted_by_bri)
        candidates = sorted(sorted_by_bri, key=lambda c: abs(_color_brightness(c) - mid_bri))
        candidates = [c for c in candidates if 70 < _color_brightness(c) < 240]
        if len(candidates) < 3:
            candidates = sorted_by_bri
        if len(candidates) >= count:
            step = len(candidates) / count
            return [candidates[int(i * step)] for i in range(count)]
        if len(candidates) >= 2:
            result = list(candidates)
            while len(result) < count:
                result.append(candidates[len(result) % len(candidates)])
            return result[:count]
    hue, sat, lig = _hex_to_hsl(primary_hex)
    colors = []
    spread = 80
    for i in range(count):
        offset = (i - (count - 1) / 2) * (spread / max(count - 1, 1))
        new_h = (hue + offset) % 360
        new_s = max(0.30, min(0.75, sat + (i % 3 - 1) * 0.10))
        new_l = max(0.45, min(0.70, lig + 0.15 + (i % 2) * 0.10))
        r, g, b = colorsys.hls_to_rgb(new_h / 360, new_l, new_s)
        colors.append(f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}")
    return colors


_SCHEME_TO_SEMANTIC = {
    "dk1": "foreground", "lt1": "background", "dk2": "primary",
    "accent1": "accent", "accent2": "secondary",
    "accent3": "tertiary", "accent4": "accent",
    "accent5": "accent", "accent6": "accent",
    "hlink": "accent", "folHlink": "accent",
    "tx1": "foreground", "tx2": "muted-foreground",
    "bg1": "background", "bg2": "muted",
}


class ComponentRenderer:

    @staticmethod
    def compute_component_bounds(
        content_area: tuple[float, float, float, float],
        orig_aspect: float,
        fit: str = "contain",
    ) -> tuple[float, float, float, float]:
        """Compute target bounds for a component, preserving original aspect ratio.

        Args:
            content_area: (left, top, width, height) in inches — available space.
            orig_aspect: original width/height ratio of the component.
            fit: "contain" (letterbox, no crop), "width" (fill width, crop height),
                 "height" (fill height, crop width), "stretch" (ignore aspect).

        Returns:
            (left, top, width, height) in inches — target bounds.
        """
        area_left, area_top, area_w, area_h = content_area

        if orig_aspect <= 0 or area_w <= 0 or area_h <= 0:
            return content_area

        area_aspect = area_w / area_h

        if fit == "stretch":
            return content_area
        elif fit == "width":
            new_h = area_w / orig_aspect
            new_top = area_top + (area_h - new_h) / 2
            return (area_left, new_top, area_w, new_h)
        elif fit == "height":
            new_w = area_h * orig_aspect
            new_left = area_left + (area_w - new_w) / 2
            return (new_left, area_top, new_w, area_h)
        else:
            if orig_aspect > area_aspect:
                new_w = area_w
                new_h = area_w / orig_aspect
                new_top = area_top + (area_h - new_h) / 2
                return (area_left, new_top, new_w, new_h)
            else:
                new_h = area_h
                new_w = area_h * orig_aspect
                new_left = area_left + (area_w - new_w) / 2
                return (new_left, area_top, new_w, new_h)

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
        if brand_spec is None or brand_spec.colors is None:
            return xml_parts

        result = dict(xml_parts)

        if "colors" in xml_parts:
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
                semantic_key = _SCHEME_TO_SEMANTIC.get(val, val)
                matched_val = brand_colors.get(semantic_key) or brand_colors.get(val)
                if matched_val:
                    parent = scheme.getparent()
                    if parent is not None:
                        srgb = etree.SubElement(parent, f"{{{a_ns}}}srgbClr")
                        srgb.set("val", matched_val)
                        for child in scheme:
                            srgb.append(copy.deepcopy(child))
                        parent.remove(scheme)

            result["colors"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        if "group" in xml_parts:
            recolored = self._recolor_group_xml(xml_parts["group"], brand_spec)
            if brand_spec and brand_spec.fonts:
                recolored = self._replace_group_fonts(recolored, brand_spec)
            result["group"] = recolored

        return result

    def _recolor_group_xml(self, group_xml: bytes, brand_spec: BrandSpec) -> bytes:
        root = etree.fromstring(group_xml)
        a_ns = _NS["a"]

        brand_colors = {}
        if brand_spec.colors:
            for role, hex_val in brand_spec.colors.items():
                brand_colors[role] = hex_val.lstrip("#").upper()

        all_colors = {}
        for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
            val = srgb.get("val", "")
            if val:
                v = val.upper()
                if v not in all_colors:
                    all_colors[v] = {"count": 0, "as_shape_fill": 0, "as_text": 0,
                                     "as_gradient": 0, "as_theme_ref": 0}
                all_colors[v]["count"] += 1

                parent = srgb.getparent()
                if parent is not None:
                    p_tag = etree.QName(parent.tag).localname if isinstance(parent.tag, str) else ""
                    gp = parent.getparent()
                    gp_tag = etree.QName(gp.tag).localname if gp is not None and isinstance(gp.tag, str) else ""

                    if p_tag in ("lnRef", "fillRef", "effectRef"):
                        all_colors[v]["as_theme_ref"] += 1
                    elif p_tag == "gs":
                        all_colors[v]["as_gradient"] += 1
                    elif p_tag == "solidFill":
                        if gp_tag in ("rPr", "endParaRPr", "defRPr"):
                            all_colors[v]["as_text"] += 1
                        else:
                            all_colors[v]["as_shape_fill"] += 1

        neutral = {"000000", "000", "FFFFFF", "FFF"}
        non_neutral = {c for c in all_colors if c not in neutral}
        if not non_neutral:
            return group_xml

        shape_fills = {c: info for c, info in all_colors.items() if info["as_shape_fill"] > 0}

        is_data_heavy = len([c for c, info in shape_fills.items()
                             if info["as_shape_fill"] >= 3 and info["as_theme_ref"] == 0]) >= 4

        if not is_data_heavy:
            saturated_fills = [c for c, info in shape_fills.items()
                              if info["as_theme_ref"] == 0 and _hex_to_hsl(c)[1] > 0.15]
            distinct_hues = len(set(int(_hex_to_hsl(c)[0] / 30) for c in saturated_fills))
            if distinct_hues >= 3:
                is_data_heavy = True

        structural_dna = sorted([
            ("secondary", brand_colors.get("SECONDARY", "2E6504")),
            ("primary", brand_colors.get("PRIMARY", "466740")),
            ("muted-foreground", brand_colors.get("MUTED-FOREGROUND", "6B7D5A")),
            ("accent", brand_colors.get("ACCENT", "7EAB77")),
            ("tertiary", brand_colors.get("TERTIARY", "7DA92F")),
            ("quaternary", brand_colors.get("QUATERNARY", "D4E3AC")),
            ("muted", brand_colors.get("MUTED", "F2F8D6")),
            ("background", brand_colors.get("BACKGROUND", "FFFEF9")),
        ], key=lambda x: _color_brightness(x[1]))

        primary_hex = brand_colors.get("PRIMARY", "466740")
        data_palette = _generate_data_palette(primary_hex, max(len(shape_fills), 6), brand_spec)

        color_map = {}

        if is_data_heavy:
            structural_fills = []
            data_fills = []

            for c, info in all_colors.items():
                if c in neutral:
                    continue
                hue, sat, lig = _hex_to_hsl(c)
                is_theme = info.get("as_theme_ref", 0) > 0
                is_text_only = info.get("as_text", 0) > 0 and info.get("as_shape_fill", 0) == 0
                shape_fill_count = info.get("as_shape_fill", 0)

                if is_theme:
                    structural_fills.append(c)
                elif is_text_only:
                    structural_fills.append(c)
                elif shape_fill_count >= 1 and sat > 0.15 and 0.15 < lig < 0.9:
                    data_fills.append(c)
                elif shape_fill_count >= 1 and sat <= 0.15:
                    structural_fills.append(c)
                else:
                    structural_fills.append(c)

            sorted_structural = sorted(structural_fills, key=_color_brightness)
            for i, src_color in enumerate(sorted_structural):
                dna_idx = int(i / max(len(sorted_structural) - 1, 1) * (len(structural_dna) - 1))
                _, dna_hex = structural_dna[dna_idx]
                color_map[src_color] = dna_hex

            sorted_data = sorted(data_fills, key=_color_brightness)
            for i, src_color in enumerate(sorted_data):
                data_idx = i % len(data_palette)
                color_map[src_color] = data_palette[data_idx]
        else:
            sorted_colors = sorted(non_neutral, key=_color_brightness)
            for i, src_color in enumerate(sorted_colors):
                dna_idx = int(i / max(len(sorted_colors) - 1, 1) * (len(structural_dna) - 1))
                _, dna_hex = structural_dna[dna_idx]
                color_map[src_color] = dna_hex

        foreground = brand_colors.get("FOREGROUND", "0D4609")
        for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
            val = srgb.get("val", "")
            if not val:
                continue
            v = val.upper()
            if v in color_map:
                srgb.set("val", color_map[v])
            elif v in ("FFFFFF", "FFF"):
                parent = srgb.getparent()
                if parent is not None:
                    gp = parent.getparent()
                    if gp is not None:
                        gp_tag = etree.QName(gp.tag).localname if isinstance(gp.tag, str) else ""
                        if gp_tag in ("rPr", "endParaRPr", "defRPr"):
                            srgb.set("val", foreground)

        return etree.tostring(root, xml_declaration=False, encoding="UTF-8")

    @staticmethod
    def _replace_group_fonts(group_xml: bytes, brand_spec: BrandSpec) -> bytes:
        if not brand_spec.fonts:
            return group_xml
        root = etree.fromstring(group_xml)
        a_ns = _NS["a"]
        heading = brand_spec.fonts.get("heading", "")
        body = brand_spec.fonts.get("body", "")
        cjk_heading = brand_spec.fonts.get("cjk_heading", brand_spec.fonts.get("heading", ""))
        cjk_body = brand_spec.fonts.get("cjk_body", brand_spec.fonts.get("body", ""))
        for tag_name in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr", f"{{{a_ns}}}defRPr"):
            for rPr in root.iter(tag_name):
                sz = rPr.get("sz")
                is_heading = sz and int(sz) >= 2000 if sz else False
                latin_font = heading if (is_heading and heading) else (body or heading)
                cjk_font = cjk_heading if (is_heading and cjk_heading) else (cjk_body or cjk_heading)
                has_latin = False
                has_ea = False
                for latin in rPr.findall(f"{{{a_ns}}}latin"):
                    has_latin = True
                    if latin_font:
                        latin.set("typeface", latin_font)
                for ea in rPr.findall(f"{{{a_ns}}}ea"):
                    has_ea = True
                    if cjk_font:
                        ea.set("typeface", cjk_font)
                for cs in rPr.findall(f"{{{a_ns}}}cs"):
                    if cjk_font:
                        cs.set("typeface", cjk_font)
                if not has_latin and latin_font:
                    latin_elem = etree.SubElement(rPr, f"{{{a_ns}}}latin")
                    latin_elem.set("typeface", latin_font)
                if not has_ea and cjk_font:
                    ea_elem = etree.SubElement(rPr, f"{{{a_ns}}}ea")
                    ea_elem.set("typeface", cjk_font)
        return etree.tostring(root, xml_declaration=False, encoding="UTF-8")

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

                    grp_xml = styled.get("group", b"")
                    if grp_xml and brand_spec:
                        try:
                            grp_bytes = grp_xml if isinstance(grp_xml, bytes) else grp_xml.encode("utf-8")
                            font_brand = brand_spec
                            if not brand_spec.fonts:
                                font_brand = BrandSpec(
                                    colors=brand_spec.colors,
                                    fonts={"heading": "Microsoft YaHei", "body": "Microsoft YaHei",
                                           "cjk_heading": "Microsoft YaHei", "cjk_body": "Microsoft YaHei"},
                                )
                            grp_bytes = ComponentRenderer._replace_group_fonts(grp_bytes, font_brand)
                            styled = dict(styled)
                            styled["group"] = grp_bytes
                        except Exception:
                            pass

                    orig_aspect = 1.0
                    group_xml = styled.get("group", b"")
                    if group_xml:
                        try:
                            grp_root = etree.fromstring(group_xml if isinstance(group_xml, bytes) else group_xml.encode("utf-8"))
                            grpSpPr = grp_root.find(f"{{{_NS['p']}}}grpSpPr")
                            if grpSpPr is not None:
                                xfrm = grpSpPr.find(f"{{{_NS['a']}}}xfrm")
                                if xfrm is not None:
                                    chExt = xfrm.find(f"{{{_NS['a']}}}chExt")
                                    if chExt is not None:
                                        cx = int(chExt.get("cx", "0"))
                                        cy = int(chExt.get("cy", "0"))
                                        if cx > 0 and cy > 0:
                                            orig_aspect = cx / cy
                        except Exception:
                            pass

                    raw_bounds = element.get("bounds", (0.9, 1.6, 11.5, 5.0))
                    content_area = (raw_bounds[0], raw_bounds[1], raw_bounds[2], raw_bounds[3])
                    fit = "contain"
                    bounds = self.compute_component_bounds(content_area, orig_aspect, fit)

                    self._inject_group_to_slide(slide, styled, bounds, brand_spec=brand_spec, stretch=(fit == "stretch"))
                    return True

        return self._fallback_group(slide, element, brand_spec)

    def _fill_group_data(self, xml_parts: dict, new_texts: list) -> dict:
        if "group" not in xml_parts or not new_texts:
            return xml_parts

        group_xml = xml_parts["group"]
        root = etree.fromstring(group_xml)
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        all_t = [t for t in root.iter(f"{{{a_ns}}}t") if t.text and t.text.strip()]
        if not all_t:
            result = dict(xml_parts)
            result["group"] = etree.tostring(root, xml_declaration=False, encoding="UTF-8")
            return result

        n_slots = len(all_t)
        n_texts = len(new_texts)

        if n_texts >= n_slots:
            for i, t_elem in enumerate(all_t):
                new_text = new_texts[i]
                if isinstance(new_text, dict):
                    new_text = new_text.get("text", "")
                t_elem.text = str(new_text)
        else:
            shape_primary = _shape_primary_slots(root, a_ns, p_ns)

            if len(shape_primary) > 0:
                filled_indices = set()
                used_slots = set()
                for i, new_text in enumerate(new_texts):
                    idx = i * len(shape_primary) // n_texts
                    idx = min(idx, len(shape_primary) - 1)
                    while idx in used_slots and idx < len(shape_primary) - 1:
                        idx += 1
                    used_slots.add(idx)
                    slot = shape_primary[idx]
                    if isinstance(new_text, dict):
                        new_text = new_text.get("text", "")
                    all_t[slot].text = str(new_text)
                    filled_indices.add(slot)

                _clear_unfilled_placeholders(root, a_ns, p_ns, filled_indices, all_t)
            else:
                step = n_slots / n_texts if n_texts > 0 else 1
                for i, new_text in enumerate(new_texts):
                    idx = min(int(i * step), n_slots - 1)
                    if isinstance(new_text, dict):
                        new_text = new_text.get("text", "")
                    all_t[idx].text = str(new_text)

        result = dict(xml_parts)
        result["group"] = etree.tostring(root, xml_declaration=False, encoding="UTF-8")
        return result

    def _inject_group_to_slide(self, slide, xml_parts: dict, bounds: tuple, brand_spec: BrandSpec | None = None,
                               stretch: bool = False) -> None:
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
                cNv.set("name", f"Component {next_id}")
                next_id += 1

            for cNv in grp_elem.iter(f"{{{ns_a}}}cNvPr"):
                if cNv.getparent() is not None and cNv.getparent().tag != f"{{{ns_p}}}nvGrpSpPr":
                    cNv.set("id", str(next_id))
                    cNv.set("name", f"Element {next_id}")
                    next_id += 1

            target_left = bounds[0] if bounds and len(bounds) > 0 else 0.9
            target_top = bounds[1] if bounds and len(bounds) > 1 else 1.5
            target_w = bounds[2] if bounds and len(bounds) > 2 else 11.5
            target_h = bounds[3] if bounds and len(bounds) > 3 else 5.2

            self._remove_unresolvable_references(grp_elem)
            self._ensure_shape_fills(grp_elem, brand_spec=brand_spec)

            self._denormalize_coordinates(grp_elem, target_left, target_top, target_w, target_h, stretch=stretch)

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

                    from pptx.opc.packuri import PackURI
                    img_part = Part(
                        partname=PackURI(f"/ppt/media/library_img_{original_rId}.png"),
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

            new_grp = etree.SubElement(sp_tree, f"{{{ns_p}}}grpSp")
            for child in grp_elem:
                new_grp.append(child)

        except Exception as exc:
            import logging
            logging.getLogger(__name__).warning("_inject_group_to_slide failed: %s", exc)

    def _denormalize_coordinates(self, grp_elem, left: float, top: float, width: float, height: float,
                                 orig_bounds_emu: tuple[int, int, int, int] | None = None,
                                 stretch: bool = False) -> None:
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        target_left_emu = int(left * 914400)
        target_top_emu = int(top * 914400)
        target_w_emu = int(width * 914400)
        target_h_emu = int(height * 914400)

        grpSpPr = grp_elem.find(f"{{{p_ns}}}grpSpPr")
        if grpSpPr is None:
            return
        grp_xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
        if grp_xfrm is None:
            return

        grp_chOff = grp_xfrm.find(f"{{{a_ns}}}chOff")
        grp_chExt = grp_xfrm.find(f"{{{a_ns}}}chExt")

        if grp_chOff is not None and grp_chExt is not None:
            try:
                orig_chOff_x = int(grp_chOff.get("x", "0"))
                orig_chOff_y = int(grp_chOff.get("y", "0"))
                orig_chExt_cx = int(grp_chExt.get("cx", "0"))
                orig_chExt_cy = int(grp_chExt.get("cy", "0"))
            except (ValueError, TypeError):
                return

            if orig_chExt_cx <= 0 or orig_chExt_cy <= 0:
                return

            scale_x = target_w_emu / orig_chExt_cx
            scale_y = target_h_emu / orig_chExt_cy

            if stretch:
                use_sx = scale_x
                use_sy = scale_y
                offset_x = 0
                offset_y = 0
            else:
                uniform_scale = min(scale_x, scale_y)
                use_sx = uniform_scale
                use_sy = uniform_scale
                actual_w = int(orig_chExt_cx * uniform_scale)
                actual_h = int(orig_chExt_cy * uniform_scale)
                offset_x = (target_w_emu - actual_w) // 2
                offset_y = (target_h_emu - actual_h) // 2

            grp_off = grp_xfrm.find(f"{{{a_ns}}}off")
            grp_ext = grp_xfrm.find(f"{{{a_ns}}}ext")
            if grp_off is not None:
                grp_off.set("x", str(target_left_emu + offset_x))
                grp_off.set("y", str(target_top_emu + offset_y))
            if grp_ext is not None:
                grp_ext.set("cx", str(target_w_emu))
                grp_ext.set("cy", str(target_h_emu))

            new_chOff_x = int((orig_chOff_x - orig_chOff_x) * use_sx)
            new_chOff_y = int((orig_chOff_y - orig_chOff_y) * use_sy)
            new_chExt_cx = int(orig_chExt_cx * use_sx)
            new_chExt_cy = int(orig_chExt_cy * use_sy)

            grp_chOff.set("x", str(new_chOff_x))
            grp_chOff.set("y", str(new_chOff_y))
            grp_chExt.set("cx", str(new_chExt_cx))
            grp_chExt.set("cy", str(new_chExt_cy))

            self._transform_child_elements(grp_elem, orig_chOff_x, orig_chOff_y, use_sx, use_sy)

            for ln in grp_elem.iter(f"{{{a_ns}}}ln"):
                w = ln.get("w")
                if w:
                    try:
                        new_w = int(int(w) * min(use_sx, use_sy))
                        if new_w > 0:
                            ln.set("w", str(new_w))
                    except (ValueError, TypeError):
                        pass

            for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr"):
                for rpr in grp_elem.iter(rpr_tag):
                    sz = rpr.get("sz")
                    if sz:
                        try:
                            new_sz = max(1400, min(7200, int(int(sz) * min(use_sy, 1.0))))
                            rpr.set("sz", str(new_sz))
                        except (ValueError, TypeError):
                            pass

            for bodyPr in grp_elem.iter(f"{{{a_ns}}}bodyPr"):
                for attr in ("lIns", "tIns", "rIns", "bIns"):
                    val = bodyPr.get(attr)
                    if val:
                        try:
                            new_val = int(int(val) * min(use_sx, use_sy))
                            if new_val >= 0:
                                bodyPr.set(attr, str(new_val))
                        except (ValueError, TypeError):
                            pass
        else:
            grp_off = grp_xfrm.find(f"{{{a_ns}}}off")
            grp_ext = grp_xfrm.find(f"{{{a_ns}}}ext")
            if grp_off is not None:
                grp_off.set("x", str(target_left_emu))
                grp_off.set("y", str(target_top_emu))
            if grp_ext is not None:
                grp_ext.set("cx", str(target_w_emu))
                grp_ext.set("cy", str(target_h_emu))

    def _transform_child_elements(self, grp_elem, chOff_x: int, chOff_y: int, scale_x: float, scale_y: float) -> None:
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        for child in grp_elem:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag in ("sp", "pic", "cxnSp", "graphicFrame"):
                spPr = child.find(f"{{{p_ns}}}spPr")
                if spPr is None:
                    continue
                xfrm = spPr.find(f"{{{a_ns}}}xfrm")
                if xfrm is None:
                    continue
                off = xfrm.find(f"{{{a_ns}}}off")
                if off is not None:
                    try:
                        orig_x = int(off.get("x", "0"))
                        orig_y = int(off.get("y", "0"))
                        off.set("x", str(int((orig_x - chOff_x) * scale_x)))
                        off.set("y", str(int((orig_y - chOff_y) * scale_y)))
                    except (ValueError, TypeError):
                        pass
                ext = xfrm.find(f"{{{a_ns}}}ext")
                if ext is not None:
                    try:
                        orig_cx = int(ext.get("cx", "0"))
                        orig_cy = int(ext.get("cy", "0"))
                        ext.set("cx", str(int(orig_cx * scale_x)))
                        ext.set("cy", str(int(orig_cy * scale_y)))
                    except (ValueError, TypeError):
                        pass

            elif tag == "grpSp":
                sub_grpSpPr = child.find(f"{{{p_ns}}}grpSpPr")
                if sub_grpSpPr is None:
                    continue
                sub_xfrm = sub_grpSpPr.find(f"{{{a_ns}}}xfrm")
                if sub_xfrm is None:
                    continue

                sub_off = sub_xfrm.find(f"{{{a_ns}}}off")
                if sub_off is not None:
                    try:
                        orig_x = int(sub_off.get("x", "0"))
                        orig_y = int(sub_off.get("y", "0"))
                        sub_off.set("x", str(int((orig_x - chOff_x) * scale_x)))
                        sub_off.set("y", str(int((orig_y - chOff_y) * scale_y)))
                    except (ValueError, TypeError):
                        pass

                sub_ext = sub_xfrm.find(f"{{{a_ns}}}ext")
                if sub_ext is not None:
                    try:
                        orig_cx = int(sub_ext.get("cx", "0"))
                        orig_cy = int(sub_ext.get("cy", "0"))
                        new_cx = int(orig_cx * scale_x)
                        new_cy = int(orig_cy * scale_y)
                        sub_ext.set("cx", str(new_cx))
                        sub_ext.set("cy", str(new_cy))
                    except (ValueError, TypeError):
                        pass

                sub_chOff = sub_xfrm.find(f"{{{a_ns}}}chOff")
                sub_chExt = sub_xfrm.find(f"{{{a_ns}}}chExt")
                if sub_chOff is not None and sub_chExt is not None:
                    try:
                        sub_chOff_x = int(sub_chOff.get("x", "0"))
                        sub_chOff_y = int(sub_chOff.get("y", "0"))
                        sub_chExt_cx = int(sub_chExt.get("cx", "0"))
                        sub_chExt_cy = int(sub_chExt.get("cy", "0"))

                        new_chOff_x = int(sub_chOff_x * scale_x)
                        new_chOff_y = int(sub_chOff_y * scale_y)
                        new_chExt_cx = int(sub_chExt_cx * scale_x)
                        new_chExt_cy = int(sub_chExt_cy * scale_y)

                        sub_chOff.set("x", str(new_chOff_x))
                        sub_chOff.set("y", str(new_chOff_y))
                        sub_chExt.set("cx", str(new_chExt_cx))
                        sub_chExt.set("cy", str(new_chExt_cy))

                        sub_scale_x = new_cx / new_chExt_cx if new_chExt_cx > 0 else 1.0
                        sub_scale_y = new_cy / new_chExt_cy if new_chExt_cy > 0 else 1.0

                        self._transform_child_elements(child, new_chOff_x, new_chOff_y, sub_scale_x, sub_scale_y)
                    except (ValueError, TypeError):
                        pass
                else:
                    self._transform_child_elements(child, chOff_x, chOff_y, scale_x, scale_y)

    def _strip_3d_effects(self, grp_elem) -> None:
        a_ns = _NS["a"]

        for tag in [f"{{{a_ns}}}scene3d", f"{{{a_ns}}}sp3d"]:
            for elem in list(grp_elem.iter(tag)):
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)

        for tag in [f"{{{a_ns}}}contourClr", f"{{{a_ns}}}bevelT", f"{{{a_ns}}}bevelB",
                    f"{{{a_ns}}}extrusionClr", f"{{{a_ns}}}prstMat"]:
            for elem in list(grp_elem.iter(tag)):
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)

    def _ensure_shape_fills(self, grp_elem, brand_spec: BrandSpec | None = None) -> None:
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        for sp in grp_elem.iter(f"{{{p_ns}}}sp"):
            spPr = sp.find(f"{{{p_ns}}}spPr")
            if spPr is None:
                continue
            has_fill = False
            for tag in ("solidFill", "gradFill", "pattFill", "noFill"):
                if spPr.find(f"{{{a_ns}}}{tag}") is not None:
                    has_fill = True
                    break
            if not has_fill:
                xfrm = spPr.find(f"{{{a_ns}}}xfrm")
                if xfrm is not None:
                    ext = xfrm.find(f"{{{a_ns}}}ext")
                    if ext is not None:
                        cx = int(ext.get("cx", "0"))
                        cy = int(ext.get("cy", "0"))
                        if cx > 50000 and cy > 50000:
                            prstGeom = spPr.find(f"{{{a_ns}}}prstGeom")
                            if prstGeom is not None:
                                idx = list(spPr).index(prstGeom) + 1
                                solidFill = etree.SubElement(spPr, f"{{{a_ns}}}solidFill")
                                srgbClr = etree.SubElement(solidFill, f"{{{a_ns}}}srgbClr")
                                muted_color = "F1F5F9"
                                if brand_spec and brand_spec.colors and brand_spec.colors.get("muted"):
                                    muted_color = brand_spec.colors["muted"].lstrip("#").upper()
                                srgbClr.set("val", muted_color)
                                spPr.insert(idx, solidFill)

    def _remove_unresolvable_references(self, grp_elem) -> None:
        p_ns = _NS["p"]
        a_ns = _NS["a"]

        to_remove = []
        for child in grp_elem.iter(f"{{{p_ns}}}graphicFrame"):
            to_remove.append(child)

        for gf in to_remove:
            parent = gf.getparent()
            if parent is not None:
                parent.remove(gf)

        for blip in list(grp_elem.iter(f"{{{a_ns}}}blip")):
            embed = blip.get(f"{{{_NS['r']}}}embed")
            if embed and not embed.startswith("rId"):
                parent = blip.getparent()
                while parent is not None:
                    grandparent = parent.getparent()
                    if grandparent is not None and grandparent.tag == f"{{{p_ns}}}sp":
                        grandparent.remove(parent)
                        break
                    parent = grandparent

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
