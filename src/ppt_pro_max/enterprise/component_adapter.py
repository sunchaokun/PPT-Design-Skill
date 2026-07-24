"""ComponentAdapter — 组件自适应适配器。

4 阶段管线让任意组件在任意品牌规范下正确显示:
  1. analyze()   — 解析组件结构（颜色语义、字号层级、形状密度）
  2. plan()      — 制定适配策略（颜色映射表、字号映射表、bounds 策略）
  3. transform() — 执行变换（颜色替换、字体替换、坐标缩放）
  4. validate()  — 验证结果（对比度检查、溢出检查、最小字号检查）
"""

from __future__ import annotations

import colorsys
import copy
from dataclasses import dataclass, field
from typing import Any

from lxml import etree

from ppt_pro_max.enterprise.brand_spec import BrandSpec


_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

_SCHEME_TO_SEMANTIC = {
    "dk1": "foreground", "lt1": "background", "dk2": "primary",
    "accent1": "accent", "accent2": "secondary",
    "accent3": "tertiary", "accent4": "accent",
    "accent5": "accent", "accent6": "accent",
    "hlink": "accent", "folHlink": "accent",
    "tx1": "foreground", "tx2": "muted-foreground",
    "bg1": "background", "bg2": "muted",
}

_MIN_FONT_PT = 11
_MIN_INSET_EMU = 36000


# ── Data structures ──

@dataclass
class ColorRole:
    hex_val: str
    role: str            # "dominant_fill" | "secondary_fill" | "data_fill" | "text" | "theme_ref"
    contexts: list[str]  # "shape_fill" | "text" | "gradient" | "theme_ref"
    count: int = 0
    index: int = 0       # for data colors ordering


@dataclass
class FontLevel:
    sz_pt: float
    role: str            # "title" | "subtitle" | "body" | "caption"
    count: int = 1


@dataclass
class FontMapping:
    latin: str
    cjk: str
    sz_pt: float


@dataclass
class ComponentAnalysis:
    color_roles: dict[str, ColorRole] = field(default_factory=dict)
    has_dark_bg: bool = False
    color_count: int = 0
    gradient_count: int = 0
    font_levels: list[FontLevel] = field(default_factory=list)
    has_cjk: bool = False
    min_font_pt: float = 14.0
    max_font_pt: float = 14.0
    shape_count: int = 0
    text_shape_count: int = 0
    text_density: float = 0.0
    aspect_ratio: float = 1.0
    orig_bounds_emu: tuple[int, int, int, int] = (0, 0, 9144000, 5486400)


@dataclass
class AdaptationPlan:
    color_map: dict[str, str] = field(default_factory=dict)
    gradient_color_map: dict[str, str] = field(default_factory=dict)
    font_map: dict[str, FontMapping] = field(default_factory=dict)
    fit_strategy: str = "contain"
    target_bounds: tuple[float, float, float, float] = (0.9, 1.6, 11.5, 5.0)
    min_font_pt: float = _MIN_FONT_PT
    force_fg_on_light_bg: bool = False
    brand_colors: dict[str, str] = field(default_factory=dict)
    is_dark_brand: bool = False


# ── ComponentAdapter ──

class ComponentAdapter:

    def adapt(self, xml_parts: dict, element: dict, brand_spec: BrandSpec | None = None) -> dict:
        """Main entry: adapt component XML to brand spec."""
        if brand_spec is None:
            brand_spec = BrandSpec()

        result = dict(xml_parts)

        if "group" not in result:
            return result

        analysis = self.analyze(result, element)

        plan = self.plan(analysis, brand_spec, element)

        result = self.transform(result, plan, analysis)

        issues = self.validate(result, brand_spec, plan)

        result["_adapted_bounds"] = plan.target_bounds
        result["_fit_strategy"] = plan.fit_strategy
        result["_validation_issues"] = issues

        return result

    # ── Phase 1: Analyze ──

    def analyze(self, xml_parts: dict, element: dict | None = None) -> ComponentAnalysis:
        """Extract structured info from component XML."""
        analysis = ComponentAnalysis()

        group_xml = xml_parts.get("group", b"")
        if not group_xml:
            return analysis

        if isinstance(group_xml, str):
            group_xml = group_xml.encode("utf-8")

        try:
            root = etree.fromstring(group_xml)
        except Exception:
            return analysis

        a_ns = _NS["a"]
        p_ns = _NS["p"]

        # Color analysis
        color_info: dict[str, dict] = {}
        for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
            val = srgb.get("val", "").upper()
            if not val:
                continue
            if val not in color_info:
                color_info[val] = {"count": 0, "contexts": []}
            color_info[val]["count"] += 1

            parent = srgb.getparent()
            context = self._classify_color_context(srgb, parent, root)
            color_info[val]["contexts"].append(context)

        # Determine dominant fill vs data fill vs text
        fill_counts: dict[str, int] = {}
        text_counts: dict[str, int] = {}
        gradient_counts: dict[str, int] = {}
        theme_ref_counts: dict[str, int] = {}

        for val, info in color_info.items():
            for ctx in info["contexts"]:
                if ctx == "shape_fill":
                    fill_counts[val] = fill_counts.get(val, 0) + 1
                elif ctx == "text":
                    text_counts[val] = text_counts.get(val, 0) + 1
                elif ctx == "gradient":
                    gradient_counts[val] = gradient_counts.get(val, 0) + 1
                elif ctx == "theme_ref":
                    theme_ref_counts[val] = theme_ref_counts.get(val, 0) + 1

        total_shapes = max(1, len(list(root.iter(f"{{{p_ns}}}sp"))))
        data_idx = 0
        for val, info in color_info.items():
            n_fill = fill_counts.get(val, 0)
            n_text = text_counts.get(val, 0)
            n_grad = gradient_counts.get(val, 0)
            n_theme = theme_ref_counts.get(val, 0)

            if n_theme > 0:
                role = "theme_ref"
            elif n_text > 0 and n_fill == 0:
                role = "text"
            elif n_fill >= total_shapes * 0.5:
                role = "dominant_fill"
            elif n_fill > 0 and n_fill < total_shapes * 0.3:
                role = "data_fill"
            elif n_fill > 0:
                role = "secondary_fill"
            else:
                role = "secondary_fill"

            analysis.color_roles[val] = ColorRole(
                hex_val=val,
                role=role,
                contexts=info["contexts"],
                count=info["count"],
                index=data_idx if role == "data_fill" else 0,
            )
            if role == "data_fill":
                data_idx += 1

        analysis.color_count = len([v for v in analysis.color_roles.values() if v.role != "theme_ref"])
        analysis.gradient_count = len(list(root.iter(f"{{{a_ns}}}gradFill")))

        # Dark background detection
        bg_colors = []
        for sp in root.iter(f"{{{p_ns}}}sp"):
            spPr = sp.find(f"{{{p_ns}}}spPr")
            if spPr is None:
                continue
            xfrm = spPr.find(f"{{{a_ns}}}xfrm")
            if xfrm is None:
                continue
            off = xfrm.find(f"{{{a_ns}}}off")
            ext = xfrm.find(f"{{{a_ns}}}ext")
            if off is not None and ext is not None:
                x = int(off.get("x", "0"))
                y = int(off.get("y", "0"))
                cx = int(ext.get("cx", "0"))
                cy = int(ext.get("cy", "0"))
                if x <= 100000 and y <= 100000 and cx > 5000000:
                    solidFill = spPr.find(f"{{{a_ns}}}solidFill")
                    if solidFill is not None:
                        srgb = solidFill.find(f"{{{a_ns}}}srgbClr")
                        if srgb is not None:
                            bg_colors.append(srgb.get("val", "").upper())

        if bg_colors:
            avg_brightness = sum(self._brightness(c) for c in bg_colors) / len(bg_colors)
            analysis.has_dark_bg = avg_brightness < 128

        # Font analysis
        all_sz: list[int] = []
        has_ea = False
        for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr", f"{{{a_ns}}}defRPr"):
            for rpr in root.iter(rpr_tag):
                sz = rpr.get("sz")
                if sz:
                    try:
                        all_sz.append(int(sz))
                    except ValueError:
                        pass
                ea = rpr.find(f"{{{a_ns}}}ea")
                if ea is not None and ea.get("typeface", ""):
                    has_ea = True

        analysis.has_cjk = has_ea

        if all_sz:
            all_pt = [s / 100 for s in all_sz]
            analysis.min_font_pt = min(all_pt)
            analysis.max_font_pt = max(all_pt)
            analysis.font_levels = self._infer_font_levels(all_pt)

        # Shape and text analysis
        sp_count = 0
        text_sp_count = 0
        total_text_len = 0
        for sp in root.iter(f"{{{p_ns}}}sp"):
            sp_count += 1
            sp_texts = [t for t in sp.iter(f"{{{a_ns}}}t") if t.text and t.text.strip()]
            if sp_texts:
                text_sp_count += 1
                total_text_len += sum(len(t.text) for t in sp_texts)

        analysis.shape_count = sp_count
        analysis.text_shape_count = text_sp_count
        analysis.text_density = total_text_len / max(sp_count, 1)

        # Aspect ratio
        grpSpPr = root.find(f"{{{p_ns}}}grpSpPr")
        if grpSpPr is not None:
            xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
            if xfrm is not None:
                chExt = xfrm.find(f"{{{a_ns}}}chExt")
                off_elem = xfrm.find(f"{{{a_ns}}}off")
                ext_elem = xfrm.find(f"{{{a_ns}}}ext")
                if chExt is not None:
                    cx = int(chExt.get("cx", "0"))
                    cy = int(chExt.get("cy", "0"))
                    if cx > 0 and cy > 0:
                        analysis.aspect_ratio = cx / cy
                        analysis.orig_bounds_emu = (
                            int(off_elem.get("x", "0")) if off_elem is not None else 0,
                            int(off_elem.get("y", "0")) if off_elem is not None else 0,
                            int(ext_elem.get("cx", "0")) if ext_elem is not None else cx,
                            int(ext_elem.get("cy", "0")) if ext_elem is not None else cy,
                        )

        return analysis

    def _classify_color_context(self, srgb_elem, parent, root) -> str:
        """Classify where a srgbClr is used."""
        if parent is None:
            return "unknown"
        p_tag = etree.QName(parent.tag).localname if isinstance(parent.tag, str) else ""
        gp = parent.getparent()
        gp_tag = etree.QName(gp.tag).localname if gp is not None and isinstance(gp.tag, str) else ""

        if p_tag in ("lnRef", "fillRef", "effectRef"):
            return "theme_ref"
        if p_tag == "gs":
            return "gradient"
        if p_tag == "solidFill":
            if gp_tag in ("rPr", "endParaRPr", "defRPr"):
                return "text"
            return "shape_fill"
        return "unknown"

    def _infer_font_levels(self, all_pt: list[float]) -> list[FontLevel]:
        """Infer font level hierarchy by clustering."""
        sorted_pt = sorted(set(all_pt), reverse=True)
        if not sorted_pt:
            return [FontLevel(sz_pt=14, role="body", count=1)]
        if len(sorted_pt) == 1:
            return [FontLevel(sz_pt=sorted_pt[0], role="body", count=len(all_pt))]

        gaps = [(sorted_pt[i] - sorted_pt[i + 1], i) for i in range(len(sorted_pt) - 1)]
        gaps.sort(reverse=True)

        n_boundaries = min(3, len(gaps))
        boundaries = sorted([g[1] for g in gaps[:n_boundaries]])

        levels: list[FontLevel] = []
        role_names = ["title", "subtitle", "body", "caption"]
        prev = 0
        for i, b in enumerate(boundaries + [len(sorted_pt) - 1]):
            level_pt = sorted_pt[prev]
            count = sum(1 for p in all_pt if sorted_pt[prev] >= p > (sorted_pt[b + 1] if b + 1 < len(sorted_pt) else 0))
            levels.append(FontLevel(
                sz_pt=level_pt,
                role=role_names[min(i, len(role_names) - 1)],
                count=max(count, 1),
            ))
            prev = b + 1

        return levels

    # ── Phase 2: Plan ──

    def plan(self, analysis: ComponentAnalysis, brand_spec: BrandSpec, element: dict) -> AdaptationPlan:
        """Build adaptation plan from analysis and brand spec."""
        brand_colors = {}
        if brand_spec.colors:
            for k, v in brand_spec.colors.items():
                brand_colors[k] = v.lstrip("#").upper()

        is_dark = self._is_brand_dark(brand_spec)

        plan = AdaptationPlan(
            brand_colors=brand_colors,
            is_dark_brand=is_dark,
            force_fg_on_light_bg=(not is_dark and analysis.has_dark_bg),
            min_font_pt=_MIN_FONT_PT,
        )

        # Color map
        plan.color_map = self._build_color_map(analysis, brand_colors, is_dark)
        plan.gradient_color_map = plan.color_map

        # Font map
        plan.font_map = self._build_font_map(analysis, brand_spec)

        # Layout
        raw_bounds = element.get("bounds", (0.9, 1.6, 11.5, 5.0))
        plan.target_bounds = tuple(float(v) for v in raw_bounds)
        plan.fit_strategy = self._determine_fit_strategy(analysis, element)

        return plan

    def _build_color_map(self, analysis: ComponentAnalysis, brand_colors: dict[str, str], is_dark: bool) -> dict[str, str]:
        """Build semantic-aware color mapping."""
        color_map: dict[str, str] = {}

        primary = brand_colors.get("PRIMARY") or brand_colors.get("primary", "2563EB")
        secondary = brand_colors.get("SECONDARY") or brand_colors.get("secondary", "7C3AED")
        muted = brand_colors.get("MUTED") or brand_colors.get("muted", "F1F5F9")
        background = brand_colors.get("BACKGROUND") or brand_colors.get("background", "FFFFFF")
        foreground = brand_colors.get("FOREGROUND") or brand_colors.get("foreground", "0F172A")
        on_primary = brand_colors.get("ON-PRIMARY") or brand_colors.get("on-primary", "FFFFFF")
        muted_fg = brand_colors.get("MUTED-FOREGROUND") or brand_colors.get("muted-foreground", "64748B")
        accent = brand_colors.get("ACCENT") or brand_colors.get("accent", "F97316")

        data_palette = self._generate_data_palette(primary, max(analysis.color_count, 6))

        data_idx = 0
        for val, role_info in analysis.color_roles.items():
            if role_info.role == "theme_ref":
                continue

            if role_info.role == "dominant_fill":
                target = primary
            elif role_info.role == "secondary_fill":
                target = muted
            elif role_info.role == "data_fill":
                target = data_palette[data_idx % len(data_palette)]
                data_idx += 1
            elif role_info.role == "text":
                if is_dark:
                    target = foreground
                else:
                    if self._brightness(val) > 200:
                        target = foreground
                    else:
                        target = val
            else:
                target = val

            color_map[val] = target.upper()

        # White text fix: only on light brand when component had dark bg
        if not is_dark and analysis.has_dark_bg:
            for val, role_info in analysis.color_roles.items():
                if role_info.role == "text" and val in ("FFFFFF", "FFF"):
                    color_map[val] = foreground

        return color_map

    def _generate_data_palette(self, primary_hex: str, count: int) -> list[str]:
        """Generate harmonious data colors from primary."""
        h = primary_hex.lstrip("#")
        hue, sat, lig = self._hex_to_hsl(h)
        colors = []
        spread = 100
        for i in range(count):
            offset = (i - (count - 1) / 2) * (spread / max(count - 1, 1))
            new_h = (hue + offset) % 360
            new_s = max(0.35, min(0.75, sat + (i % 3 - 1) * 0.08))
            new_l = max(0.45, min(0.68, lig + 0.12 + (i % 2) * 0.08))
            r, g, b = colorsys.hls_to_rgb(new_h / 360, new_l, new_s)
            colors.append(f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}")
        return colors

    def _build_font_map(self, analysis: ComponentAnalysis, brand_spec: BrandSpec) -> dict[str, FontMapping]:
        """Build level-aware font mapping."""
        brand_fonts = brand_spec.fonts or {}
        heading = brand_fonts.get("heading", "Inter")
        body = brand_fonts.get("body", "Inter")
        cjk_heading = brand_fonts.get("cjk_heading", "Microsoft YaHei")
        cjk_body = brand_fonts.get("cjk_body", "Microsoft YaHei")

        font_map: dict[str, FontMapping] = {}
        for level in analysis.font_levels:
            if level.role in ("title", "subtitle"):
                font_map[level.role] = FontMapping(latin=heading, cjk=cjk_heading, sz_pt=level.sz_pt)
            else:
                font_map[level.role] = FontMapping(latin=body, cjk=cjk_body, sz_pt=level.sz_pt)

        if "body" not in font_map:
            font_map["body"] = FontMapping(latin=body, cjk=cjk_body, sz_pt=14)

        return font_map

    def _determine_fit_strategy(self, analysis: ComponentAnalysis, element: dict) -> str:
        """Choose best fit strategy based on component characteristics."""
        if analysis.text_density > 30:
            return "width"
        if analysis.aspect_ratio > 2.5:
            return "width"
        if analysis.aspect_ratio < 0.5:
            return "height"
        if analysis.shape_count > 20:
            return "contain"
        return "contain"

    # ── Phase 3: Transform ──

    def transform(self, xml_parts: dict, plan: AdaptationPlan, analysis: ComponentAnalysis) -> dict:
        """Execute color, font, and layout transformations."""
        result = dict(xml_parts)

        if "group" not in result:
            return result

        group_xml = result["group"]
        if isinstance(group_xml, str):
            group_xml = group_xml.encode("utf-8")

        try:
            root = etree.fromstring(group_xml)
        except Exception:
            return result

        a_ns = _NS["a"]
        p_ns = _NS["p"]

        # Step 1: Color transform
        self._transform_colors(root, plan, analysis)

        # Step 2: Font transform
        self._transform_fonts(root, plan, analysis)

        # Step 3: Layout transform
        self._transform_layout(root, plan, analysis)

        result["group"] = etree.tostring(root, xml_declaration=False, encoding="UTF-8")
        return result

    def _transform_colors(self, root, plan: AdaptationPlan, analysis: ComponentAnalysis):
        """Transform colors: srgbClr replacement, gradient handling, white text fix."""
        a_ns = _NS["a"]

        # Replace srgbClr values
        for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
            val = srgb.get("val", "").upper()
            if val in plan.color_map:
                srgb.set("val", plan.color_map[val])

        # Replace schemeClr (not inside style elements)
        for scheme in list(root.iter(f"{{{a_ns}}}schemeClr")):
            if self._is_inside_style(scheme):
                continue
            val = scheme.get("val", "")
            semantic_key = _SCHEME_TO_SEMANTIC.get(val, val)
            brand_val = plan.brand_colors.get(semantic_key.upper()) or plan.brand_colors.get(semantic_key)
            if not brand_val:
                brand_val = plan.brand_colors.get(val.upper()) or plan.brand_colors.get(val)
            if brand_val:
                parent = scheme.getparent()
                if parent is not None:
                    idx = list(parent).index(scheme)
                    srgb = etree.Element(f"{{{a_ns}}}srgbClr")
                    srgb.set("val", brand_val.upper())
                    for child in scheme:
                        srgb.append(copy.deepcopy(child))
                    parent.remove(scheme)
                    parent.insert(idx, srgb)

        # White text fix on light background
        if plan.force_fg_on_light_bg:
            fg_hex = plan.brand_colors.get("FOREGROUND") or plan.brand_colors.get("foreground", "0F172A")
            for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
                val = srgb.get("val", "").upper()
                if val in ("FFFFFF", "FFF"):
                    parent = srgb.getparent()
                    if parent is not None:
                        gp = parent.getparent()
                        if gp is not None:
                            gp_tag = etree.QName(gp.tag).localname if isinstance(gp.tag, str) else ""
                            if gp_tag in ("rPr", "endParaRPr", "defRPr"):
                                srgb.set("val", fg_hex.upper())

    def _is_inside_style(self, elem) -> bool:
        """Check if element is inside an <a:style> element."""
        p = elem.getparent()
        while p is not None:
            if p.tag == f"{{{_NS['a']}}}style":
                return True
            p = p.getparent()
        return False

    def _transform_fonts(self, root, plan: AdaptationPlan, analysis: ComponentAnalysis):
        """Transform fonts: level-aware replacement, CJK addition."""
        a_ns = _NS["a"]

        for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr", f"{{{a_ns}}}defRPr"):
            for rpr in root.iter(rpr_tag):
                sz = rpr.get("sz")
                if sz:
                    try:
                        sz_pt = int(sz) / 100
                        level_role = self._find_closest_level(sz_pt, analysis.font_levels)
                    except (ValueError, TypeError):
                        level_role = "body"
                else:
                    level_role = "body"

                font_mapping = plan.font_map.get(level_role, plan.font_map.get("body"))
                if font_mapping is None:
                    continue

                latin_font = font_mapping.latin
                cjk_font = font_mapping.cjk

                has_latin = False
                for latin in rpr.findall(f"{{{a_ns}}}latin"):
                    has_latin = True
                    if latin_font:
                        latin.set("typeface", latin_font)
                if not has_latin and latin_font:
                    latin_elem = etree.SubElement(rpr, f"{{{a_ns}}}latin")
                    latin_elem.set("typeface", latin_font)

                for tag in (f"{{{a_ns}}}ea", f"{{{a_ns}}}cs"):
                    elem = rpr.find(tag)
                    if elem is not None:
                        if cjk_font:
                            elem.set("typeface", cjk_font)
                    else:
                        if cjk_font:
                            new_elem = etree.SubElement(rpr, tag)
                            new_elem.set("typeface", cjk_font)

    def _find_closest_level(self, sz_pt: float, font_levels: list[FontLevel]) -> str:
        """Find closest font level by size."""
        if not font_levels:
            return "body"
        best_role = "body"
        best_dist = float("inf")
        for level in font_levels:
            dist = abs(level.sz_pt - sz_pt)
            if dist < best_dist:
                best_dist = dist
                best_role = level.role
        return best_role

    def _transform_layout(self, root, plan: AdaptationPlan, analysis: ComponentAnalysis):
        """Transform layout: bounds, scaling, font size protection."""
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        target_left, target_top, target_w, target_h = plan.target_bounds
        target_left_emu = int(target_left * 914400)
        target_top_emu = int(target_top * 914400)
        target_w_emu = int(target_w * 914400)
        target_h_emu = int(target_h * 914400)

        grpSpPr = root.find(f"{{{p_ns}}}grpSpPr")
        if grpSpPr is None:
            return
        grp_xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
        if grp_xfrm is None:
            return

        chOff = grp_xfrm.find(f"{{{a_ns}}}chOff")
        chExt = grp_xfrm.find(f"{{{a_ns}}}chExt")
        if chOff is None or chExt is None:
            return

        orig_chOff_x = int(chOff.get("x", "0"))
        orig_chOff_y = int(chOff.get("y", "0"))
        orig_chExt_cx = int(chExt.get("cx", "0"))
        orig_chExt_cy = int(chExt.get("cy", "0"))

        if orig_chExt_cx <= 0 or orig_chExt_cy <= 0:
            return

        scale_x = target_w_emu / orig_chExt_cx
        scale_y = target_h_emu / orig_chExt_cy

        if plan.fit_strategy == "stretch":
            use_sx, use_sy = scale_x, scale_y
        elif plan.fit_strategy == "width":
            use_sx = scale_x
            use_sy = scale_x
        elif plan.fit_strategy == "height":
            use_sx = scale_y
            use_sy = scale_y
        else:
            uniform = min(scale_x, scale_y)
            use_sx = use_sy = uniform

        actual_w = int(orig_chExt_cx * use_sx)
        actual_h = int(orig_chExt_cy * use_sy)
        offset_x = (target_w_emu - actual_w) // 2
        offset_y = (target_h_emu - actual_h) // 2

        off_elem = grp_xfrm.find(f"{{{a_ns}}}off")
        ext_elem = grp_xfrm.find(f"{{{a_ns}}}ext")
        if off_elem is not None:
            off_elem.set("x", str(target_left_emu + offset_x))
            off_elem.set("y", str(target_top_emu + offset_y))
        if ext_elem is not None:
            ext_elem.set("cx", str(actual_w))
            ext_elem.set("cy", str(actual_h))

        new_chOff_x = int(orig_chOff_x * use_sx)
        new_chOff_y = int(orig_chOff_y * use_sy)
        chOff.set("x", str(new_chOff_x))
        chOff.set("y", str(new_chOff_y))
        chExt.set("cx", str(actual_w))
        chExt.set("cy", str(actual_h))

        self._transform_child_elements(root, orig_chOff_x, orig_chOff_y, use_sx, use_sy)

        # Adaptive font scaling
        self._scale_fonts_adaptive(root, use_sy, plan.min_font_pt)

        # bodyPr inset protection
        self._protect_body_pr_insets(root)

    def _transform_child_elements(self, grp_elem, chOff_x: int, chOff_y: int, scale_x: float, scale_y: float):
        """Transform child element coordinates recursively."""
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

                for ln in child.iter(f"{{{a_ns}}}ln"):
                    w = ln.get("w")
                    if w:
                        try:
                            new_w = int(int(w) * min(scale_x, scale_y))
                            if new_w > 0:
                                ln.set("w", str(new_w))
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

                        sub_chOff.set("x", str(int(sub_chOff_x * scale_x)))
                        sub_chOff.set("y", str(int(sub_chOff_y * scale_y)))
                        sub_chExt.set("cx", str(int(sub_chExt_cx * scale_x)))
                        sub_chExt.set("cy", str(int(sub_chExt_cy * scale_y)))
                    except (ValueError, TypeError):
                        pass

    def _scale_fonts_adaptive(self, root, scale_y: float, min_font_pt: float):
        """Scale font sizes with minimum protection."""
        a_ns = _NS["a"]
        min_sz = int(min_font_pt * 100)

        for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr", f"{{{a_ns}}}defRPr"):
            for rpr in root.iter(rpr_tag):
                sz = rpr.get("sz")
                if not sz:
                    continue
                try:
                    orig_sz = int(sz)
                    if scale_y < 1.0:
                        new_sz = max(min_sz, int(orig_sz * scale_y))
                    else:
                        new_sz = orig_sz
                    rpr.set("sz", str(new_sz))
                except (ValueError, TypeError):
                    pass

    def _protect_body_pr_insets(self, root):
        """Ensure bodyPr insets don't become zero after scaling."""
        a_ns = _NS["a"]
        for bodyPr in root.iter(f"{{{a_ns}}}bodyPr"):
            for attr in ("lIns", "tIns", "rIns", "bIns"):
                val = bodyPr.get(attr)
                if val is not None:
                    try:
                        v = int(val)
                        if v < _MIN_INSET_EMU:
                            bodyPr.set(attr, str(_MIN_INSET_EMU))
                    except (ValueError, TypeError):
                        pass

    # ── Phase 4: Validate ──

    def validate(self, result: dict, brand_spec: BrandSpec, plan: AdaptationPlan) -> list[str]:
        """Validate adaptation result and auto-fix issues."""
        issues: list[str] = []

        group_xml = result.get("group")
        if not group_xml:
            return issues

        if isinstance(group_xml, str):
            group_xml = group_xml.encode("utf-8")

        try:
            root = etree.fromstring(group_xml)
        except Exception:
            return issues

        a_ns = _NS["a"]
        p_ns = _NS["p"]

        # V1: Contrast check
        text_colors: set[str] = set()
        fill_colors: set[str] = set()

        for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
            val = srgb.get("val", "").upper()
            if not val:
                continue
            parent = srgb.getparent()
            if parent is None:
                continue
            gp = parent.getparent()
            if gp is not None:
                gp_tag = etree.QName(gp.tag).localname if isinstance(gp.tag, str) else ""
                p_tag = etree.QName(parent.tag).localname if isinstance(parent.tag, str) else ""
                if p_tag == "solidFill":
                    if gp_tag in ("rPr", "endParaRPr", "defRPr"):
                        text_colors.add(val)
                    else:
                        fill_colors.add(val)

        for tc in list(text_colors):
            for fc in fill_colors:
                ratio = self._contrast_ratio(tc, fc)
                if ratio < 3.0:
                    issues.append(f"Low contrast: text #{tc} on fill #{fc} (ratio={ratio:.1f})")
                    fg = plan.brand_colors.get("FOREGROUND") or plan.brand_colors.get("foreground", "0F172A")
                    if not plan.is_dark_brand:
                        self._replace_text_color(root, tc, fg.upper())

        # V2: Minimum font size
        for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr", f"{{{a_ns}}}defRPr"):
            for rpr in root.iter(rpr_tag):
                sz = rpr.get("sz")
                if sz:
                    try:
                        sz_val = int(sz)
                        min_sz = int(plan.min_font_pt * 100)
                        if sz_val < min_sz:
                            rpr.set("sz", str(min_sz))
                            issues.append(f"Font size {sz_val / 100}pt below minimum, raised to {plan.min_font_pt}pt")
                    except (ValueError, TypeError):
                        pass

        # V3: bodyPr inset check
        for bodyPr in root.iter(f"{{{a_ns}}}bodyPr"):
            for attr in ("lIns", "tIns", "rIns", "bIns"):
                val = bodyPr.get(attr)
                if val is not None:
                    try:
                        if int(val) < _MIN_INSET_EMU:
                            bodyPr.set(attr, str(_MIN_INSET_EMU))
                            issues.append(f"bodyPr {attr} below minimum, raised to {_MIN_INSET_EMU}")
                    except (ValueError, TypeError):
                        pass

        result["group"] = etree.tostring(root, xml_declaration=False, encoding="UTF-8")

        return issues

    def _replace_text_color(self, root, old_color: str, new_color: str):
        """Replace a specific text color in the XML."""
        a_ns = _NS["a"]
        for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
            val = srgb.get("val", "").upper()
            if val == old_color.upper():
                parent = srgb.getparent()
                if parent is not None:
                    gp = parent.getparent()
                    if gp is not None:
                        gp_tag = etree.QName(gp.tag).localname if isinstance(gp.tag, str) else ""
                        if gp_tag in ("rPr", "endParaRPr", "defRPr"):
                            srgb.set("val", new_color)

    # ── Utilities ──

    @staticmethod
    def _brightness(hex_color: str) -> float:
        h = hex_color.lstrip("#")
        if len(h) < 6:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2 if len(h) == 3 else h + "0" * (6 - len(h))
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return 0.299 * r + 0.587 * g + 0.114 * b

    @staticmethod
    def _hex_to_hsl(hex_val: str):
        h = hex_val.lstrip("#")
        if len(h) < 6:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2 if len(h) == 3 else h + "0" * (6 - len(h))
        r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
        hue, light, sat = colorsys.rgb_to_hls(r, g, b)
        return hue * 360, sat, light

    @staticmethod
    def _relative_luminance(hex_color: str) -> float:
        h = hex_color.lstrip("#")
        if len(h) < 6:
            return 0.5
        r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
        r = r / 12.92 if r <= 0.03928 else ((r + 0.055) / 1.055) ** 2.4
        g = g / 12.92 if g <= 0.03928 else ((g + 0.055) / 1.055) ** 2.4
        b = b / 12.92 if b <= 0.03928 else ((b + 0.055) / 1.055) ** 2.4
        return 0.2126 * r + 0.7152 * g + 0.0722 * b

    def _contrast_ratio(self, hex1: str, hex2: str) -> float:
        l1 = self._relative_luminance(hex1)
        l2 = self._relative_luminance(hex2)
        lighter = max(l1, l2)
        darker = min(l1, l2)
        return (lighter + 0.05) / (darker + 0.05)

    def _is_brand_dark(self, brand_spec: BrandSpec) -> bool:
        if brand_spec.dark_mode:
            return True
        bg = (brand_spec.colors or {}).get("background") or (brand_spec.colors or {}).get("BACKGROUND", "#FFFFFF")
        return self._brightness(bg.lstrip("#")) < 128
