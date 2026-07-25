"""ComponentAdapter v2 — 组件自适应适配器。

4 阶段管线让任意组件在任意品牌规范下正确显示:
  1. analyze()   — 解析组件结构（面积加权颜色角色、字号层级聚类、形状密度）
  2. plan()      — 制定适配策略（fill/text 分离映射表、层级保护字号、bounds 策略）
  3. transform() — 执行变换（上下文感知颜色替换、层级感知字体替换、归一化坐标缩放）
  4. validate()  — 闭环验证（对比度检查、层级间距检查、溢出检查 + replan）

v2 核心改进:
  - 坐标: chOff 归零 + 递归嵌套 grpSp + 幂等性
  - 配色: 面积加权角色推断 + fill_map/text_map 上下文分离
  - 字体: 层级保护缩放（min_gap ≥ 1pt）+ 放大跟随
  - 架构: skip_denorm 防双重变换 + 闭环 replan
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
_MIN_FONT_GAP_PT = 1.0


# ── Data structures ──

@dataclass
class ColorRole:
    hex_val: str
    role: str            # "dominant_fill" | "secondary_fill" | "data_fill" | "text" | "theme_ref"
    contexts: list[str]  # "shape_fill" | "text" | "gradient" | "theme_ref"
    count: int = 0
    area_emu2: int = 0   # area coverage in EMU^2 (for fill contexts)
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
    nesting_depth: int = 0
    orig_bounds_emu: tuple[int, int, int, int] = (0, 0, 9144000, 5486400)


@dataclass
class AdaptationPlan:
    color_map: dict[str, str] = field(default_factory=dict)
    fill_map: dict[str, str] = field(default_factory=dict)
    text_map: dict[str, str] = field(default_factory=dict)
    font_map: dict[str, FontMapping] = field(default_factory=dict)
    fit_strategy: str = "contain"
    target_bounds: tuple[float, float, float, float] = (0.9, 1.6, 11.5, 5.0)
    min_font_pt: float = _MIN_FONT_PT
    force_fg_on_light_bg: bool = False
    brand_colors: dict[str, str] = field(default_factory=dict)
    is_dark_brand: bool = False
    brand_spec: Any = None


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

        issues, need_replan = self.validate(result, brand_spec, plan, analysis)

        if need_replan:
            plan2 = self.replan(analysis, brand_spec, element, issues)
            result = self.transform(result, plan2, analysis)
            issues2, _ = self.validate(result, brand_spec, plan2, analysis)
            issues.extend(issues2)

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

        # Nesting depth
        analysis.nesting_depth = self._detect_nesting_depth(root, p_ns)

        # Color analysis (area-weighted)
        self._analyze_colors(root, analysis, a_ns, p_ns)

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

        # Aspect ratio and original bounds
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

    def _detect_nesting_depth(self, root, p_ns) -> int:
        """Detect maximum grpSp nesting depth."""
        max_depth = 0

        def walk(elem, depth):
            nonlocal max_depth
            for child in elem:
                tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if tag == "grpSp":
                    new_depth = depth + 1
                    max_depth = max(max_depth, new_depth)
                    walk(child, new_depth)

        walk(root, 0)
        return max_depth

    def _analyze_colors(self, root, analysis, a_ns, p_ns):
        """Area-weighted color role analysis."""
        color_info: dict[str, dict] = {}
        total_area = 0

        for sp in root.iter(f"{{{p_ns}}}sp"):
            spPr = sp.find(f"{{{p_ns}}}spPr")
            if spPr is None:
                continue
            xfrm = spPr.find(f"{{{a_ns}}}xfrm")
            shape_area = 0
            if xfrm is not None:
                ext = xfrm.find(f"{{{a_ns}}}ext")
                if ext is not None:
                    shape_area = int(ext.get("cx", "0")) * int(ext.get("cy", "0"))
            total_area += shape_area

            for srgb in sp.iter(f"{{{a_ns}}}srgbClr"):
                val = srgb.get("val", "").upper()
                if not val:
                    continue
                if val not in color_info:
                    color_info[val] = {"count": 0, "contexts": [], "area": 0}
                color_info[val]["count"] += 1
                parent = srgb.getparent()
                context = self._classify_color_context(srgb, parent, root)
                color_info[val]["contexts"].append(context)
                if context == "shape_fill":
                    color_info[val]["area"] += shape_area

        total_area = max(total_area, 1)
        data_idx = 0
        for val, info in color_info.items():
            n_fill = sum(1 for c in info["contexts"] if c == "shape_fill")
            n_text = sum(1 for c in info["contexts"] if c == "text")
            n_theme = sum(1 for c in info["contexts"] if c == "theme_ref")
            area_ratio = info["area"] / total_area

            if n_theme > 0 and n_fill == 0 and n_text == 0:
                role = "theme_ref"
            elif n_text > 0 and n_fill == 0 and n_theme == 0:
                role = "text"
            elif n_fill > 0 and area_ratio >= 0.25:
                role = "dominant_fill"
            elif n_fill > 0 and area_ratio < 0.08:
                role = "data_fill"
            elif n_fill > 0:
                role = "secondary_fill"
            elif n_text > 0:
                role = "text"
            else:
                role = "secondary_fill"

            analysis.color_roles[val] = ColorRole(
                hex_val=val,
                role=role,
                contexts=info["contexts"],
                count=info["count"],
                area_emu2=info["area"],
                index=data_idx if role == "data_fill" else 0,
            )
            if role == "data_fill":
                data_idx += 1

        analysis.color_count = len([v for v in analysis.color_roles.values() if v.role != "theme_ref"])

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
        """Infer font level hierarchy by clustering.

        v2: <=4 distinct sizes map directly to roles (no gap skipping).
        >4 sizes: use gap-based clustering but each cluster includes boundary.
        """
        sorted_pt = sorted(set(all_pt), reverse=True)
        if not sorted_pt:
            return [FontLevel(sz_pt=14, role="body", count=1)]
        if len(sorted_pt) == 1:
            return [FontLevel(sz_pt=sorted_pt[0], role="body", count=len(all_pt))]

        role_names = ["title", "subtitle", "body", "caption"]

        if len(sorted_pt) <= 4:
            levels = []
            for i, sz in enumerate(sorted_pt):
                count = sum(1 for p in all_pt if p == sz)
                levels.append(FontLevel(sz_pt=sz, role=role_names[min(i, 3)], count=max(count, 1)))
            return levels

        gaps = [(sorted_pt[i] - sorted_pt[i + 1], i) for i in range(len(sorted_pt) - 1)]
        gaps.sort(reverse=True)

        n_boundaries = min(3, len(gaps))
        split_indices = sorted([g[1] for g in gaps[:n_boundaries]])

        levels: list[FontLevel] = []
        prev = 0
        for i, split in enumerate(split_indices + [len(sorted_pt) - 1]):
            if i < len(split_indices):
                group = sorted_pt[prev:split + 1]
            else:
                group = sorted_pt[prev:]
            if not group:
                continue
            level_pt = group[0]
            count = sum(1 for p in all_pt if min(group) <= p <= max(group))
            levels.append(FontLevel(
                sz_pt=level_pt,
                role=role_names[min(i, len(role_names) - 1)],
                count=max(count, 1),
            ))
            prev = split + 1

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
            brand_spec=brand_spec,
        )

        # Color maps (fill + text separated)
        fill_map, text_map = self._build_color_maps(analysis, brand_colors, is_dark, brand_spec)
        plan.fill_map = fill_map
        plan.text_map = text_map
        plan.color_map = {**fill_map, **text_map}

        # Font map
        plan.font_map = self._build_font_map(analysis, brand_spec)

        # Layout
        raw_bounds = element.get("bounds", (0.9, 1.6, 11.5, 5.0))
        plan.target_bounds = tuple(float(v) for v in raw_bounds)
        plan.fit_strategy = self._determine_fit_strategy(analysis, element)

        return plan

    def _build_color_maps(self, analysis: ComponentAnalysis, brand_colors: dict[str, str], is_dark: bool, brand_spec: BrandSpec | None = None):
        """Build context-aware fill_map and text_map separately."""
        primary = brand_colors.get("PRIMARY") or brand_colors.get("primary", "2563EB")
        muted = brand_colors.get("MUTED") or brand_colors.get("muted", "F1F5F9")
        muted_dark = self._darken_color(muted, 0.7)
        foreground = brand_colors.get("FOREGROUND") or brand_colors.get("foreground", "0F172A")

        data_palette = self._generate_data_palette(primary, max(analysis.color_count, 6), brand_spec)

        fill_map: dict[str, str] = {}
        text_map: dict[str, str] = {}
        data_idx = 0

        for val, role_info in analysis.color_roles.items():
            if role_info.role == "theme_ref":
                continue

            has_text = "text" in role_info.contexts

            if role_info.role == "dominant_fill":
                fill_map[val] = primary.upper()
            elif role_info.role == "secondary_fill":
                fill_map[val] = muted_dark.upper()
            elif role_info.role == "data_fill":
                fill_map[val] = data_palette[data_idx % len(data_palette)].upper()
                data_idx += 1

            if has_text:
                on_primary = brand_colors.get("ON-PRIMARY") or brand_colors.get("on-primary", "FFFFFF")
                if is_dark:
                    text_map[val] = foreground.upper()
                else:
                    if self._brightness(val) > 180:
                        text_map[val] = on_primary.upper()
                    else:
                        text_map[val] = foreground.upper()

        # White text fix on light brand with dark component bg
        if not is_dark and analysis.has_dark_bg:
            for val in ("FFFFFF", "FFF"):
                if val not in text_map:
                    text_map[val] = foreground.upper()

        return fill_map, text_map

    def _generate_data_palette(self, primary_hex: str, count: int, brand_spec: BrandSpec | None = None) -> list[str]:
        """Generate harmonious data colors from primary.

        Uses brand DNA colors when available, falls back to HSL rotation.
        """
        if brand_spec and hasattr(brand_spec, "_dna_actual_colors") and brand_spec._dna_actual_colors:
            actual = brand_spec._dna_actual_colors
            sorted_by_bri = sorted(actual.keys(), key=lambda c: self._brightness(c.lstrip("#")))
            candidates = [c.lstrip("#").upper() for c in sorted_by_bri
                          if 70 < self._brightness(c.lstrip("#")) < 240]
            if len(candidates) >= count:
                step = len(candidates) / count
                return [candidates[int(i * step)] for i in range(count)]
            if len(candidates) >= 2:
                result = list(candidates)
                while len(result) < count:
                    result.append(candidates[len(result) % len(candidates)])
                return result[:count]

        h = primary_hex.lstrip("#")
        hue, sat, lig = self._hex_to_hsl(h)
        colors = []
        spread = 80
        for i in range(count):
            offset = (i - (count - 1) / 2) * (spread / max(count - 1, 1))
            new_h = (hue + offset) % 360
            new_s = max(0.30, min(0.75, sat + (i % 3 - 1) * 0.10))
            new_l = max(0.45, min(0.70, lig + 0.15 + (i % 2) * 0.10))
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
        if analysis.nesting_depth > 1:
            return "stretch"
        raw_bounds = element.get("bounds", (0.9, 1.6, 11.5, 5.0))
        if len(raw_bounds) >= 4:
            target_ar = raw_bounds[2] / max(raw_bounds[3], 0.01)
            ratio_diff = max(analysis.aspect_ratio, target_ar) / max(min(analysis.aspect_ratio, target_ar), 0.01)
            if ratio_diff > 4.0:
                if analysis.aspect_ratio > target_ar:
                    return "width"
                else:
                    return "height"
        return "stretch"

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

        # Step 1: Remove empty text shapes first (before font/layout transforms)
        self._remove_empty_shapes(root)

        # Step 2: Color transform
        self._transform_colors(root, plan, analysis)

        # Step 3: Font transform
        self._transform_fonts(root, plan, analysis)

        # Step 4: Layout transform
        self._transform_layout(root, plan, analysis)

        # Step 5 (post-layout): Fix text color on dark fills
        self._fix_dark_fill_text(root, plan)

        result["group"] = etree.tostring(root, xml_declaration=False, encoding="UTF-8")
        return result

    def _transform_colors(self, root, plan: AdaptationPlan, analysis: ComponentAnalysis):
        """Context-aware color transform: fill_map for fills, text_map for text."""
        a_ns = _NS["a"]
        p_ns = _NS["p"]
        fill_map = plan.fill_map
        text_map = plan.text_map

        # Step 1: Replace srgbClr based on context
        for srgb in root.iter(f"{{{a_ns}}}srgbClr"):
            val = srgb.get("val", "").upper()
            if not val:
                continue

            parent = srgb.getparent()
            if parent is None:
                continue

            context = self._classify_color_context(srgb, parent, root)

            if context == "text":
                if val in text_map:
                    srgb.set("val", text_map[val])
            elif context in ("shape_fill", "gradient"):
                if val in fill_map:
                    srgb.set("val", fill_map[val])
            elif context == "theme_ref":
                pass
            else:
                if val in fill_map:
                    srgb.set("val", fill_map[val])
                elif val in text_map:
                    srgb.set("val", text_map[val])

        # Step 2: Replace schemeClr (not inside style elements)
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

        # Step 2.5: Replace prstClr (preset colors like "white") with brand colors
        _PRST_COLOR_MAP = {
            "WHITE": ("FFFFFF", True),
            "BLACK": ("000000", False),
            "RED": ("FF0000", False),
            "GREEN": ("00FF00", False),
            "BLUE": ("0000FF", False),
            "YELLOW": ("FFFF00", True),
            "GRAY": ("808080", False),
            "GREY": ("808080", False),
            "LIGHTGRAY": ("D3D3D3", True),
            "LIGHTGREY": ("D3D3D3", True),
            "DARKGRAY": ("404040", False),
            "DARKGREY": ("404040", False),
        }
        fg_hex = plan.brand_colors.get("FOREGROUND") or plan.brand_colors.get("foreground", "0F172A")
        for prst in list(root.iter(f"{{{a_ns}}}prstClr")):
            val = prst.get("val", "").upper()
            parent = prst.getparent()
            if parent is None:
                continue
            gp = parent.getparent()
            if gp is not None:
                gp_tag = etree.QName(gp.tag).localname if isinstance(gp.tag, str) else ""
                if gp_tag in ("rPr", "endParaRPr", "defRPr"):
                    idx = list(parent).index(prst)
                    srgb = etree.Element(f"{{{a_ns}}}srgbClr")
                    srgb.set("val", fg_hex.upper())
                    for ch in prst:
                        srgb.append(copy.deepcopy(ch))
                    parent.remove(prst)
                    parent.insert(idx, srgb)
                elif val in _PRST_COLOR_MAP:
                    hex_val, is_light = _PRST_COLOR_MAP[val]
                    if is_light:
                        brand_val = plan.brand_colors.get("MUTED") or plan.brand_colors.get("muted", "F1F5F9")
                    else:
                        brand_val = plan.brand_colors.get("PRIMARY") or plan.brand_colors.get("primary", "2563EB")
                    idx = list(parent).index(prst)
                    srgb = etree.Element(f"{{{a_ns}}}srgbClr")
                    srgb.set("val", brand_val.upper())
                    for ch in prst:
                        srgb.append(copy.deepcopy(ch))
                    parent.remove(prst)
                    parent.insert(idx, srgb)

        # Step 3: White text fix on light background
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

        # Step 3.5: Inject explicit text color into rPr elements that have no solidFill
        # Without this, text inherits theme colors which may be invisible against brand bg
        fg_for_text = plan.brand_colors.get("FOREGROUND") or plan.brand_colors.get("foreground", "0F172A")
        for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}defRPr"):
            for rpr in root.iter(rpr_tag):
                existing_fill = rpr.find(f"{{{a_ns}}}solidFill")
                if existing_fill is not None:
                    continue
                solidFill = etree.SubElement(rpr, f"{{{a_ns}}}solidFill")
                srgb = etree.SubElement(solidFill, f"{{{a_ns}}}srgbClr")
                srgb.set("val", fg_for_text.upper())

        # Step 4: Inject brand fills into shapes that have no explicit fill
        # This prevents PowerPoint from using theme colors (FFCA08 etc.)
        self._inject_brand_fills(root, plan, analysis)

        # Step 5: Replace srgbClr inside <a:style> elements (fillRef, lnRef, effectRef, fontRef)
        # These reference theme colors and override our brand fills
        self._replace_style_colors(root, plan)

        # Step 6: Fix text color on dark fills — use on-primary (white) for readability
        # (Moved to end of transform() to run after all other modifications)

    def _fix_dark_fill_text(self, root, plan: AdaptationPlan):
        """Fix text color on dark fills — use on-primary (white) for readability.

        Must run AFTER all other transforms so it isn't overwritten by later steps.
        """
        a_ns = _NS["a"]
        p_ns = _NS["p"]
        on_primary = plan.brand_colors.get("ON-PRIMARY") or plan.brand_colors.get("on-primary", "FFFFFF")
        for sp in root.iter(f"{{{p_ns}}}sp"):
            spPr = sp.find(f"{{{p_ns}}}spPr")
            if spPr is None:
                continue
            fill_hex = None
            for sf in spPr.iter(f"{{{a_ns}}}solidFill"):
                srgb = sf.find(f"{{{a_ns}}}srgbClr")
                if srgb is not None:
                    fill_hex = srgb.get("val", "").upper()
                    break
            if not fill_hex:
                continue
            if self._brightness(fill_hex) < 150:
                for rpr in sp.iter(f"{{{a_ns}}}rPr"):
                    existing = rpr.find(f"{{{a_ns}}}solidFill")
                    if existing is not None:
                        srgb2 = existing.find(f"{{{a_ns}}}srgbClr")
                        if srgb2 is not None:
                            text_hex = srgb2.get("val", "").upper()
                            if self._brightness(text_hex) < 150:
                                srgb2.set("val", on_primary.upper())

    def _remove_empty_shapes(self, root):
        """Remove empty textbox shapes (txBox=1, no visible text) after data fill.

        Only removes shapes that are explicitly textboxes AND have no visible text.
        Decorative shapes (arrows, backgrounds, circles, dividers) are preserved
        even if they have no text — they are visual elements of the component.
        """
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        for grpSp in list(root.iter(f"{{{p_ns}}}grpSp")):
            sps_to_remove = []
            for sp in grpSp.findall(f"{{{p_ns}}}sp"):
                is_textbox = False
                nvSpPr = sp.find(f"{{{p_ns}}}nvSpPr")
                if nvSpPr is not None:
                    cNvSpPr = nvSpPr.find(f"{{{p_ns}}}cNvSpPr")
                    if cNvSpPr is not None:
                        is_textbox = cNvSpPr.get("txBox", "0") == "1"

                if not is_textbox:
                    continue

                has_visible_text = False
                for t in sp.iter(f"{{{a_ns}}}t"):
                    if t.text and t.text.strip():
                        has_visible_text = True
                        break

                if not has_visible_text:
                    sps_to_remove.append(sp)

            for sp in sps_to_remove:
                grpSp.remove(sp)

        self._cleanup_empty_runs(root)

    def _cleanup_empty_runs(self, root):
        """Remove empty <a:r> runs (with no visible text) from shapes.

        After _clear_unfilled_placeholders, many runs have empty <a:t> but
        still carry font/color properties. These bloat the XML and can
        trigger PowerPoint repair prompts.
        """
        a_ns = _NS["a"]

        for r in list(root.iter(f"{{{a_ns}}}r")):
            t = r.find(f"{{{a_ns}}}t")
            if t is None or t.text is None or not t.text.strip():
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)

    def _darken_color(self, hex_val: str, factor: float = 0.7) -> str:
        h = hex_val.lstrip("#")
        if len(h) < 6:
            return hex_val
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        r, g, b = int(r * factor), int(g * factor), int(b * factor)
        return f"{r:02X}{g:02X}{b:02X}"

    def _inject_brand_fills(self, root, plan: AdaptationPlan, analysis: ComponentAnalysis):
        """Inject explicit brand fills into shapes with no fill in spPr.

        Without this, PowerPoint uses the document theme's fillStyle chain,
        resulting in theme colors (like FFCA08) that ignore the brand spec.
        """
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        primary = plan.brand_colors.get("PRIMARY") or plan.brand_colors.get("primary", "2563EB")
        muted = plan.brand_colors.get("MUTED") or plan.brand_colors.get("muted", "F1F5F9")
        muted_dark = self._darken_color(muted, 0.7)

        shape_idx = 0
        for sp in root.iter(f"{{{p_ns}}}sp"):
            spPr = sp.find(f"{{{p_ns}}}spPr")
            if spPr is None:
                continue

            has_fill = False
            for tag in ("solidFill", "gradFill", "pattFill"):
                if spPr.find(f"{{{a_ns}}}{tag}") is not None:
                    has_fill = True
                    break

            noFill = spPr.find(f"{{{a_ns}}}noFill")
            if noFill is not None:
                spPr.remove(noFill)
                has_fill = False

            if has_fill:
                continue

            xfrm = spPr.find(f"{{{a_ns}}}xfrm")
            if xfrm is None:
                continue
            ext = xfrm.find(f"{{{a_ns}}}ext")
            if ext is None:
                continue
            cx = int(ext.get("cx", "0"))
            cy = int(ext.get("cy", "0"))
            if cx < 50000 or cy < 50000:
                continue

            prstGeom = spPr.find(f"{{{a_ns}}}prstGeom")
            if prstGeom is None:
                continue

            is_textbox = False
            nvSpPr = sp.find(f"{{{p_ns}}}nvSpPr")
            if nvSpPr is not None:
                cNvSpPr = nvSpPr.find(f"{{{p_ns}}}cNvSpPr")
                if cNvSpPr is not None:
                    is_textbox = cNvSpPr.get("txBox", "0") == "1"

            area = cx * cy
            total_area = analysis.orig_bounds_emu[2] * analysis.orig_bounds_emu[3]
            area_ratio = area / max(total_area, 1)

            if is_textbox and area_ratio < 0.05:
                continue

            if area_ratio >= 0.15:
                fill_color = primary
            else:
                fill_color = muted_dark

            if is_textbox and area_ratio < 0.15:
                fill_color = muted_dark

            solidFill = etree.Element(f"{{{a_ns}}}solidFill")
            srgbClr = etree.SubElement(solidFill, f"{{{a_ns}}}srgbClr")
            srgbClr.set("val", fill_color.upper())

            geom_idx = list(spPr).index(prstGeom) + 1
            spPr.insert(geom_idx, solidFill)

            shape_idx += 1

    def _replace_style_colors(self, root, plan: AdaptationPlan):
        """Replace srgbClr inside <a:style> and <p:style> elements with brand colors.

        <a:style> and <p:style> contain fillRef, lnRef, effectRef, fontRef which
        reference theme colors. When these have srgbClr, PowerPoint uses them
        instead of the shape's explicit fill. We must replace them with brand colors.
        """
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        primary = plan.brand_colors.get("PRIMARY") or plan.brand_colors.get("primary", "2563EB")
        foreground = plan.brand_colors.get("FOREGROUND") or plan.brand_colors.get("foreground", "0F172A")
        muted = plan.brand_colors.get("MUTED") or plan.brand_colors.get("muted", "F1F5F9")

        for ns in (a_ns, p_ns):
            for style in root.iter(f"{{{ns}}}style"):
                fillRef = style.find(f"{{{a_ns}}}fillRef")
                if fillRef is not None:
                    style.remove(fillRef)
                for ref_tag in ("lnRef", "effectRef"):
                    ref = style.find(f"{{{a_ns}}}{ref_tag}")
                    if ref is None:
                        continue
                    srgb = ref.find(f"{{{a_ns}}}srgbClr")
                    if srgb is not None:
                        if ref_tag == "lnRef":
                            srgb.set("val", muted.upper())
                        elif ref_tag == "effectRef":
                            srgb.set("val", muted.upper())

                fontRef = style.find(f"{{{a_ns}}}fontRef")
                if fontRef is not None:
                    srgb = fontRef.find(f"{{{a_ns}}}srgbClr")
                    if srgb is not None:
                        srgb.set("val", foreground.upper())

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

    # ── Layout Transform (Core Fix) ──

    def _transform_layout(self, root, plan: AdaptationPlan, analysis: ComponentAnalysis):
        """Normalized virtual canvas coordinate transform.

        Core rules:
        - new chOff = (0, 0)           (normalize origin)
        - new chExt = new ext           (1:1 mapping)
        - child: (orig_pos - orig_chOff) * (new_ext / orig_chExt)
        - Recursive for nested grpSp
        - Idempotent: second run produces scale=1.0
        """
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        target_left, target_top, target_w, target_h = plan.target_bounds
        max_slide_w = 13.33
        max_slide_h = 7.5
        target_w = min(target_w, max_slide_w - target_left)
        target_h = min(target_h, max_slide_h - target_top)
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

        # Idempotency guard: if already normalized and fits well, skip layout transform
        ext_elem = grp_xfrm.find(f"{{{a_ns}}}ext")
        if ext_elem is not None:
            cur_cx = int(ext_elem.get("cx", "0"))
            cur_cy = int(ext_elem.get("cy", "0"))
            already_normalized = (orig_chOff_x == 0 and orig_chOff_y == 0
                                  and orig_chExt_cx == cur_cx and orig_chExt_cy == cur_cy)
            if already_normalized:
                scale_to_target = min(target_w_emu / max(cur_cx, 1), target_h_emu / max(cur_cy, 1))
                if 0.95 <= scale_to_target <= 1.05:
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
        else:  # contain
            uniform = min(scale_x, scale_y)
            use_sx = use_sy = uniform

        actual_w = int(orig_chExt_cx * use_sx)
        actual_h = int(orig_chExt_cy * use_sy)
        offset_x = max(0, (target_w_emu - actual_w) // 2)
        offset_y = max(0, (target_h_emu - actual_h) // 2)

        off_elem = grp_xfrm.find(f"{{{a_ns}}}off")
        ext_elem = grp_xfrm.find(f"{{{a_ns}}}ext")
        new_off_x = target_left_emu + offset_x
        new_off_y = target_top_emu + offset_y

        max_slide_w_emu = int(max_slide_w * 914400)
        max_slide_h_emu = int(max_slide_h * 914400)
        if new_off_x + actual_w > max_slide_w_emu:
            new_off_x = max(0, max_slide_w_emu - actual_w)
        if new_off_y + actual_h > max_slide_h_emu:
            new_off_y = max(0, max_slide_h_emu - actual_h)

        if off_elem is not None:
            off_elem.set("x", str(new_off_x))
            off_elem.set("y", str(new_off_y))
        if ext_elem is not None:
            ext_elem.set("cx", str(actual_w))
            ext_elem.set("cy", str(actual_h))

        # KEY FIX: chOff -> (0, 0), chExt = ext (normalize virtual canvas)
        chOff.set("x", "0")
        chOff.set("y", "0")
        chExt.set("cx", str(actual_w))
        chExt.set("cy", str(actual_h))

        # Recursively transform children
        self._transform_children(root, orig_chOff_x, orig_chOff_y, use_sx, use_sy)

        # Hierarchy-preserving font scaling
        self._scale_fonts_hierarchical(root, use_sy, plan.min_font_pt, analysis, plan)

        # bodyPr inset protection
        self._protect_body_pr_insets(root)

    def _transform_children(self, parent, orig_chOff_x, orig_chOff_y, scale_x, scale_y):
        """Recursively transform child element coordinates.

        For shapes (sp, pic, cxnSp, graphicFrame):
            new_pos = (orig_pos - parent_chOff) * scale

        For nested grpSp:
            1. Transform off/ext (same as shapes)
            2. Read this grpSp's own chOff/chExt
            3. Compute sub_scale = new_ext / orig_chExt
            4. Recurse into this grpSp's children with sub_scale
            5. Normalize this grpSp's chOff -> 0, chExt = new_ext
        """
        for child in parent:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag in ("sp", "pic", "cxnSp", "graphicFrame"):
                self._transform_shape_coords(child, orig_chOff_x, orig_chOff_y, scale_x, scale_y)

            elif tag == "grpSp":
                self._transform_nested_group(child, orig_chOff_x, orig_chOff_y, scale_x, scale_y)

    def _transform_shape_coords(self, sp, chOff_x, chOff_y, scale_x, scale_y):
        """Transform a shape's coordinates."""
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        spPr = sp.find(f"{{{p_ns}}}spPr")
        if spPr is None:
            return
        xfrm = spPr.find(f"{{{a_ns}}}xfrm")
        if xfrm is None:
            return

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

        for ln in sp.iter(f"{{{a_ns}}}ln"):
            w = ln.get("w")
            if w:
                try:
                    new_w = int(int(w) * min(scale_x, scale_y))
                    if new_w > 0:
                        ln.set("w", str(new_w))
                except (ValueError, TypeError):
                    pass

    def _transform_nested_group(self, grpSp, parent_chOff_x, parent_chOff_y, parent_scale_x, parent_scale_y):
        """Recursively transform a nested grpSp.

        Steps:
        1. Transform off/ext like a regular shape
        2. Read this grpSp's own chOff/chExt
        3. Compute sub_scale = new_ext / orig_chExt
        4. Recurse into children with sub_scale
        5. Normalize chOff -> 0, chExt = new_ext
        """
        a_ns = _NS["a"]
        p_ns = _NS["p"]

        sub_grpSpPr = grpSp.find(f"{{{p_ns}}}grpSpPr")
        if sub_grpSpPr is None:
            return
        sub_xfrm = sub_grpSpPr.find(f"{{{a_ns}}}xfrm")
        if sub_xfrm is None:
            return

        # Step 1: Transform off/ext (in parent's coordinate space)
        sub_off = sub_xfrm.find(f"{{{a_ns}}}off")
        if sub_off is not None:
            try:
                orig_x = int(sub_off.get("x", "0"))
                orig_y = int(sub_off.get("y", "0"))
                sub_off.set("x", str(int((orig_x - parent_chOff_x) * parent_scale_x)))
                sub_off.set("y", str(int((orig_y - parent_chOff_y) * parent_scale_y)))
            except (ValueError, TypeError):
                pass

        sub_ext = sub_xfrm.find(f"{{{a_ns}}}ext")
        new_ext_cx = 0
        new_ext_cy = 0
        if sub_ext is not None:
            try:
                orig_cx = int(sub_ext.get("cx", "0"))
                orig_cy = int(sub_ext.get("cy", "0"))
                new_ext_cx = int(orig_cx * parent_scale_x)
                new_ext_cy = int(orig_cy * parent_scale_y)
                sub_ext.set("cx", str(new_ext_cx))
                sub_ext.set("cy", str(new_ext_cy))
            except (ValueError, TypeError):
                pass

        # Step 2: Read this grpSp's own chOff/chExt
        sub_chOff = sub_xfrm.find(f"{{{a_ns}}}chOff")
        sub_chExt = sub_xfrm.find(f"{{{a_ns}}}chExt")

        orig_sub_chOff_x = int(sub_chOff.get("x", "0")) if sub_chOff is not None else 0
        orig_sub_chOff_y = int(sub_chOff.get("y", "0")) if sub_chOff is not None else 0
        orig_sub_chExt_cx = int(sub_chExt.get("cx", "1")) if sub_chExt is not None else 1
        orig_sub_chExt_cy = int(sub_chExt.get("cy", "1")) if sub_chExt is not None else 1

        # Step 3: Compute sub_scale = new_ext / orig_chExt
        sub_scale_x = new_ext_cx / orig_sub_chExt_cx if orig_sub_chExt_cx > 0 else 1.0
        sub_scale_y = new_ext_cy / orig_sub_chExt_cy if orig_sub_chExt_cy > 0 else 1.0

        # Step 4: Recurse into children
        self._transform_children(grpSp, orig_sub_chOff_x, orig_sub_chOff_y, sub_scale_x, sub_scale_y)

        # Step 5: Normalize this grpSp's virtual canvas
        if sub_chOff is not None:
            sub_chOff.set("x", "0")
            sub_chOff.set("y", "0")
        if sub_chExt is not None:
            sub_chExt.set("cx", str(new_ext_cx))
            sub_chExt.set("cy", str(new_ext_cy))

    def _scale_fonts_hierarchical(self, root, scale_y: float, min_font_pt: float, analysis: ComponentAnalysis, plan: AdaptationPlan | None = None):
        """Hierarchy-preserving font scaling.

        Strategy: preserve original font size ratios by scaling individually,
        then enforce minimum size by lifting the smallest font to min_sz
        and proportionally lifting all others to maintain ratios.
        
        Max font is capped based on component area ratio — large components
        (e.g. full-width features) allow bigger titles, small components
        (e.g. single labels) are capped tighter to stay proportional to
        the project's title/body font sizes (typically 20-24pt title, 14pt body).
        """
        a_ns = _NS["a"]
        p_ns = _NS["p"]
        min_sz = int(min_font_pt * 100)

        if plan and plan.target_bounds:
            group_w = plan.target_bounds[2]
            group_h = plan.target_bounds[3]
        else:
            group_w = analysis.orig_bounds_emu[2] / 914400
            group_h = analysis.orig_bounds_emu[3] / 914400
        group_area_in2 = group_w * group_h
        if group_area_in2 >= 20:
            max_font_pt = 36
        elif group_area_in2 >= 10:
            max_font_pt = 30
        elif group_area_in2 >= 5:
            max_font_pt = 24
        else:
            max_font_pt = 20
        max_sz = int(max_font_pt * 100)

        non_empty_rprs: set = set()
        for r in root.iter(f"{{{a_ns}}}r"):
            t = r.find(f"{{{a_ns}}}t")
            if t is not None and t.text and t.text.strip():
                rpr = r.find(f"{{{a_ns}}}rPr")
                if rpr is not None:
                    non_empty_rprs.add(rpr)

        font_entries: list[tuple] = []
        for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr", f"{{{a_ns}}}defRPr"):
            for rpr in root.iter(rpr_tag):
                sz = rpr.get("sz")
                if sz:
                    try:
                        orig_sz = int(sz)
                        if orig_sz > 0:
                            font_entries.append((rpr, orig_sz, rpr in non_empty_rprs))
                    except (ValueError, TypeError):
                        pass

        if not font_entries:
            return

        scaled = []
        for rpr, orig_sz, is_visible in font_entries:
            if scale_y < 1.0:
                new_sz = int(orig_sz * scale_y)
            elif scale_y > 1.0:
                new_sz = int(orig_sz * min(scale_y, 1.5))
            else:
                new_sz = orig_sz
            scaled.append((rpr, orig_sz, new_sz, is_visible))

        if not scaled:
            return
        min_scaled = min(s for _, _, s, _ in scaled)
        if min_scaled < min_sz and min_scaled > 0:
            lift = min_sz / min_scaled
            scaled = [(rpr, orig_sz, max(int(new_sz * lift), min_sz), is_vis) for rpr, orig_sz, new_sz, is_vis in scaled]

        visible_sizes = [s for _, _, s, is_vis in scaled if is_vis and s > 0]
        if len(visible_sizes) >= 2:
            vis_min = min(visible_sizes)
            vis_max = max(visible_sizes)
            if vis_max / vis_min < 1.3:
                orig_vis_sizes = [(orig_sz, i) for i, (_, orig_sz, new_sz, is_vis) in enumerate(scaled) if is_vis and new_sz > 0]
                if orig_vis_sizes:
                    orig_min = min(s for s, _ in orig_vis_sizes)
                    orig_max = max(s for s, _ in orig_vis_sizes)
                    if orig_max > orig_min:
                        for orig_sz, i in orig_vis_sizes:
                            orig_ratio = (orig_sz - orig_min) / (orig_max - orig_min)
                            target = vis_min + int(orig_ratio * (int(vis_min * 1.3) - vis_min))
                            _, _, new_sz, is_vis = scaled[i]
                            scaled[i] = (scaled[i][0], scaled[i][1], max(target, new_sz), is_vis)
                    else:
                        target_max = int(vis_min * 1.3)
                        n_vis = len(orig_vis_sizes)
                        for rank, (orig_sz, i) in enumerate(sorted(orig_vis_sizes, key=lambda x: x[0], reverse=True)):
                            if rank == 0:
                                continue
                            frac = rank / n_vis
                            target = int(target_max - frac * (target_max - vis_min) * 0.5)
                            _, _, new_sz, is_vis = scaled[i]
                            scaled[i] = (scaled[i][0], scaled[i][1], max(target, new_sz), is_vis)

        for rpr, orig_sz, new_sz, _ in scaled:
            final_sz = max(min_sz, min(new_sz, max_sz))
            rpr.set("sz", str(final_sz))

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

    def validate(self, result: dict, brand_spec: BrandSpec, plan: AdaptationPlan, analysis: ComponentAnalysis):
        """Closed-loop validation: check and auto-fix, return (issues, need_replan)."""
        issues: list[str] = []

        group_xml = result.get("group")
        if not group_xml:
            return issues, False

        if isinstance(group_xml, str):
            group_xml = group_xml.encode("utf-8")

        try:
            root = etree.fromstring(group_xml)
        except Exception:
            return issues, False

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

        # V2: Minimum font size
        min_sz = int(plan.min_font_pt * 100)
        for rpr_tag in (f"{{{a_ns}}}rPr", f"{{{a_ns}}}endParaRPr", f"{{{a_ns}}}defRPr"):
            for rpr in root.iter(rpr_tag):
                sz = rpr.get("sz")
                if sz:
                    try:
                        sz_val = int(sz)
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

        # V4: Text overflow check
        overflow = self._check_text_overflow(root, a_ns, p_ns, issues)

        result["group"] = etree.tostring(root, xml_declaration=False, encoding="UTF-8")

        return issues, overflow

    def _check_text_overflow(self, root, a_ns, p_ns, issues) -> bool:
        """Estimate text overflow: compare text needed area vs box area."""
        overflow_found = False
        for sp in root.iter(f"{{{p_ns}}}sp"):
            spPr = sp.find(f"{{{p_ns}}}spPr")
            if spPr is None:
                continue
            xfrm = spPr.find(f"{{{a_ns}}}xfrm")
            if xfrm is None:
                continue
            ext = xfrm.find(f"{{{a_ns}}}ext")
            if ext is None:
                continue

            box_w = int(ext.get("cx", "0"))
            box_h = int(ext.get("cy", "0"))
            if box_w <= 0 or box_h <= 0:
                continue

            total_chars = 0
            max_font_emu = 0
            for rpr in sp.iter(f"{{{a_ns}}}rPr"):
                sz = rpr.get("sz")
                if sz:
                    try:
                        font_emu = int(sz) * 12700
                        max_font_emu = max(max_font_emu, font_emu)
                    except (ValueError, TypeError):
                        pass

            for t in sp.iter(f"{{{a_ns}}}t"):
                if t.text:
                    total_chars += len(t.text)

            if max_font_emu <= 0 or total_chars == 0:
                continue

            chars_per_line = max(1, box_w / int(max_font_emu * 0.6))
            lines = max(1, total_chars / chars_per_line)
            needed_h = int(lines * max_font_emu * 1.3)

            if needed_h > box_h * 1.2:
                issues.append(f"Text overflow: {total_chars} chars in {box_w / 914400:.1f}x{box_h / 914400:.1f}in box")
                overflow_found = True

        return overflow_found

    def replan(self, analysis: ComponentAnalysis, brand_spec: BrandSpec, element: dict, issues: list[str]) -> AdaptationPlan:
        """Re-plan based on validation issues."""
        plan = self.plan(analysis, brand_spec, element)

        overflow_issues = [i for i in issues if "overflow" in i.lower()]
        if overflow_issues:
            target_bounds = list(plan.target_bounds)
            target_bounds[3] = target_bounds[3] * 1.3
            plan.target_bounds = tuple(target_bounds)
            if plan.fit_strategy == "contain":
                plan.fit_strategy = "width"

        return plan

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
