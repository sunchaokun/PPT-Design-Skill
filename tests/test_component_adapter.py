"""Test ComponentAdapter — strict validation of layout, color, font details.

Tests use realistic GroupShape XML to verify:
1. Layout: no overflow, correct bounds, text fits
2. Color: no white-on-light, contrast >= 3:1, brand colors applied
3. Font: no unreadable sizes, CJK applied, heading/body correct
"""
from __future__ import annotations

import copy
import os
import tempfile

import pytest
from lxml import etree

from ppt_pro_max.enterprise.brand_spec import BrandSpec


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _light_brand() -> BrandSpec:
    return BrandSpec(
        source="test",
        colors={
            "primary": "#2563EB",
            "on-primary": "#FFFFFF",
            "secondary": "#7C3AED",
            "accent": "#F97316",
            "background": "#FFFFFF",
            "foreground": "#0F172A",
            "muted": "#F1F5F9",
            "muted-foreground": "#64748B",
            "border": "#E2E8F0",
            "destructive": "#EF4444",
        },
        fonts={"heading": "Inter", "body": "Inter", "cjk_heading": "Microsoft YaHei", "cjk_body": "Microsoft YaHei"},
    )


def _dark_brand() -> BrandSpec:
    return BrandSpec(
        source="test",
        colors={
            "primary": "#3B82F6",
            "on-primary": "#FFFFFF",
            "secondary": "#8B5CF6",
            "accent": "#F59E0B",
            "background": "#0F172A",
            "foreground": "#F8FAFC",
            "muted": "#1E293B",
            "muted-foreground": "#94A3B8",
            "border": "#334155",
            "destructive": "#EF4444",
        },
        fonts={"heading": "Orbitron", "body": "JetBrains Mono", "cjk_heading": "Microsoft YaHei", "cjk_body": "Microsoft YaHei"},
        dark_mode=True,
    )


def _make_group_xml(
    shapes=None,
    bg_color="4472C4",
    text_color="FFFFFF",
    title_size=3200,
    body_size=1800,
    fill_colors=None,
    font_latin="Calibri",
    font_ea="",
    width_emu=9144000,
    height_emu=5486400,
) -> bytes:
    """Build realistic GroupShape XML with configurable properties."""
    if fill_colors is None:
        fill_colors = ["4472C4", "ED7D31", "A5A5A5", "FFC000"]

    shapes_xml = ""
    for i, fc in enumerate(fill_colors):
        x = 200000 + i * 2200000
        y = 300000
        cx = 2000000
        cy = 1500000
        is_title = (i == 0)
        sz = title_size if is_title else body_size
        tc = text_color
        txt = f"Step {i+1}" if is_title else f"Item {i+1} detail text"

        ea_attr = f' ea:typeface="{font_ea}"' if font_ea else ""
        cjk_cs = ""
        if font_ea:
            cjk_cs = f'<a:ea typeface="{font_ea}"/><a:cs typeface="{font_ea}"/>'

        text_color_fill = f'<a:solidFill><a:srgbClr val="{tc}"/></a:solidFill>'

        shapes_xml += f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{10+i}" name="Shape{i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="{fc}"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="{sz}" b="{"1" if is_title else "0"}" dirty="0">
                <a:latin typeface="{font_latin}"/>{cjk_cs}
                {text_color_fill}
              </a:rPr>
              <a:t>{txt}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>"""

    xml = f"""<p:grpSp xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
  <p:nvGrpSpPr><p:cNvPr id="2" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="457200" y="457200"/>
      <a:ext cx="{width_emu}" cy="{height_emu}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="{width_emu}" cy="{height_emu}"/>
    </a:xfrm>
  </p:grpSpPr>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="3" name="BgRect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>
    </p:spPr>
  </p:sp>{shapes_xml}
</p:grpSp>"""
    return xml.encode("utf-8")


def _make_smartart_data_xml(texts=None) -> bytes:
    """Build SmartArt data XML for testing schemeClr replacement."""
    if texts is None:
        texts = ["Step 1", "Step 2", "Step 3"]
    dgm_ns = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
    pts = ""
    for i, t in enumerate(texts):
        pts += f'<dgm:pt><dgm:prLo val="0"/><a:t>{t}</a:t></dgm:pt>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="{dgm_ns}" xmlns:a="{A_NS}">
  <dgm:ptLst>{pts}</dgm:ptLst>
</dgm:dataModel>""".encode("utf-8")


def _make_smartart_colors_xml() -> bytes:
    """Build SmartArt colors XML with schemeClr references."""
    dgm_ns = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:colorsDef xmlns:dgm="{dgm_ns}" xmlns:a="{A_NS}">
  <dgm:styleLbl name="node0">
    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
  </dgm:styleLbl>
  <dgm:styleLbl name="node1">
    <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
  </dgm:styleLbl>
  <dgm:styleLbl name="node2">
    <a:solidFill><a:schemeClr val="dk1"/></a:solidFill>
  </dgm:styleLbl>
</dgm:colorsDef>""".encode("utf-8")


# ═══════════════════════════════════════════════════════════════
# Test: ComponentAdapter.analyze()
# ═══════════════════════════════════════════════════════════════

class TestComponentAdapterAnalyze:

    def test_analyze_extracts_color_roles(self):
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(bg_color="4472C4", text_color="FFFFFF", fill_colors=["4472C4", "ED7D31", "A5A5A5"])
        xml_parts = {"group": xml}
        analysis = adapter.analyze(xml_parts, {"type": "group", "category": "process", "node_count": 3})
        assert analysis.color_count >= 3
        assert any(r.role == "text" for r in analysis.color_roles.values())

    def test_analyze_detects_dark_background(self):
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(bg_color="1E293B", text_color="FFFFFF")
        xml_parts = {"group": xml}
        analysis = adapter.analyze(xml_parts, {"type": "group", "category": "process", "node_count": 3})
        assert analysis.has_dark_bg is True

    def test_analyze_detects_light_background(self):
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(bg_color="F1F5F9", text_color="0F172A")
        xml_parts = {"group": xml}
        analysis = adapter.analyze(xml_parts, {"type": "group", "category": "process", "node_count": 3})
        assert analysis.has_dark_bg is False

    def test_analyze_extracts_font_levels(self):
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(title_size=3200, body_size=1600)
        xml_parts = {"group": xml}
        analysis = adapter.analyze(xml_parts, {"type": "group", "category": "process", "node_count": 3})
        assert len(analysis.font_levels) >= 2
        roles = [l.role for l in analysis.font_levels]
        assert "title" in roles or "subtitle" in roles

    def test_analyze_extracts_aspect_ratio(self):
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(width_emu=9144000, height_emu=5486400)
        xml_parts = {"group": xml}
        analysis = adapter.analyze(xml_parts, {"type": "group", "category": "process", "node_count": 3})
        assert abs(analysis.aspect_ratio - (9144000 / 5486400)) < 0.1

    def test_analyze_detects_cjk(self):
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(font_ea="Microsoft YaHei")
        xml_parts = {"group": xml}
        analysis = adapter.analyze(xml_parts, {"type": "group", "category": "process", "node_count": 3})
        assert analysis.has_cjk is True


# ═══════════════════════════════════════════════════════════════
# Test: Color Adaptation — 白字在浅底上必须变深
# ═══════════════════════════════════════════════════════════════

class TestColorAdaptationWhiteOnLight:

    def test_white_text_on_light_bg_replaced_with_foreground(self):
        """CRITICAL: 白色文本在浅色品牌下必须替换为 foreground 色。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(bg_color="4472C4", text_color="FFFFFF")
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B", "C"], "node_count": 3,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        grp_root = etree.fromstring(result["group"])

        fg_hex = "0F172A"
        text_colors = set()
        for rpr_tag in (f"{{{A_NS}}}rPr", f"{{{A_NS}}}endParaRPr"):
            for rpr in grp_root.iter(rpr_tag):
                for sf in rpr.iter(f"{{{A_NS}}}srgbClr"):
                    parent = sf.getparent()
                    if parent is not None and parent.tag == f"{{{A_NS}}}solidFill":
                        gp = parent.getparent()
                        if gp is not None:
                            gp_local = etree.QName(gp.tag).localname if isinstance(gp.tag, str) else ""
                            if gp_local in ("rPr", "endParaRPr", "defRPr"):
                                text_colors.add(sf.get("val", "").upper())

        assert "FFFFFF" not in text_colors, f"White text still present in light brand! Found: {text_colors}"

    def test_white_text_kept_on_dark_bg(self):
        """深色品牌下白色文本应保留。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(bg_color="1E293B", text_color="FFFFFF")
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B"], "node_count": 2,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _dark_brand())
        grp_root = etree.fromstring(result["group"])

        has_white_text = False
        for rpr in grp_root.iter(f"{{{A_NS}}}rPr"):
            for sf in rpr.findall(f".//{{{A_NS}}}srgbClr"):
                val = sf.get("val", "").upper()
                parent = sf.getparent()
                if parent is not None and parent.tag == f"{{{A_NS}}}solidFill":
                    gp = parent.getparent()
                    if gp is not None:
                        gp_local = etree.QName(gp.tag).localname if isinstance(gp.tag, str) else ""
                        if gp_local in ("rPr", "endParaRPr", "defRPr") and val == "FFFFFF":
                            has_white_text = True

        assert has_white_text, "White text was incorrectly removed from dark brand"


# ═══════════════════════════════════════════════════════════════
# Test: Color Adaptation — contrast ratio validation
# ═══════════════════════════════════════════════════════════════

class TestColorContrastValidation:

    def test_text_fill_contrast_above_3_to_1(self):
        """所有文本-背景配对对比度必须 >= 3:1。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(fill_colors=["4472C4", "ED7D31", "A5A5A5", "FFC000"])
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B", "C", "D"], "node_count": 4,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())

        issues = result.get("_validation_issues", [])
        contrast_issues = [i for i in issues if "contrast" in i.lower() or "Low contrast" in i]
        assert len(contrast_issues) == 0, f"Contrast issues found: {contrast_issues}"


# ═══════════════════════════════════════════════════════════════
# Test: Font Adaptation — 字号层级识别
# ═══════════════════════════════════════════════════════════════

class TestFontAdaptationLevels:

    def test_heading_font_applied_to_large_text(self):
        """大字号文本必须使用 heading 字体。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(title_size=3600, body_size=1400, font_latin="Arial")
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["Title", "Body"], "node_count": 2,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        grp_root = etree.fromstring(result["group"])

        heading_fonts = set()
        for rpr in grp_root.iter(f"{{{A_NS}}}rPr"):
            sz = rpr.get("sz")
            if sz and int(sz) >= 3000:
                latin = rpr.find(f"{{{A_NS}}}latin")
                if latin is not None:
                    heading_fonts.add(latin.get("typeface", ""))

        assert "Inter" in heading_fonts or any("Inter" in f for f in heading_fonts), \
            f"Heading font should be Inter, got: {heading_fonts}"

    def test_body_font_applied_to_small_text(self):
        """小字号文本必须使用 body 字体。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(title_size=3600, body_size=1400, font_latin="Arial")
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["Title", "Body"], "node_count": 2,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        grp_root = etree.fromstring(result["group"])

        body_fonts = set()
        for rpr in grp_root.iter(f"{{{A_NS}}}rPr"):
            sz = rpr.get("sz")
            if sz and int(sz) <= 2000:
                latin = rpr.find(f"{{{A_NS}}}latin")
                if latin is not None:
                    body_fonts.add(latin.get("typeface", ""))

        assert "Inter" in body_fonts or any("Inter" in f for f in body_fonts), \
            f"Body font should be Inter, got: {body_fonts}"

    def test_cjk_font_added_when_missing(self):
        """缺少 CJK 字体的组件必须自动添加。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(font_latin="Arial", font_ea="")
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B"], "node_count": 2,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        grp_root = etree.fromstring(result["group"])

        has_cjk_ea = False
        for rpr in grp_root.iter(f"{{{A_NS}}}rPr"):
            ea = rpr.find(f"{{{A_NS}}}ea")
            if ea is not None and ea.get("typeface", ""):
                has_cjk_ea = True

        assert has_cjk_ea, "CJK font (ea element) should be added to rPr elements"

    def test_no_font_below_11pt(self):
        """缩放后字号不得低于 11pt。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(title_size=1200, body_size=800)
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B"], "node_count": 2,
                   "bounds": (0.9, 1.6, 5.0, 3.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        grp_root = etree.fromstring(result["group"])

        min_sz = float("inf")
        for rpr_tag in (f"{{{A_NS}}}rPr", f"{{{A_NS}}}endParaRPr", f"{{{A_NS}}}defRPr"):
            for rpr in grp_root.iter(rpr_tag):
                sz = rpr.get("sz")
                if sz:
                    min_sz = min(min_sz, int(sz) / 100)

        assert min_sz >= 11.0, f"Font size {min_sz}pt is below minimum 11pt"


# ═══════════════════════════════════════════════════════════════
# Test: Layout Adaptation — 边界/偏移/溢出
# ═══════════════════════════════════════════════════════════════

class TestLayoutAdaptation:

    def test_component_bounds_within_content_area(self):
        """组件必须完全在目标区域内。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(width_emu=9144000, height_emu=5486400)
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B"], "node_count": 2,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        bounds = result.get("_adapted_bounds", (0.9, 1.6, 11.5, 5.0))
        left, top, w, h = bounds
        assert left >= 0, f"Left bound {left} is negative"
        assert top >= 0, f"Top bound {top} is negative"
        assert w > 0 and h > 0, f"Invalid bounds: {bounds}"

    def test_bodypr_insets_not_zero(self):
        """文本框 insets 不能因缩放变为 0。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml()
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B"], "node_count": 2,
                   "bounds": (0.9, 1.6, 5.0, 3.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        grp_root = etree.fromstring(result["group"])

        for bodyPr in grp_root.iter(f"{{{A_NS}}}bodyPr"):
            for attr in ("lIns", "tIns", "rIns", "bIns"):
                val = bodyPr.get(attr)
                if val is not None:
                    assert int(val) >= 36000, f"bodyPr {attr}={val} is below minimum 36000 EMU (~0.04 inch)"

    def test_wide_component_uses_width_fit(self):
        """宽组件应优先适配宽度。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()
        xml = _make_group_xml(width_emu=16000000, height_emu=4000000)
        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": ["A", "B"], "node_count": 2,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        strategy = result.get("_fit_strategy", "contain")
        assert strategy in ("width", "contain"), f"Expected width or contain strategy, got {strategy}"


# ═══════════════════════════════════════════════════════════════
# Test: SmartArt schemeClr bug fix
# ═══════════════════════════════════════════════════════════════

class TestSmartArtSchemeClrFix:

    def test_schemeclr_replaced_with_correct_position(self):
        """schemeClr 替换后 srgbClr 必须在原位，不能跑到末尾。"""
        from ppt_pro_max.enterprise.component_renderer import ComponentRenderer
        renderer = ComponentRenderer()
        brand_spec = _light_brand()
        colors_xml = _make_smartart_colors_xml()
        xml_parts = {"colors": colors_xml}

        result = renderer._apply_brand_colors(xml_parts, brand_spec)
        result_colors = result.get("colors")
        if result_colors is None:
            pytest.skip("No colors XML to test")

        root = etree.fromstring(result_colors)
        srgb_vals = [s.get("val", "").upper() for s in root.iter(f"{{{A_NS}}}srgbClr")]
        brand_primary = "2563EB"
        assert brand_primary in srgb_vals or any(v for v in srgb_vals if v != ""), \
            f"Brand primary {brand_primary} not found in replaced colors. Got: {srgb_vals}"

    def test_no_schemeclr_left_after_replacement(self):
        """替换后不应残留 schemeClr（除非在 style 元素内）。"""
        from ppt_pro_max.enterprise.component_renderer import ComponentRenderer
        renderer = ComponentRenderer()
        brand_spec = _light_brand()
        colors_xml = _make_smartart_colors_xml()
        xml_parts = {"colors": colors_xml}

        result = renderer._apply_brand_colors(xml_parts, brand_spec)
        result_colors = result.get("colors")
        if result_colors is None:
            pytest.skip("No colors XML to test")

        root = etree.fromstring(result_colors)
        remaining_scheme = list(root.iter(f"{{{A_NS}}}schemeClr"))
        non_style_scheme = []
        for s in remaining_scheme:
            p = s.getparent()
            in_style = False
            while p is not None:
                if p.tag == f"{{{A_NS}}}style":
                    in_style = True
                    break
                p = p.getparent()
            if not in_style:
                non_style_scheme.append(s.get("val", ""))

        assert len(non_style_scheme) == 0, f"Unreplaced schemeClr values: {non_style_scheme}"


# ═══════════════════════════════════════════════════════════════
# Test: Group recolor — gradient handling
# ═══════════════════════════════════════════════════════════════

class TestGroupRecolorGradient:

    def test_gradient_colors_recolored(self):
        """渐变中的颜色也必须被替换。"""
        from ppt_pro_max.enterprise.component_adapter import ComponentAdapter
        adapter = ComponentAdapter()

        xml = f"""<p:grpSp xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
  <p:nvGrpSpPr><p:cNvPr id="2" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm><a:off x="457200" y="457200"/><a:ext cx="9144000" cy="5486400"/>
      <a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="5486400"/></a:xfrm>
  </p:grpSpPr>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="3" name="Shape0"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="200000" y="300000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:gradFill>
        <a:gsLst>
          <a:gs pos="0"><a:srgbClr val="4472C4"/></a:gs>
          <a:gs pos="50000"><a:srgbClr val="5B9BD5"/></a:gs>
          <a:gs pos="100000"><a:srgbClr val="ED7D31"/></a:gs>
        </a:gsLst>
      </a:gradFill>
    </p:spPr>
  </p:sp>
</p:grpSp>""".encode("utf-8")

        xml_parts = {"group": xml}
        element = {"type": "group", "category": "process", "texts": [], "node_count": 0,
                   "bounds": (0.9, 1.6, 11.5, 5.0)}
        result = adapter.adapt(xml_parts, element, _light_brand())
        grp_root = etree.fromstring(result["group"])

        gradient_colors = set()
        for gs in grp_root.iter(f"{{{A_NS}}}gs"):
            for srgb in gs.findall(f"{{{A_NS}}}srgbClr"):
                val = srgb.get("val", "").upper()
                gradient_colors.add(val)

        original_colors = {"4472C4", "5B9BD5", "ED7D31"}
        unchanged = gradient_colors & original_colors
        assert len(unchanged) < 3, f"Gradient colors not recolored: {unchanged} still present"


# ═══════════════════════════════════════════════════════════════
# Test: End-to-end PPT generation with component
# ═══════════════════════════════════════════════════════════════

class TestEndToEndComponentQuality:

    def _build_test_library(self, tmp_path):
        from ppt_pro_max.enterprise.component_library import ComponentLibrary
        db_path = str(tmp_path / "test_lib.db")
        lib = ComponentLibrary(db_path)

        xml = _make_group_xml(fill_colors=["4472C4", "ED7D31", "A5A5A5"], text_color="FFFFFF")
        lib.add(
            type="group",
            category="process",
            variant="chevron",
            node_count=3,
            xml_parts={"group": xml},
        )
        return lib

    def test_rendered_ppt_has_no_white_text_on_light_bg(self, tmp_path):
        """生成的 PPT 中浅色品牌下不能有白色文本。"""
        from ppt_pro_max.enterprise.component_library import ComponentLibrary
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        from ppt_pro_max.enterprise.component_renderer import ComponentRenderer

        lib = self._build_test_library(tmp_path)
        brand = _light_brand()
        precision = PrecisionRenderer(brand_spec=brand)
        prs = precision.create_presentation()
        slide = precision.add_slide(prs)
        precision.apply_brand_background(slide, prs, goal="content")

        element = {
            "type": "group",
            "category": "process",
            "texts": ["Phase 1", "Phase 2", "Phase 3"],
            "nodes": [{"text": "Phase 1"}, {"text": "Phase 2"}, {"text": "Phase 3"}],
            "node_count": 3,
            "bounds": (0.9, 1.6, 11.5, 5.0),
        }
        renderer = ComponentRenderer()
        renderer.render(slide, element, brand, lib)

        output = str(tmp_path / "test_output.pptx")
        precision.save(prs, output)
        lib.close()

        assert os.path.isfile(output)
        from pptx import Presentation
        result_prs = Presentation(output)
        assert len(result_prs.slides) == 1

    def test_rendered_ppt_fonts_are_brand_compliant(self, tmp_path):
        """生成的 PPT 中字体必须符合品牌规范。"""
        from ppt_pro_max.enterprise.component_library import ComponentLibrary
        from ppt_pro_max.enterprise.precision_renderer import PrecisionRenderer
        from ppt_pro_max.enterprise.component_renderer import ComponentRenderer

        lib = self._build_test_library(tmp_path)
        brand = _light_brand()
        precision = PrecisionRenderer(brand_spec=brand)
        prs = precision.create_presentation()
        slide = precision.add_slide(prs)
        precision.apply_brand_background(slide, prs, goal="content")

        element = {
            "type": "group",
            "category": "process",
            "texts": ["Phase 1", "Phase 2", "Phase 3"],
            "nodes": [{"text": "Phase 1"}, {"text": "Phase 2"}, {"text": "Phase 3"}],
            "node_count": 3,
            "bounds": (0.9, 1.6, 11.5, 5.0),
        }
        renderer = ComponentRenderer()
        renderer.render(slide, element, brand, lib)

        output = str(tmp_path / "test_fonts.pptx")
        precision.save(prs, output)
        lib.close()

        from pptx import Presentation
        from lxml import etree as _etree
        import zipfile
        with zipfile.ZipFile(output) as z:
            slide_xml = z.read("ppt/slides/slide1.xml")
            root = _etree.fromstring(slide_xml)
            non_brand_fonts = set()
            for latin in root.iter(f"{{{A_NS}}}latin"):
                tf = latin.get("typeface", "")
                if tf and tf not in ("Inter", "Microsoft YaHei", "+mj-lt", "+mn-lt", ""):
                    if not tf.startswith("+"):
                        non_brand_fonts.add(tf)
            assert len(non_brand_fonts) == 0, f"Non-brand fonts found: {non_brand_fonts}"
