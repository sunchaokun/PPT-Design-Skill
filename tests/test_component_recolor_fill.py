"""Test component recoloring and smart text fill for GroupShape components."""

import pytest
from lxml import etree

from ppt_pro_max.enterprise.brand_spec import BrandSpec
from ppt_pro_max.enterprise.component_renderer import (
    ComponentRenderer,
    _is_placeholder,
    _classify_text_slots,
    _color_brightness,
    _LIPSUM_WORDS,
)

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _make_group_xml(colors_and_texts):
    """Build a minimal group XML with given srgbClr fills and <a:t> texts."""
    shapes = []
    for i, (color, text) in enumerate(colors_and_texts):
        shapes.append(
            f'<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'xmlns:a="{_A}">'
            f"<p:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></p:spPr>"
            f"<p:txBody><a:p><a:r><a:rPr><a:solidFill>"
            f"<a:srgbClr val=\"FFFFFF\"/></a:solidFill></a:rPr>"
            f"<a:t>{text}</a:t></a:r></a:p></p:txBody>"
            f"</p:sp>"
        )
    inner = "".join(shapes)
    return (
        f'<p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:a="{_A}">'
        f"<p:grpSpPr/>"
        f"{inner}"
        f"</p:grpSp>"
    ).encode("utf-8")


class TestColorBrightness:
    def test_white(self):
        assert _color_brightness("FFFFFF") > 240

    def test_black(self):
        assert _color_brightness("000000") < 10

    def test_blue_vs_yellow(self):
        assert _color_brightness("FFFF00") > _color_brightness("0000FF")


class TestIsPlaceholder:
    def test_single_letter(self):
        assert _is_placeholder("A")
        assert _is_placeholder("Z")

    def test_lipsum_phrase(self):
        assert _is_placeholder("Aenean commodo ligula")

    def test_real_text(self):
        assert not _is_placeholder("Revenue Growth")
        assert not _is_placeholder("Q3 2024")

    def test_empty(self):
        assert not _is_placeholder("")
        assert not _is_placeholder(None)


class TestClassifyTextSlots:
    def test_short_words_are_primary(self):
        xml = _make_group_xml([("7B9CFE", "Strategy"), ("000000", "A"), ("7B9CFE", "Vision")])
        root = etree.fromstring(xml)
        t_elems = [t for t in root.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
        slots = _classify_text_slots(t_elems)
        assert 0 in slots
        assert 2 in slots

    def test_all_single_letters_fallback(self):
        xml = _make_group_xml([("7B9CFE", "A"), ("7B9CFE", "A"), ("7B9CFE", "A")])
        root = etree.fromstring(xml)
        t_elems = [t for t in root.iter(f"{{{_A}}}t") if t.text and t.text.strip()]
        slots = _classify_text_slots(t_elems)
        assert len(slots) >= 1


class TestRecolorGroupXml:
    def test_recolors_fill_colors(self):
        xml = _make_group_xml([("7B9CFE", "Test1"), ("85898F", "Test2")])
        brand = BrandSpec(colors={
            "primary": "#00AA00",
            "accent": "#FF6600",
            "muted": "#E0FFE0",
            "foreground": "#1E293B",
            "on-primary": "#FFFFFF",
        })
        renderer = ComponentRenderer()
        result = renderer._recolor_group_xml(xml, brand)
        root = etree.fromstring(result)

        fill_vals = set()
        for sf in root.iter(f"{{{_A}}}solidFill"):
            srgb = sf.find(f"{{{_A}}}srgbClr")
            if srgb is not None:
                fill_vals.add(srgb.get("val", "").upper())

        assert "7B9CFE" not in fill_vals
        assert "85898F" not in fill_vals
        assert "00AA00" in fill_vals or "E0FFE0" in fill_vals or "FF6600" in fill_vals

    def test_preserves_white_and_black_text(self):
        xml = _make_group_xml([("7B9CFE", "Test")])
        brand = BrandSpec(colors={
            "primary": "#00AA00",
            "foreground": "#1E293B",
            "on-primary": "#FFFFFF",
        })
        renderer = ComponentRenderer()
        result = renderer._recolor_group_xml(xml, brand)
        root = etree.fromstring(result)

        text_fill_vals = []
        for rpr in root.iter(f"{{{_A}}}rPr"):
            sf = rpr.find(f"{{{_A}}}solidFill")
            if sf is not None:
                srgb = sf.find(f"{{{_A}}}srgbClr")
                if srgb is not None:
                    text_fill_vals.append(srgb.get("val", "").upper())

        assert any(v in ("FFFFFF", "1E293B") for v in text_fill_vals)

    def test_no_fill_colors_returns_unchanged(self):
        xml = b'<p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:grpSpPr/></p:grpSp>'
        brand = BrandSpec(colors={"primary": "#00AA00"})
        renderer = ComponentRenderer()
        result = renderer._recolor_group_xml(xml, brand)
        assert result == xml


class TestFillGroupDataSmart:
    def test_fewer_texts_clears_placeholders(self):
        xml = _make_group_xml([
            ("7B9CFE", "Aenean commodo"),
            ("7B9CFE", "A"),
            ("7B9CFE", "ligula eget"),
            ("7B9CFE", "A"),
            ("7B9CFE", "dolor massa"),
        ])
        renderer = ComponentRenderer()
        result = renderer._fill_group_data({"group": xml}, ["Strategy", "Vision"])
        root = etree.fromstring(result["group"])
        texts = [t.text for t in root.iter(f"{{{_A}}}t") if t.text is not None]
        assert "Strategy" in texts
        assert "Vision" in texts
        assert "Aenean" not in texts
        assert "ligula" not in texts

    def test_equal_texts_replaces_all(self):
        xml = _make_group_xml([("7B9CFE", "Old1"), ("7B9CFE", "Old2")])
        renderer = ComponentRenderer()
        result = renderer._fill_group_data({"group": xml}, ["New1", "New2"])
        root = etree.fromstring(result["group"])
        texts = [t.text for t in root.iter(f"{{{_A}}}t") if t.text is not None]
        assert texts == ["New1", "New2"]

    def test_dict_texts(self):
        xml = _make_group_xml([("7B9CFE", "Old1"), ("7B9CFE", "Old2")])
        renderer = ComponentRenderer()
        result = renderer._fill_group_data({"group": xml}, [{"text": "New1"}, {"text": "New2"}])
        root = etree.fromstring(result["group"])
        texts = [t.text for t in root.iter(f"{{{_A}}}t") if t.text is not None]
        assert texts == ["New1", "New2"]

    def test_no_texts_returns_unchanged(self):
        xml = _make_group_xml([("7B9CFE", "Keep")])
        renderer = ComponentRenderer()
        result = renderer._fill_group_data({"group": xml}, [])
        root = etree.fromstring(result["group"])
        texts = [t.text for t in root.iter(f"{{{_A}}}t") if t.text is not None]
        assert "Keep" in texts


class TestApplyBrandColorsGroup:
    def test_group_xml_gets_recolored(self):
        xml = _make_group_xml([("7B9CFE", "Test1"), ("85898F", "Test2")])
        brand = BrandSpec(colors={"primary": "#00AA00", "accent": "#FF6600", "muted": "#E0FFE0"})
        renderer = ComponentRenderer()
        result = renderer._apply_brand_colors({"group": xml}, brand)
        root = etree.fromstring(result["group"])

        fill_vals = set()
        for sf in root.iter(f"{{{_A}}}solidFill"):
            srgb = sf.find(f"{{{_A}}}srgbClr")
            if srgb is not None:
                fill_vals.add(srgb.get("val", "").upper())

        assert "7B9CFE" not in fill_vals
        assert "85898F" not in fill_vals

    def test_no_brand_returns_unchanged(self):
        xml = _make_group_xml([("7B9CFE", "Test")])
        renderer = ComponentRenderer()
        result = renderer._apply_brand_colors({"group": xml}, None)
        assert result["group"] == xml


class TestEndToEndRecolorAndFill:
    def test_green_brand_with_few_texts(self):
        xml = _make_group_xml([
            ("7B9CFE", "Aenean commodo"),
            ("7B9CFE", "A"),
            ("85898F", "ligula eget"),
            ("85898F", "A"),
            ("7B9CFE", "dolor"),
        ])
        brand = BrandSpec(colors={
            "primary": "#00AA00",
            "accent": "#FF6600",
            "muted": "#E0FFE0",
            "foreground": "#1E293B",
            "on-primary": "#FFFFFF",
        })
        renderer = ComponentRenderer()
        filled = renderer._fill_group_data({"group": xml}, ["Strategy", "Vision"])
        styled = renderer._apply_brand_colors(filled, brand)
        root = etree.fromstring(styled["group"])

        fill_vals = set()
        for sf in root.iter(f"{{{_A}}}solidFill"):
            srgb = sf.find(f"{{{_A}}}srgbClr")
            if srgb is not None:
                fill_vals.add(srgb.get("val", "").upper())

        assert "7B9CFE" not in fill_vals
        assert "85898F" not in fill_vals

        texts = [t.text for t in root.iter(f"{{{_A}}}t") if t.text is not None]
        assert "Strategy" in texts
        assert "Vision" in texts
        assert "Aenean" not in texts
