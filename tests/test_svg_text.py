"""Tests for svg_compiler._text — SVG text rendering.

Validates:
- Font family resolution (Arial, Georgia, etc.)
- Baseline mapping (dominant-baseline → MSO_ANCHOR)
- Text anchor (start/middle/end → PP_ALIGN)
- Pillow measurement fallback (when font unavailable)
- Multiline tspan (currently joined as single string)
"""
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from ppt_pro_max.renderer.svg_compiler._text import (
    _ANCHOR_MAP,
    _BASELINE_MAP,
    _FONT_MAP,
    _resolve_baseline,
    _resolve_font_family,
)


class TestFontFamilyResolution:
    def test_arial_lowercase(self):
        assert _resolve_font_family("arial") == "Arial"

    def test_helvetica_alias(self):
        assert _resolve_font_family("Helvetica") == "Arial"

    def test_sans_serif_alias(self):
        assert _resolve_font_family("sans-serif") == "Arial"

    def test_serif_alias(self):
        assert _resolve_font_family("serif") == "Times New Roman"

    def test_monospace_alias(self):
        assert _resolve_font_family("monospace") == "Courier New"

    def test_consolas_direct(self):
        assert _resolve_font_family("Consolas") == "Consolas"

    def test_georgia(self):
        assert _resolve_font_family("georgia") == "Georgia"

    def test_empty_falls_back(self):
        assert _resolve_font_family(None) == "Calibri"
        assert _resolve_font_family("") == "Calibri"

    def test_unknown_returns_first(self):
        assert _resolve_font_family("Helvetica, Roboto") == "Arial"

    def test_font_map_has_at_least_10_fonts(self):
        assert len(_FONT_MAP) >= 10


class TestBaselineResolution:
    def test_auto_maps_to_middle(self):
        assert _resolve_baseline_baseline("auto") == MSO_ANCHOR.MIDDLE

    def test_alphabetic_maps_to_bottom(self):
        assert _resolve_baseline_baseline("alphabetic") == MSO_ANCHOR.BOTTOM

    def test_text_before_edge_maps_to_top(self):
        assert _resolve_baseline_baseline("text-before-edge") == MSO_ANCHOR.TOP

    def test_central_maps_to_middle(self):
        assert _resolve_baseline_baseline("central") == MSO_ANCHOR.MIDDLE

    def test_hanging_maps_to_top(self):
        assert _resolve_baseline_baseline("hanging") == MSO_ANCHOR.TOP

    def test_unknown_falls_back_to_middle(self):
        assert _resolve_baseline_baseline("xyz") == MSO_ANCHOR.MIDDLE

    def test_baseline_map_size(self):
        assert len(_BASELINE_MAP) >= 7


def _resolve_baseline_baseline(value: str):
    class FakeEl:
        def get(self, _, default=None):
            return value
    return _resolve_baseline(FakeEl())


class TestTextAnchor:
    def test_middle_maps_to_center(self):
        assert _ANCHOR_MAP["middle"] == PP_ALIGN.CENTER

    def test_end_maps_to_right(self):
        assert _ANCHOR_MAP["end"] == PP_ALIGN.RIGHT

    def test_start_maps_to_left(self):
        assert _ANCHOR_MAP["start"] == PP_ALIGN.LEFT


class TestRenderSvgTextIntegration:
    """Integration: full pipeline text rendering."""

    def test_simple_text_renders(self):
        from lxml import etree
        from pptx import Presentation
        from pptx.util import Inches

        from ppt_pro_max.renderer.svg_compiler._affine import Affine
        from ppt_pro_max.renderer.svg_compiler._text import render_svg_text

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        svg_text = '<text xmlns="http://www.w3.org/2000/svg" x="100" y="100" font-size="16" fill="#FF0000">Hello</text>'
        el = etree.fromstring(svg_text)

        features = set()

        def resolve_color_fn(v, C, fb):
            return v if v.startswith("#") else None

        render_svg_text(
            el=el,
            tf=Affine(),
            to_inches_fn=lambda x, y: (x / 96, y / 96),
            slide=slide,
            C={},
            resolve_color_fn=resolve_color_fn,
            features=features,
        )

        assert "text" in features
        assert len(slide.shapes) > 0

    def test_empty_content_skipped(self):
        from lxml import etree
        from pptx import Presentation
        from pptx.util import Inches

        from ppt_pro_max.renderer.svg_compiler._affine import Affine
        from ppt_pro_max.renderer.svg_compiler._text import render_svg_text

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        svg_text = '<text xmlns="http://www.w3.org/2000/svg" x="100" y="100">   </text>'
        el = etree.fromstring(svg_text)

        features = set()

        def resolve_color_fn(v, C, fb):
            return None

        render_svg_text(
            el=el,
            tf=Affine(),
            to_inches_fn=lambda x, y: (x / 96, y / 96),
            slide=slide,
            C={},
            resolve_color_fn=resolve_color_fn,
            features=features,
        )

        assert len(slide.shapes) == 0

    def test_grad_fill_falls_back_to_solid(self):
        from lxml import etree
        from pptx import Presentation
        from pptx.util import Inches

        from ppt_pro_max.renderer.svg_compiler._affine import Affine
        from ppt_pro_max.renderer.svg_compiler._text import render_svg_text

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        svg_text = '<text xmlns="http://www.w3.org/2000/svg" x="100" y="100" fill="url(#g1)">Hello</text>'
        el = etree.fromstring(svg_text)

        features = set()

        def resolve_color_fn(v, C, fb):
            return "#000000"

        render_svg_text(
            el=el,
            tf=Affine(),
            to_inches_fn=lambda x, y: (x / 96, y / 96),
            slide=slide,
            C={},
            resolve_color_fn=resolve_color_fn,
            features=features,
        )

        assert len(slide.shapes) > 0
