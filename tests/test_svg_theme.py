"""Tests for svg_compiler._theme — C dict → SVG defaults, mood→gradient presets."""
import pytest
from ppt_pro_max.renderer.svg_compiler._theme import (
    available_mood_gradients,
    c_to_svg_style,
    mood_gradient,
    mood_gradient_def,
    svg_defaults,
    _MOOD_GRADIENT_PRESETS,
)
from ppt_pro_max.renderer.svg_compiler._paint import GradientDef


C_SAMPLE = {
    "primary": "#1D78FA",
    "on-primary": "#FFFFFF",
    "secondary": "#64748B",
    "accent": "#FF5500",
    "background": "#FFFFFF",
    "foreground": "#0A1E3D",
    "muted": "#F0F4F8",
    "muted-foreground": "#6B7B8D",
    "border": "#DEE5EF",
    "destructive": "#EF4444",
    "text_dark": "#0A1E3D",
    "text_body": "#333333",
    "text_muted": "#666666",
    "card_bg": "#FFFFFF",
    "font_body": "Inter",
}


class TestSvgDefaults:
    def test_basic_defaults(self):
        d = svg_defaults(C_SAMPLE)
        assert d["fill"] == "#1D78FA"
        assert d["stroke"] == "#0A1E3D"
        assert d["stroke-width"] == "1"
        assert d["font-family"] == "Inter"
        assert d["font-size"] == "14"
        assert d["color"] == "#0A1E3D"
        assert d["opacity"] == "1"

    def test_empty_C(self):
        d = svg_defaults({})
        assert d["fill"] == "#000000"
        assert d["stroke"] == "#000000"
        assert d["font-family"] == "Arial"

    def test_none_C(self):
        d = svg_defaults(None)
        assert d["fill"] == "#000000"

    def test_foreground_fallback_to_text_dark(self):
        C = {"text_dark": "#111111"}
        d = svg_defaults(C)
        assert d["stroke"] == "#111111"

    def test_color_fallback_to_foreground(self):
        C = {"foreground": "#222222"}
        d = svg_defaults(C)
        assert d["color"] == "#222222"

    def test_hex_passthrough(self):
        d = svg_defaults({"primary": "#FF0000"})
        assert d["fill"] == "#FF0000"


class TestMoodGradient:
    def test_tech_mood_svg(self):
        svg = mood_gradient("tech", C_SAMPLE)
        assert 'id="mood-grad"' in svg
        assert "linearGradient" in svg
        assert "#1D78FA" in svg
        assert "#FF5500" in svg
        assert "<defs>" in svg
        assert "</defs>" in svg

    def test_dark_mood_svg(self):
        svg = mood_gradient("dark", C_SAMPLE)
        assert "linearGradient" in svg
        assert "#0A1E3D" in svg

    def test_unknown_mood_falls_back_to_professional(self):
        svg = mood_gradient("nonexistent_mood", C_SAMPLE)
        assert "linearGradient" in svg
        assert "#1D78FA" in svg

    def test_custom_grad_id(self):
        svg = mood_gradient("tech", C_SAMPLE, grad_id="my-grad")
        assert 'id="my-grad"' in svg

    def test_empty_C_uses_hex_fallbacks(self):
        svg = mood_gradient("tech", {})
        assert "linearGradient" in svg
        assert "#000000" in svg

    def test_all_moods_produce_valid_svg(self):
        for mood in _MOOD_GRADIENT_PRESETS:
            svg = mood_gradient(mood, C_SAMPLE)
            assert "<defs>" in svg, f"mood={mood} missing <defs>"
            assert "</defs>" in svg, f"mood={mood} missing </defs>"
            assert "Gradient" in svg, f"mood={mood} missing Gradient element"


class TestMoodGradientDef:
    def test_tech_returns_gradient_def(self):
        gd = mood_gradient_def("tech", C_SAMPLE)
        assert isinstance(gd, GradientDef)
        assert gd.gradient_type == "linear"
        assert len(gd.stops) == 3
        assert gd.stops[0][0] == 0
        assert gd.stops[0][1] == "#1D78FA"
        assert gd.stops[0][2] == 1.0

    def test_dark_returns_gradient_def(self):
        gd = mood_gradient_def("dark", C_SAMPLE)
        assert gd.gradient_type == "linear"
        assert len(gd.stops) == 2

    def test_unknown_mood_falls_back(self):
        gd = mood_gradient_def("nonexistent", C_SAMPLE)
        assert isinstance(gd, GradientDef)
        assert gd.gradient_type == "linear"

    def test_empty_C(self):
        gd = mood_gradient_def("tech", {})
        assert isinstance(gd, GradientDef)
        assert gd.stops[0][1] == "#000000"

    def test_linear_has_direction(self):
        gd = mood_gradient_def("tech", C_SAMPLE)
        assert gd.x1 == 0.0
        assert gd.y1 == 0.0
        assert gd.x2 == 1.0
        assert gd.y2 == 1.0

    def test_vertical_gradient_direction(self):
        gd = mood_gradient_def("dark", C_SAMPLE)
        assert gd.x1 == 0.0
        assert gd.y1 == 0.0
        assert gd.x2 == 0.0
        assert gd.y2 == 1.0


class TestCToSvgStyle:
    def test_basic_css(self):
        css = c_to_svg_style(C_SAMPLE)
        assert "--primary:#1D78FA" in css
        assert "--accent:#FF5500" in css
        assert "--text-dark:#0A1E3D" in css
        assert ":root{" in css
        assert css.endswith(";}")

    def test_empty_C(self):
        css = c_to_svg_style({})
        assert css == ""

    def test_none_C(self):
        css = c_to_svg_style(None)
        assert css == ""

    def test_partial_C(self):
        css = c_to_svg_style({"primary": "#FF0000", "accent": "#00FF00"})
        assert "--primary:#FF0000" in css
        assert "--accent:#00FF00" in css

    def test_key_mapping(self):
        css = c_to_svg_style({"text_dark": "#111", "text_body": "#222", "card_bg": "#FFF"})
        assert "--text-dark:#111" in css
        assert "--text-body:#222" in css
        assert "--card-bg:#FFF" in css


class TestAvailableMoodGradients:
    def test_returns_list(self):
        moods = available_mood_gradients()
        assert isinstance(moods, list)
        assert len(moods) >= 15

    def test_contains_key_moods(self):
        moods = available_mood_gradients()
        for key in ("tech", "dark", "warm", "elegant", "professional", "mckinsey"):
            assert key in moods, f"missing mood: {key}"


class TestPresetIntegrity:
    def test_all_presets_have_required_keys(self):
        for mood, preset in _MOOD_GRADIENT_PRESETS.items():
            assert "type" in preset, f"mood={mood} missing type"
            assert "stops" in preset, f"mood={mood} missing stops"
            assert preset["type"] in ("linear", "radial"), f"mood={mood} bad type"
            assert len(preset["stops"]) >= 2, f"mood={mood} needs >=2 stops"

    def test_all_stops_have_three_elements(self):
        for mood, preset in _MOOD_GRADIENT_PRESETS.items():
            for i, stop in enumerate(preset["stops"]):
                assert len(stop) == 3, f"mood={mood} stop[{i}] has {len(stop)} elements, expected 3"
                pos, color_key, opacity = stop
                assert 0 <= pos <= 1, f"mood={mood} stop[{i}] pos={pos} out of range"
                assert 0 < opacity <= 1, f"mood={mood} stop[{i}] opacity={opacity} out of range"
