"""Tests for theme mapper."""

from ppt_pro_max.renderer.theme_mapper import ThemeMapper, _PRESET_THEMES


def test_preset_themes_count():
    assert len(_PRESET_THEMES) == 5


def test_map_preset_theme():
    mapper = ThemeMapper()
    theme = mapper.map({}, theme_name="dark-tech")
    assert theme["name"] == "Dark Tech"
    assert theme["dark_mode"] is True


def test_map_custom_design_system():
    mapper = ThemeMapper()
    ds = {
        "colors": {
            "primary": "#FF0000",
            "background": "#000000",
            "foreground": "#FFFFFF",
        },
        "typography": {"heading": "Arial", "body": "Arial"},
    }
    theme = mapper.map(ds)
    assert theme["dark_mode"] is True
    assert theme["colors"]["primary"] == "#FF0000"


def test_dark_mode_detection():
    mapper = ThemeMapper()
    light = mapper._is_dark_mode({"background": "#FFFFFF"})
    dark = mapper._is_dark_mode({"background": "#0F172A"})
    assert light is False
    assert dark is True
