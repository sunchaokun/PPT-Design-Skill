"""Tests for Phase 3: Style System Expansion.

Covers:
  1. 5 new palettes: ink-wash, ink-wash-dark, cyber-neon-pro, sci-paper, zen-minimal
  2. 5 new font pairs: ink-wash-serif, chinese-calligraphy, chinese-classical, sci-serif, tech-display
  3. 5 new decoration styles: brush-stroke, seal-stamp, neon-glow, sci-grid, glass-panel
  4. 4 new layout variants: scroll, ink-wash, sci-dense, hero-image
  5. 5 new mood mappings: ink-wash, chinese-traditional, zen, sci, neon
  6. Chinese mood keywords in _detect_moods
  7. 5 new CJK companions: STKaiti, STXingkai, LiSu, Orbitron, JetBrains Mono
  8. 5 new preset atoms
"""

from __future__ import annotations

import pytest

from ppt_pro_max.renderer.theme_composer import (
    COLOR_PALETTES,
    FONT_PAIRS,
    DECORATION_STYLES,
    LAYOUT_VARIANTS,
    _MOOD_PALETTE_MAP,
    _MOOD_FONT_MAP,
    _MOOD_DECORATION_MAP,
    _MOOD_LAYOUT_MAP,
    _PRESET_ATOM_MAP,
    ThemeComposer,
)
from ppt_pro_max.renderer.theme_mapper import CJK_COMPANIONS


_REQUIRED_PALETTE_KEYS = {
    "primary", "on-primary", "secondary", "accent", "background",
    "foreground", "muted", "muted-foreground", "border", "destructive",
}


class TestNewPalettes:
    @pytest.mark.parametrize("name", [
        "ink-wash", "ink-wash-dark", "cyber-neon-pro", "sci-paper", "zen-minimal",
    ])
    def test_palette_exists(self, name):
        assert name in COLOR_PALETTES, f"Palette {name!r} missing"

    @pytest.mark.parametrize("name", [
        "ink-wash", "ink-wash-dark", "cyber-neon-pro", "sci-paper", "zen-minimal",
    ])
    def test_palette_has_all_keys(self, name):
        palette = COLOR_PALETTES[name]
        missing = _REQUIRED_PALETTE_KEYS - set(palette.keys())
        assert not missing, f"Palette {name!r} missing keys: {missing}"

    def test_ink_wash_palette_values(self):
        p = COLOR_PALETTES["ink-wash"]
        assert p["accent"] == "#C41E3A"
        assert p["background"] == "#F5F0E8"

    def test_ink_wash_dark_is_dark(self):
        p = COLOR_PALETTES["ink-wash-dark"]
        bg = p["background"].lstrip("#")
        r, g, b = int(bg[0:2], 16), int(bg[2:4], 16), int(bg[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        assert luminance < 0.5, "ink-wash-dark background should be dark"

    def test_cyber_neon_pro_is_dark(self):
        p = COLOR_PALETTES["cyber-neon-pro"]
        bg = p["background"].lstrip("#")
        r, g, b = int(bg[0:2], 16), int(bg[2:4], 16), int(bg[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        assert luminance < 0.5

    def test_sci_paper_has_academic_blue(self):
        p = COLOR_PALETTES["sci-paper"]
        assert p["primary"] == "#1E3A5F"

    def test_zen_minimal_has_green_primary(self):
        p = COLOR_PALETTES["zen-minimal"]
        assert p["primary"] == "#5B7553"


class TestNewFontPairs:
    @pytest.mark.parametrize("name", [
        "ink-wash-serif", "chinese-calligraphy", "chinese-classical",
        "sci-serif", "tech-display",
    ])
    def test_font_pair_exists(self, name):
        assert name in FONT_PAIRS, f"Font pair {name!r} missing"

    @pytest.mark.parametrize("name", [
        "ink-wash-serif", "chinese-calligraphy", "chinese-classical",
        "sci-serif", "tech-display",
    ])
    def test_font_pair_has_heading_and_body(self, name):
        fp = FONT_PAIRS[name]
        assert "heading" in fp
        assert "body" in fp

    def test_ink_wash_serif_values(self):
        fp = FONT_PAIRS["ink-wash-serif"]
        assert fp["heading"] == "STKaiti"
        assert fp["body"] == "FangSong"

    def test_tech_display_values(self):
        fp = FONT_PAIRS["tech-display"]
        assert fp["heading"] == "Orbitron"
        assert fp["body"] == "JetBrains Mono"


class TestNewDecorationStyles:
    @pytest.mark.parametrize("name", [
        "brush-stroke", "seal-stamp", "neon-glow", "sci-grid", "glass-panel",
    ])
    def test_decoration_exists(self, name):
        assert name in DECORATION_STYLES, f"Decoration style {name!r} missing"

    @pytest.mark.parametrize("name", [
        "brush-stroke", "seal-stamp", "neon-glow", "sci-grid", "glass-panel",
    ])
    def test_decoration_has_name_and_description(self, name):
        d = DECORATION_STYLES[name]
        assert "name" in d
        assert "description" in d

    def test_brush_stroke_flags(self):
        d = DECORATION_STYLES["brush-stroke"]
        assert d.get("brush_divider") is True

    def test_seal_stamp_flags(self):
        d = DECORATION_STYLES["seal-stamp"]
        assert d.get("seal_decoration") is True

    def test_neon_glow_flags(self):
        d = DECORATION_STYLES["neon-glow"]
        assert d.get("neon_accent") is True

    def test_sci_grid_flags(self):
        d = DECORATION_STYLES["sci-grid"]
        assert d.get("grid_background") is True

    def test_glass_panel_flags(self):
        d = DECORATION_STYLES["glass-panel"]
        assert d.get("glass_card") is True


class TestNewLayoutVariants:
    @pytest.mark.parametrize("name", ["scroll", "ink-wash", "sci-dense", "hero-image"])
    def test_layout_exists(self, name):
        assert name in LAYOUT_VARIANTS, f"Layout variant {name!r} missing"

    @pytest.mark.parametrize("name", ["scroll", "ink-wash", "sci-dense", "hero-image"])
    def test_layout_has_required_keys(self, name):
        lv = LAYOUT_VARIANTS[name]
        for key in ("name", "content_margin_left", "content_margin_right",
                    "title_alignment", "card_style", "description"):
            assert key in lv, f"Layout {name!r} missing key {key!r}"

    def test_scroll_centered(self):
        lv = LAYOUT_VARIANTS["scroll"]
        assert lv["title_alignment"] == "center"

    def test_ink_wash_has_vertical_text(self):
        lv = LAYOUT_VARIANTS["ink-wash"]
        assert lv.get("vertical_text_area") is True

    def test_sci_dense_narrow_margins(self):
        lv = LAYOUT_VARIANTS["sci-dense"]
        assert lv["content_margin_left"] < 1.0
        assert lv.get("dense_mode") is True

    def test_hero_image_zero_margins(self):
        lv = LAYOUT_VARIANTS["hero-image"]
        assert lv["content_margin_left"] == 0.0


class TestNewMoodMappings:
    @pytest.mark.parametrize("mood", ["ink-wash", "chinese-traditional", "zen", "sci", "neon"])
    def test_mood_in_palette_map(self, mood):
        assert mood in _MOOD_PALETTE_MAP, f"Mood {mood!r} missing from _MOOD_PALETTE_MAP"

    @pytest.mark.parametrize("mood", ["ink-wash", "chinese-traditional", "zen", "sci", "neon"])
    def test_mood_in_font_map(self, mood):
        assert mood in _MOOD_FONT_MAP, f"Mood {mood!r} missing from _MOOD_FONT_MAP"

    @pytest.mark.parametrize("mood", ["ink-wash", "chinese-traditional", "zen", "sci", "neon"])
    def test_mood_in_decoration_map(self, mood):
        assert mood in _MOOD_DECORATION_MAP, f"Mood {mood!r} missing from _MOOD_DECORATION_MAP"

    @pytest.mark.parametrize("mood", ["ink-wash", "chinese-traditional", "zen", "sci", "neon"])
    def test_mood_in_layout_map(self, mood):
        assert mood in _MOOD_LAYOUT_MAP, f"Mood {mood!r} missing from _MOOD_LAYOUT_MAP"

    def test_mood_palette_refs_valid(self):
        for mood in ["ink-wash", "chinese-traditional", "zen", "sci", "neon"]:
            for palette_name in _MOOD_PALETTE_MAP[mood]:
                assert palette_name in COLOR_PALETTES, \
                    f"Mood {mood!r} references invalid palette {palette_name!r}"

    def test_mood_font_refs_valid(self):
        for mood in ["ink-wash", "chinese-traditional", "zen", "sci", "neon"]:
            for font_name in _MOOD_FONT_MAP[mood]:
                assert font_name in FONT_PAIRS, \
                    f"Mood {mood!r} references invalid font pair {font_name!r}"

    def test_mood_decoration_refs_valid(self):
        for mood in ["ink-wash", "chinese-traditional", "zen", "sci", "neon"]:
            for deco_name in _MOOD_DECORATION_MAP[mood]:
                assert deco_name in DECORATION_STYLES, \
                    f"Mood {mood!r} references invalid decoration {deco_name!r}"

    def test_mood_layout_refs_valid(self):
        for mood in ["ink-wash", "chinese-traditional", "zen", "sci", "neon"]:
            for layout_name in _MOOD_LAYOUT_MAP[mood]:
                assert layout_name in LAYOUT_VARIANTS, \
                    f"Mood {mood!r} references invalid layout {layout_name!r}"


class TestChineseMoodKeywords:
    def test_detect_ink_wash_chinese(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("水墨风格PPT")
        assert "ink-wash" in moods

    def test_detect_chinese_traditional(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("国风演示")
        assert "chinese-traditional" in moods or "ink-wash" in moods

    def test_detect_zen_chinese(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("禅意极简")
        assert "zen" in moods

    def test_detect_sci_chinese(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("科研论文展示")
        assert "sci" in moods

    def test_detect_neon_chinese(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("赛博霓虹风格")
        assert "neon" in moods


class TestNewCJKCompanions:
    @pytest.mark.parametrize("font", ["STKaiti", "STXingkai", "LiSu", "Orbitron", "JetBrains Mono"])
    def test_cjk_companion_exists(self, font):
        assert font in CJK_COMPANIONS, f"CJK companion for {font!r} missing"

    @pytest.mark.parametrize("font", ["STKaiti", "STXingkai", "LiSu", "Orbitron", "JetBrains Mono"])
    def test_cjk_companion_has_heading_and_body(self, font):
        comp = CJK_COMPANIONS[font]
        assert "heading" in comp
        assert "body" in comp

    def test_stkaiti_companion(self):
        comp = CJK_COMPANIONS["STKaiti"]
        assert comp["heading"] == "STKaiti"
        assert comp["body"] == "FangSong"

    def test_orbitron_companion(self):
        comp = CJK_COMPANIONS["Orbitron"]
        assert comp["heading"] == "Microsoft YaHei"


class TestNewPresetAtoms:
    @pytest.mark.parametrize("preset", [
        "ink-wash", "chinese-traditional", "zen", "sci", "neon",
    ])
    def test_preset_atom_exists(self, preset):
        assert preset in _PRESET_ATOM_MAP, f"Preset atom {preset!r} missing"

    @pytest.mark.parametrize("preset", [
        "ink-wash", "chinese-traditional", "zen", "sci", "neon",
    ])
    def test_preset_atom_refs_valid(self, preset):
        atom = _PRESET_ATOM_MAP[preset]
        assert atom["palette"] in COLOR_PALETTES
        assert atom["fonts"] in FONT_PAIRS
        assert atom["decoration"] in DECORATION_STYLES
        assert atom["layout"] in LAYOUT_VARIANTS


class TestStyleSystemIntegration:
    def test_composer_resolves_ink_wash(self):
        tc = ThemeComposer()
        result = tc.compose(style="水墨")
        assert result is not None
        atoms = result.get("atoms", {})
        assert "ink-wash" in atoms.get("moods", [])

    def test_composer_resolves_neon(self):
        tc = ThemeComposer()
        result = tc.compose(style="neon cyberpunk")
        assert result is not None
        atoms = result.get("atoms", {})
        assert "neon" in atoms.get("moods", [])

    def test_all_new_palettes_compose(self):
        tc = ThemeComposer()
        for name in ["ink-wash", "ink-wash-dark", "cyber-neon-pro", "sci-paper", "zen-minimal"]:
            result = tc.compose(palette=name)
            assert result is not None, f"Palette {name!r} failed to compose"
            atoms = result.get("atoms", {})
            assert atoms.get("palette") == name
