"""Tests for P3: mood_words expansion in ThemeComposer."""

from __future__ import annotations

from ppt_pro_max.renderer.theme_composer import (
    ThemeComposer,
    _MOOD_PALETTE_MAP,
    _MOOD_FONT_MAP,
    _MOOD_DECORATION_MAP,
    _MOOD_LAYOUT_MAP,
)


class TestMoodDetectionExpansion:

    def test_international_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("international consulting report")
        assert "international" in moods

    def test_cream_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("cream colored warm design")
        assert "cream" in moods

    def test_frosted_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("frosted glass effect")
        assert "frosted" in moods

    def test_mckinsey_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("mckinsey style presentation")
        assert "mckinsey" in moods

    def test_bcg_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("bcg matrix analysis")
        assert "consulting" in moods

    def test_bain_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("bain strategy deck")
        assert "consulting" in moods

    def test_deloitte_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("deloitte audit report")
        assert "consulting" in moods

    def test_pastel_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("pastel colors theme")
        assert "pastel" in moods

    def test_retro_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("retro vintage style")
        assert "retro" in moods

    def test_gov_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("government policy report")
        assert "government" in moods

    def test_legal_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("legal compliance review")
        assert "legal" in moods

    def test_pharma_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("pharmaceutical clinical trial")
        assert "pharma" in moods

    def test_realestate_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("real estate investment")
        assert "realestate" in moods

    def test_automotive_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("automotive industry trends")
        assert "automotive" in moods

    def test_aviation_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("aviation safety report")
        assert "aviation" in moods

    def test_energy_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("renewable energy transition")
        assert "energy" in moods

    def test_telecom_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("telecom 5g deployment")
        assert "telecom" in moods

    def test_logistics_detected(self):
        tc = ThemeComposer()
        moods = tc._detect_moods("logistics supply chain")
        assert "logistics" in moods


class TestMoodMapCoverage:

    def test_all_detected_moods_have_palette(self):
        tc = ThemeComposer()
        test_styles = [
            "international", "cream", "frosted", "mckinsey",
            "consulting", "pastel", "retro", "government",
            "legal", "pharma", "realestate", "automotive",
            "aviation", "energy", "telecom", "logistics",
        ]
        for style in test_styles:
            moods = tc._detect_moods(style)
            for mood in moods:
                assert mood in _MOOD_PALETTE_MAP, f"Mood '{mood}' (from style '{style}') missing from _MOOD_PALETTE_MAP"

    def test_all_detected_moods_have_font(self):
        tc = ThemeComposer()
        test_styles = [
            "international", "cream", "frosted", "mckinsey",
            "consulting", "pastel", "retro", "government",
            "legal", "pharma", "realestate", "automotive",
            "aviation", "energy", "telecom", "logistics",
        ]
        for style in test_styles:
            moods = tc._detect_moods(style)
            for mood in moods:
                assert mood in _MOOD_FONT_MAP, f"Mood '{mood}' (from style '{style}') missing from _MOOD_FONT_MAP"

    def test_all_detected_moods_have_decoration(self):
        tc = ThemeComposer()
        test_styles = [
            "international", "cream", "frosted", "mckinsey",
            "consulting", "pastel", "retro", "government",
            "legal", "pharma", "realestate", "automotive",
            "aviation", "energy", "telecom", "logistics",
        ]
        for style in test_styles:
            moods = tc._detect_moods(style)
            for mood in moods:
                assert mood in _MOOD_DECORATION_MAP, f"Mood '{mood}' (from style '{style}') missing from _MOOD_DECORATION_MAP"

    def test_all_detected_moods_have_layout(self):
        tc = ThemeComposer()
        test_styles = [
            "international", "cream", "frosted", "mckinsey",
            "consulting", "pastel", "retro", "government",
            "legal", "pharma", "realestate", "automotive",
            "aviation", "energy", "telecom", "logistics",
        ]
        for style in test_styles:
            moods = tc._detect_moods(style)
            for mood in moods:
                assert mood in _MOOD_LAYOUT_MAP, f"Mood '{mood}' (from style '{style}') missing from _MOOD_LAYOUT_MAP"


class TestMoodComposeIntegration:

    def test_international_style_composes(self):
        tc = ThemeComposer()
        result = tc.compose(style="international", seed=42)
        assert "atoms" in result
        assert result["atoms"]["palette"]

    def test_mckinsey_style_composes(self):
        tc = ThemeComposer()
        result = tc.compose(style="mckinsey consulting deck", seed=42)
        assert "atoms" in result
        assert result["atoms"]["palette"] in _MOOD_PALETTE_MAP.get("mckinsey", [])

    def test_cream_style_composes(self):
        tc = ThemeComposer()
        result = tc.compose(style="cream warm", seed=42)
        assert "atoms" in result
        assert result["atoms"]["palette"]

    def test_frosted_style_composes(self):
        tc = ThemeComposer()
        result = tc.compose(style="frosted glass", seed=42)
        assert "atoms" in result
        assert result["atoms"]["palette"]

    def test_consulting_style_composes(self):
        tc = ThemeComposer()
        result = tc.compose(style="bcg strategy", seed=42)
        assert "atoms" in result
        assert result["atoms"]["palette"]
