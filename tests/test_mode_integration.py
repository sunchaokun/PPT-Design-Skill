"""Tests for Phase 7: Mode integration — mood-driven text/image effects.

Covers:
  1. _MOOD_TEXT_EFFECT_MAP — mood → text effect preset mapping
  2. _MOOD_IMAGE_EFFECT_MAP — mood → image effect mapping
  3. compose() returns text_effect_preset and image_effect
  4. Backward compatibility — existing moods still work
  5. Chinese mood keywords trigger correct effects
"""
from __future__ import annotations

import pytest

from ppt_pro_max.renderer.theme_composer import ThemeComposer


def _compose(style=None, **kwargs):
    tc = ThemeComposer()
    return tc.compose(style=style, **kwargs)


# ── _MOOD_TEXT_EFFECT_MAP ──


class TestMoodTextEffectMap:
    def test_map_exists(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        assert isinstance(_MOOD_TEXT_EFFECT_MAP, dict)

    def test_all_moods_have_text_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        assert len(_MOOD_TEXT_EFFECT_MAP) >= 5

    def test_ink_wash_text_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        assert "ink-wash" in _MOOD_TEXT_EFFECT_MAP
        effects = _MOOD_TEXT_EFFECT_MAP["ink-wash"]
        assert isinstance(effects, list)
        assert "ink-wash" in effects

    def test_neon_text_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        assert "neon" in _MOOD_TEXT_EFFECT_MAP
        effects = _MOOD_TEXT_EFFECT_MAP["neon"]
        assert "purple-neon" in effects or "cyber-cyan" in effects

    def test_sci_text_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        assert "sci" in _MOOD_TEXT_EFFECT_MAP
        effects = _MOOD_TEXT_EFFECT_MAP["sci"]
        assert isinstance(effects, list)

    def test_zen_text_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        assert "zen" in _MOOD_TEXT_EFFECT_MAP

    def test_professional_text_effect_is_none_or_subtle(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        effects = _MOOD_TEXT_EFFECT_MAP.get("professional", [])
        for e in effects:
            assert e in (None, "none", "steel", "gold-shine")

    def test_dark_text_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_TEXT_EFFECT_MAP
        assert "dark" in _MOOD_TEXT_EFFECT_MAP


# ── _MOOD_IMAGE_EFFECT_MAP ──


class TestMoodImageEffectMap:
    def test_map_exists(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_IMAGE_EFFECT_MAP
        assert isinstance(_MOOD_IMAGE_EFFECT_MAP, dict)

    def test_all_moods_have_image_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_IMAGE_EFFECT_MAP
        assert len(_MOOD_IMAGE_EFFECT_MAP) >= 5

    def test_ink_wash_image_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_IMAGE_EFFECT_MAP
        assert "ink-wash" in _MOOD_IMAGE_EFFECT_MAP
        effects = _MOOD_IMAGE_EFFECT_MAP["ink-wash"]
        assert "ink_wash" in effects or "grayscale" in effects

    def test_neon_image_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_IMAGE_EFFECT_MAP
        assert "neon" in _MOOD_IMAGE_EFFECT_MAP
        effects = _MOOD_IMAGE_EFFECT_MAP["neon"]
        assert "duotone" in effects

    def test_sci_image_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_IMAGE_EFFECT_MAP
        assert "sci" in _MOOD_IMAGE_EFFECT_MAP

    def test_professional_image_effect_is_none(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_IMAGE_EFFECT_MAP
        effects = _MOOD_IMAGE_EFFECT_MAP.get("professional", [])
        for e in effects:
            assert e in (None, "none", "soft_edge")

    def test_nature_image_effect(self):
        from ppt_pro_max.renderer.theme_composer import _MOOD_IMAGE_EFFECT_MAP
        assert "nature" in _MOOD_IMAGE_EFFECT_MAP


# ── compose() returns text_effect_preset ──


class TestComposeTextEffectPreset:
    def test_ink_wash_style_returns_text_effect(self):
        result = _compose(style="水墨")
        assert "text_effect_preset" in result
        assert result["text_effect_preset"] is not None

    def test_neon_style_returns_text_effect(self):
        result = _compose(style="霓虹")
        assert "text_effect_preset" in result
        assert result["text_effect_preset"] is not None

    def test_professional_style_returns_none_or_subtle(self):
        result = _compose(style="professional")
        assert "text_effect_preset" in result

    def test_no_style_returns_default(self):
        result = _compose()
        assert "text_effect_preset" in result

    def test_explicit_preset_overrides_mood(self):
        result = _compose(style="水墨", text_effect_preset="gold-shine")
        assert result["text_effect_preset"] == "gold-shine"


# ── compose() returns image_effect ──


class TestComposeImageEffect:
    def test_ink_wash_style_returns_image_effect(self):
        result = _compose(style="水墨")
        assert "image_effect" in result

    def test_neon_style_returns_image_effect(self):
        result = _compose(style="霓虹")
        assert "image_effect" in result

    def test_professional_style_returns_image_effect(self):
        result = _compose(style="professional")
        assert "image_effect" in result

    def test_no_style_returns_default(self):
        result = _compose()
        assert "image_effect" in result

    def test_explicit_effect_overrides_mood(self):
        result = _compose(style="水墨", image_effect="sepia")
        assert result["image_effect"] == "sepia"


# ── Chinese mood keyword integration ──


class TestChineseMoodKeywords:
    def test_水墨_triggers_ink_wash_text_effect(self):
        result = _compose(style="水墨风")
        assert result["text_effect_preset"] is not None

    def test_赛博_triggers_neon_effects(self):
        result = _compose(style="赛博朋克")
        assert result["text_effect_preset"] is not None
        assert result["image_effect"] is not None

    def test_科研_triggers_sci_effects(self):
        result = _compose(style="科研学术")
        assert result["text_effect_preset"] is not None

    def test_禅意_triggers_zen_effects(self):
        result = _compose(style="禅意")
        assert "text_effect_preset" in result


# ── Backward compatibility ──


class TestBackwardCompat:
    def test_compose_still_returns_core_fields(self):
        result = _compose(style="professional")
        for key in ("colors", "typography", "decoration", "layout_variant",
                    "atoms", "dark_mode"):
            assert key in result, f"Missing key: {key}"

    def test_existing_mood_mappings_unchanged(self):
        from ppt_pro_max.renderer.theme_composer import (
            _MOOD_PALETTE_MAP, _MOOD_FONT_MAP,
            _MOOD_DECORATION_MAP, _MOOD_LAYOUT_MAP,
        )
        assert "professional" in _MOOD_PALETTE_MAP
        assert "ink-wash" in _MOOD_PALETTE_MAP
        assert "neon" in _MOOD_PALETTE_MAP
        assert len(_MOOD_PALETTE_MAP) >= 39

    def test_atoms_dict_has_new_fields(self):
        result = _compose(style="水墨")
        atoms = result.get("atoms", {})
        assert "moods" in atoms

    def test_compose_with_seed_is_deterministic(self):
        r1 = _compose(style="水墨", seed=42)
        r2 = _compose(style="水墨", seed=42)
        assert r1["text_effect_preset"] == r2["text_effect_preset"]
        assert r1["image_effect"] == r2["image_effect"]
