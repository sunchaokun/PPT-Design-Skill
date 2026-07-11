"""Tests for DensityProfile."""

from __future__ import annotations

import pytest


class TestDensityProfile:

    def test_get_profile_low(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile
        p = get_density_profile(1)
        assert p.title_size == 36
        assert p.max_bullets == 3

    def test_get_profile_high(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile
        p = get_density_profile(10)
        assert p.title_size == 18
        assert p.max_bullets == 15

    def test_get_profile_clamped_low(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile
        p = get_density_profile(0)
        assert p.title_size == 36

    def test_get_profile_clamped_high(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile
        p = get_density_profile(99)
        assert p.title_size == 18

    def test_apply_density_truncates_bullets(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile, apply_density_to_bullets
        p = get_density_profile(1)
        bullets = ["short", "a" * 100, "medium text"]
        result = apply_density_to_bullets(bullets, p)
        assert len(result) == 3
        assert result[1].endswith("…")

    def test_apply_density_limits_count(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile, apply_density_to_bullets
        p = get_density_profile(1)
        bullets = [f"bullet {i}" for i in range(10)]
        result = apply_density_to_bullets(bullets, p)
        assert len(result) == 3

    def test_apply_density_preserves_short(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile, apply_density_to_bullets
        p = get_density_profile(5)
        result = apply_density_to_bullets(["hello", "world"], p)
        assert result == ["hello", "world"]

    def test_profile_mid_range(self):
        from ppt_pro_max.enterprise.density_profile import get_density_profile
        p = get_density_profile(5)
        assert p.title_size == 28
        assert p.bullet_size == 12
        assert p.line_spacing == 1.15

    def test_density_affects_rendered_font(self, tmp_path):
        from ppt_pro_max.enterprise.pipeline import EnterprisePipeline
        from ppt_pro_max.enterprise.density_profile import get_density_profile
        from pptx import Presentation
        from pptx.util import Pt
        pipeline = EnterprisePipeline()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        profile = get_density_profile(10)
        pipeline._populate_slide(slide, {"title": "Small Title"}, prs, profile)
        for para in slide.shapes.title.text_frame.paragraphs:
            for run in para.runs:
                assert run.font.size == Pt(18)
