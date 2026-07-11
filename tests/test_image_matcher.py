"""Tests for ImageMatcher."""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from PIL import Image


def _make_png(path: Path, name: str = "") -> str:
    img = Image.new("RGB", (100, 50), color="red")
    img.save(str(path))
    return str(path)


class TestImageMatcher:

    def test_match_by_goal(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        hero = _make_png(tmp_path / "hero.png")
        pool = [hero]
        designs = [{"goal": "hook", "title": "Welcome"}]
        result = match_images(pool, designs)
        assert result[0]["image"] == hero

    def test_match_by_title_keyword(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        dashboard = _make_png(tmp_path / "dashboard.png")
        pool = [dashboard]
        designs = [{"goal": "content", "title": "Product Dashboard Overview"}]
        result = match_images(pool, designs)
        assert result[0]["image"] == dashboard

    def test_no_match_fallback_first_unused(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        random_img = _make_png(tmp_path / "random123.png")
        pool = [random_img]
        designs = [{"goal": "content", "title": "Some Slide"}]
        result = match_images(pool, designs)
        assert result[0]["image"] == random_img

    def test_no_double_assign(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        hero = _make_png(tmp_path / "hero.png")
        team = _make_png(tmp_path / "team.png")
        pool = [hero, team]
        designs = [
            {"goal": "hook", "title": "Welcome"},
            {"goal": "team", "title": "Our Team"},
        ]
        result = match_images(pool, designs)
        assert result[0]["image"] == hero
        assert result[1]["image"] == team
        assert result[0]["image"] != result[1]["image"]

    def test_existing_image_not_overwritten(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        hero = _make_png(tmp_path / "hero.png")
        existing = _make_png(tmp_path / "existing.png")
        pool = [hero]
        designs = [{"goal": "hook", "title": "Welcome", "image": existing}]
        result = match_images(pool, designs)
        assert result[0]["image"] == existing

    def test_empty_pool_no_crash(self):
        from ppt_pro_max.enterprise.image_matcher import match_images
        designs = [{"goal": "hook", "title": "Welcome"}]
        result = match_images([], designs)
        assert "image" not in result[0]

    def test_empty_designs_no_crash(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        pool = [_make_png(tmp_path / "img.png")]
        result = match_images(pool, [])
        assert result == []

    def test_product_goal_matches_screenshot(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        ss = _make_png(tmp_path / "screenshot.png")
        pool = [ss]
        designs = [{"goal": "product", "title": "Our Product"}]
        result = match_images(pool, designs)
        assert result[0]["image"] == ss

    def test_more_slides_than_images(self, tmp_path):
        from ppt_pro_max.enterprise.image_matcher import match_images
        hero = _make_png(tmp_path / "hero.png")
        pool = [hero]
        designs = [
            {"goal": "hook", "title": "Welcome"},
            {"goal": "problem", "title": "Pain"},
            {"goal": "solution", "title": "Fix"},
        ]
        result = match_images(pool, designs)
        assert result[0]["image"] == hero
        assert "image" not in result[1]
        assert "image" not in result[2]
