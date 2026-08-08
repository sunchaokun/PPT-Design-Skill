"""Tests for P5: Image size classification + image_prompt auto-generation."""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from PIL import Image

from ppt_pro_max.enterprise.image_matcher import (
    match_images,
    classify_image_size,
    assign_images_by_size,
    auto_generate_image_prompts,
)


def _make_png(path: Path, w: int = 200, h: int = 150) -> Path:
    img = Image.new("RGB", (w, h), color="blue")
    img.save(str(path))
    return path


class TestImageSizeClassification:

    def test_large_image_classified_as_background(self, tmp_path):
        p = _make_png(tmp_path / "hero.png", w=1600, h=900)
        assert classify_image_size(str(p)) == "background"

    def test_medium_image_classified_as_scene(self, tmp_path):
        p = _make_png(tmp_path / "product.png", w=1000, h=800)
        assert classify_image_size(str(p)) == "scene"

    def test_small_image_classified_as_icon(self, tmp_path):
        p = _make_png(tmp_path / "icon.png", w=400, h=300)
        assert classify_image_size(str(p)) == "icon"

    def test_exact_boundary_1500_is_scene(self, tmp_path):
        p = _make_png(tmp_path / "edge.png", w=1500, h=1000)
        assert classify_image_size(str(p)) == "scene"

    def test_exact_boundary_800_is_icon(self, tmp_path):
        p = _make_png(tmp_path / "edge.png", w=800, h=600)
        assert classify_image_size(str(p)) == "icon"

    def test_nonexistent_file_returns_unknown(self):
        assert classify_image_size("/nonexistent/img.png") == "unknown"


class TestAssignImagesBySize:

    def test_background_image_goes_to_hook(self, tmp_path):
        bg = _make_png(tmp_path / "hero.png", w=1600, h=900)
        designs = [{"goal": "hook", "title": "Welcome"}, {"goal": "content", "title": "Details"}]
        result = assign_images_by_size([str(bg)], designs)
        assert result[0].get("image") == str(bg)

    def test_scene_image_goes_to_features(self, tmp_path):
        bg = _make_png(tmp_path / "hero.png", w=1600, h=900)
        scene = _make_png(tmp_path / "product.png", w=1000, h=800)
        designs = [{"goal": "hook", "title": "Welcome"}, {"goal": "features", "title": "Features"}]
        result = assign_images_by_size([str(bg), str(scene)], designs)
        assert result[0].get("image") == str(bg)
        assert result[1].get("image") == str(scene)

    def test_icon_image_assigned_to_cards(self, tmp_path):
        scene = _make_png(tmp_path / "product.png", w=1000, h=800)
        icon = _make_png(tmp_path / "icon.png", w=400, h=300)
        designs = [
            {"goal": "features", "title": "Features", "cards": [{"title": "A"}, {"title": "B"}]},
        ]
        result = assign_images_by_size([str(scene), str(icon)], designs)
        assert result[0].get("image") == str(scene)
        assert any(card.get("image") for card in result[0].get("cards", []))

    def test_no_images_no_crash(self):
        designs = [{"goal": "hook", "title": "Welcome"}]
        result = assign_images_by_size([], designs)
        assert result[0].get("image") is None

    def test_existing_image_not_overwritten(self, tmp_path):
        bg = _make_png(tmp_path / "hero.png", w=1600, h=900)
        other = _make_png(tmp_path / "other.png", w=1600, h=900)
        designs = [{"goal": "hook", "title": "Welcome", "image": str(other)}]
        result = assign_images_by_size([str(bg)], designs)
        assert result[0]["image"] == str(other)


class TestAutoGenerateImagePrompts:

    def test_hook_gets_cover_prompt(self):
        designs = [{"goal": "hook", "title": "Welcome to Our Platform"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None
        assert "cover" in result[0]["image_prompt"].lower() or "presentation" in result[0]["image_prompt"].lower()

    def test_problem_gets_challenge_prompt(self):
        designs = [{"goal": "problem", "title": "Current Challenges"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None
        assert "challenge" in result[0]["image_prompt"].lower() or "dark" in result[0]["image_prompt"].lower()

    def test_solution_gets_dashboard_prompt(self):
        designs = [{"goal": "solution", "title": "Our Solution"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None

    def test_features_gets_grid_prompt(self):
        designs = [{"goal": "features", "title": "Key Features"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None

    def test_cta_gets_contact_prompt(self):
        designs = [{"goal": "cta", "title": "Get Started"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None

    def test_page_with_image_no_prompt(self, tmp_path):
        img = _make_png(tmp_path / "existing.png")
        designs = [{"goal": "hook", "title": "Welcome", "image": str(img)}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is None

    def test_page_with_image_prompt_already_set_not_overwritten(self):
        designs = [{"goal": "hook", "title": "Welcome", "image_prompt": "custom prompt"}]
        result = auto_generate_image_prompts(designs)
        assert result[0]["image_prompt"] == "custom prompt"

    def test_title_included_in_prompt(self):
        designs = [{"goal": "hook", "title": "AI Revolution"}]
        result = auto_generate_image_prompts(designs)
        assert "AI Revolution" in result[0]["image_prompt"]

    def test_data_goal_gets_diagram_prompt(self):
        designs = [{"goal": "data", "title": "Architecture"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None

    def test_content_goal_gets_generic_prompt(self):
        designs = [{"goal": "content", "title": "Market Analysis"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None

    def test_exercise_goal_gets_workshop_prompt(self):
        designs = [{"goal": "exercise", "title": "Practice Time"}]
        result = auto_generate_image_prompts(designs)
        assert result[0].get("image_prompt") is not None
