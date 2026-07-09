"""Tests for image fetcher."""

import os
from ppt_pro_max.renderer.image_fetcher import ImageFetcher


def test_placeholder_mode_returns_none():
    fetcher = ImageFetcher(mode="placeholder")
    result = fetcher.fetch("abstract technology")
    assert result is None


def test_available_modes():
    modes = ImageFetcher.available_modes()
    assert "placeholder" in modes
    assert "search" in modes
    assert "generate" in modes
    assert "enhance" in modes


def test_search_mode_no_api_key():
    fetcher = ImageFetcher(mode="search", unsplash_access_key="", pexels_api_key="")
    original_env = os.environ.get("UNSPLASH_ACCESS_KEY", "")
    os.environ.pop("UNSPLASH_ACCESS_KEY", None)
    try:
        result = fetcher.fetch("test keywords that do not exist xyz123")
        assert result is None
    finally:
        if original_env:
            os.environ["UNSPLASH_ACCESS_KEY"] = original_env


def test_build_image_prompt():
    fetcher = ImageFetcher(mode="generate")
    prompt = fetcher._build_image_prompt("AI innovation", "curiosity", "hook")
    assert "curiosity" in prompt or "mysterious" in prompt
    assert "AI innovation" in prompt


def test_cache_dir_created():
    import tempfile
    cache_dir = os.path.join(tempfile.gettempdir(), "test_ppt_images")
    fetcher = ImageFetcher(mode="placeholder", image_cache_dir=cache_dir)
    assert os.path.exists(cache_dir)


def test_check_cache_miss():
    import tempfile
    cache_dir = os.path.join(tempfile.gettempdir(), "test_ppt_images_miss")
    os.makedirs(cache_dir, exist_ok=True)
    fetcher = ImageFetcher(mode="placeholder", image_cache_dir=cache_dir)
    result = fetcher._check_cache("nonexistent_key_xyz")
    assert result is None


def test_kimi_provider_routes_to_enhance():
    fetcher = ImageFetcher(mode="generate", llm_provider="kimi", llm_api_key="fake")
    assert fetcher.llm_provider == "kimi"
