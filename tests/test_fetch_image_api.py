"""Tests for standalone fetch_image() API and 'image' CLI subcommand."""

import os
import sys
from unittest.mock import patch

import pytest


class TestFetchImageAPI:
    """Test the public fetch_image() API in __init__.py."""

    def test_import(self):
        from ppt_pro_max import fetch_image
        assert callable(fetch_image)

    def test_placeholder_mode_returns_none_path(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("test keywords", mode="placeholder")
        assert result["path"] is None
        assert result["mode"] == "placeholder"

    def test_returns_dict_with_all_keys(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("city skyline", mode="placeholder", width=2560, height=1440)
        assert "path" in result
        assert "mode" in result
        assert "provider" in result
        assert "keywords" in result
        assert "width" in result
        assert "height" in result

    def test_keywords_echoed_back(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("AI robot", mode="placeholder")
        assert result["keywords"] == "AI robot"

    def test_dimensions_echoed_back(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("test", mode="placeholder", width=1920, height=1080)
        assert result["width"] == 1920
        assert result["height"] == 1080

    def test_default_dimensions(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("test", mode="placeholder")
        assert result["width"] == 1920
        assert result["height"] == 1080

    def test_mode_echoed_back(self):
        from ppt_pro_max import fetch_image
        for mode in ("placeholder", "search", "generate", "enhance", "auto"):
            result = fetch_image("test", mode=mode)
            assert result["mode"] == mode

    def test_auto_mode_default(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("test")
        assert result["mode"] == "auto"

    def test_provider_reflected(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("test", mode="generate", llm_provider="seedream", llm_api_key="fake-key")
        assert result["provider"] == "seedream"

    def test_no_auto_detect(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("test", mode="placeholder", auto_detect=False)
        assert result["path"] is None

    def test_search_mode_no_api_key_returns_none(self):
        from ppt_pro_max import fetch_image
        clean_env = {
            "UNSPLASH_ACCESS_KEY": "",
            "PEXELS_API_KEY": "",
        }
        with patch.dict(os.environ, clean_env, clear=False):
            result = fetch_image("nonexistent_xyz_abc", mode="search", auto_detect=False)
            assert result["path"] is None
            assert result["mode"] == "search"

    def test_emotion_and_goal_forwarded(self):
        from ppt_pro_max import fetch_image
        result = fetch_image("team meeting", mode="placeholder", emotion="confidence", goal="hook")
        assert result["path"] is None
        assert result["keywords"] == "team meeting"


class TestFetchImageCLI:
    """Test the 'ppt-design image' CLI subcommand."""

    def test_image_help(self):
        from ppt_pro_max.cli import _run_image_command
        with pytest.raises(SystemExit) as exc_info:
            _run_image_command(["--help"])
        assert exc_info.value.code == 0

    def test_image_placeholder_mode(self):
        from ppt_pro_max.cli import _run_image_command
        with pytest.raises(SystemExit) as exc_info:
            _run_image_command(["test keywords", "--image-mode", "placeholder"])
        assert exc_info.value.code == 1

    def test_image_placeholder_verbose(self):
        from ppt_pro_max.cli import _run_image_command
        with pytest.raises(SystemExit) as exc_info:
            _run_image_command(["test keywords", "--image-mode", "placeholder", "-v"])
        assert exc_info.value.code == 1

    def test_image_custom_dimensions(self):
        from ppt_pro_max.cli import _run_image_command
        with pytest.raises(SystemExit) as exc_info:
            _run_image_command([
                "futuristic city",
                "--image-mode", "placeholder",
                "--width", "2560",
                "--height", "1440",
            ])
        assert exc_info.value.code == 1

    def test_image_search_no_key(self):
        from ppt_pro_max.cli import _run_image_command
        with patch.dict(os.environ, {"UNSPLASH_ACCESS_KEY": "", "PEXELS_API_KEY": ""}, clear=False):
            with pytest.raises(SystemExit) as exc_info:
                _run_image_command(["nonexistent_xyz", "--image-mode", "search", "--no-auto-detect"])
                assert exc_info.value.code == 1

    def test_main_dispatches_image_subcommand(self):
        from ppt_pro_max.cli import main
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "image", "test", "--image-mode", "placeholder"]
            with pytest.raises(SystemExit) as exc_info:
                main()
            assert exc_info.value.code == 1
        finally:
            sys.argv = old_argv

    def test_main_preserves_normal_ppt_generation(self):
        from ppt_pro_max.cli import main
        old_argv = sys.argv
        try:
            sys.argv = ["ppt-design", "test topic", "--dry-run"]
            main()
        except SystemExit as e:
            assert e.code == 0
        finally:
            sys.argv = old_argv


class TestFetchImageWithMock:
    """Test fetch_image() with mocked ImageFetcher to verify plumbing."""

    def test_successful_generate_returns_path(self):
        from ppt_pro_max import fetch_image
        with patch("ppt_pro_max.renderer.image_fetcher.ImageFetcher") as MockFetcher:
            mock_instance = MockFetcher.return_value
            mock_instance.fetch.return_value = "/tmp/generated_image.png"
            mock_instance.llm_provider = "seedream"

            result = fetch_image(
                "futuristic city",
                mode="generate",
                llm_provider="seedream",
                llm_api_key="fake",
            )

            assert result["path"] == "/tmp/generated_image.png"
            assert result["mode"] == "generate"
            assert result["provider"] == "seedream"
            assert result["keywords"] == "futuristic city"
            assert result["width"] == 1920
            assert result["height"] == 1080
            MockFetcher.assert_called_once()
            mock_instance.fetch.assert_called_once()

    def test_fetch_called_with_correct_params(self):
        from ppt_pro_max import fetch_image
        with patch("ppt_pro_max.renderer.image_fetcher.ImageFetcher") as MockFetcher:
            mock_instance = MockFetcher.return_value
            mock_instance.fetch.return_value = None
            mock_instance.llm_provider = "gpt-image"

            result = fetch_image(
                "product demo",
                mode="search",
                emotion="confidence",
                goal="hook",
                width=1280,
                height=720,
                unsplash_access_key="test-key",
            )

            mock_instance.fetch.assert_called_once_with(
                "product demo",
                emotion="confidence",
                goal="hook",
                width=1280,
                height=720,
            )
            assert result["path"] is None
            assert result["width"] == 1280
            assert result["height"] == 720
