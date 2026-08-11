import os
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture(autouse=True)
def _reset_dotenv_state():
    import ppt_pro_max.build_helpers as bh
    bh._LOADED = False
    yield
    bh._LOADED = False


@pytest.fixture
def clean_env(monkeypatch):
    env_keys = [
        "ARK_API_KEY", "OPENAI_API_KEY", "GEMINI_API_KEY",
        "DASHSCOPE_API_KEY", "MOONSHOT_API_KEY",
        "PPT_IMAGE_LLM_API_KEY", "PPT_IMAGE_LLM_PROVIDER",
        "UNSPLASH_ACCESS_KEY", "PEXELS_API_KEY",
    ]
    for k in env_keys:
        monkeypatch.delenv(k, raising=False)
    yield


class TestLoadDotenvBasic:
    def test_no_dotenv_files_no_crash(self, clean_env):
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        with patch.object(Path, "is_file", return_value=False):
            bh._load_dotenv()
        assert bh._LOADED is True

    def test_idempotent(self, clean_env):
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        with patch.object(Path, "is_file", return_value=False):
            bh._load_dotenv()
        assert bh._LOADED is True
        with patch.object(Path, "is_file", side_effect=AssertionError("should not be called")):
            bh._load_dotenv()
        assert bh._LOADED is True

    def test_dotenv_not_installed(self, clean_env):
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        with patch.dict("sys.modules", {"dotenv": None}):
            bh._load_dotenv()
        assert bh._LOADED is True

    def test_loads_from_cwd(self, clean_env, tmp_path, monkeypatch):
        env_file = tmp_path / ".env"
        env_file.write_text("ARK_API_KEY=test-cwd-key\n", encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        bh._load_dotenv()
        assert os.environ.get("ARK_API_KEY") == "test-cwd-key"

    def test_override_false_preserves_existing(self, monkeypatch):
        monkeypatch.setenv("ARK_API_KEY", "existing-key")
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        bh._load_dotenv()
        assert os.environ.get("ARK_API_KEY") == "existing-key"

    def test_loads_multiple_env_files(self, clean_env, tmp_path, monkeypatch):
        cwd_env = tmp_path / ".env"
        cwd_env.write_text("ARK_API_KEY=cwd-key\nMOONSHOT_API_KEY=cwd-moonshot\n", encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        bh._load_dotenv()
        assert os.environ.get("ARK_API_KEY") == "cwd-key"
        assert os.environ.get("MOONSHOT_API_KEY") == "cwd-moonshot"

    def test_skill_dirs_scanned(self, clean_env, tmp_path, monkeypatch):
        skill_dir = tmp_path / ".agents" / "skills" / "ppt-design-skill"
        skill_dir.mkdir(parents=True)
        skill_env = skill_dir / ".env"
        skill_env.write_text("GEMINI_API_KEY=skill-gemini\n", encoding="utf-8")
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        with patch("ppt_pro_max.build_helpers.Path") as MockPath:
            MockPath.cwd.return_value = tmp_path / "nonexistent"
            MockPath.home.return_value = tmp_path
            MockPath.__truediv__ = lambda self, key: Path(str(self)) / key
            mock_pkg = MagicMock()
            mock_pkg.resolve.return_value.parent = tmp_path / "pkg"
            mock_pkg.__truediv__ = lambda self, key: tmp_path / "pkg" / key
            MockPath.return_value = mock_pkg
            real_is_file = Path.is_file
            def smart_is_file(self):
                s = str(self)
                if s == str(skill_env):
                    return True
                if s == str(tmp_path / "nonexistent" / ".env"):
                    return False
                if "pkg" in s and s.endswith(".env"):
                    return False
                if s == str(tmp_path / ".ppt-pro-max" / ".env"):
                    return False
                return real_is_file(self)
            with patch.object(Path, "is_file", smart_is_file), patch.object(Path, "iterdir", return_value=[skill_dir]):
                bh._load_dotenv()
        assert os.environ.get("GEMINI_API_KEY") == "skill-gemini"

    def test_nonexistent_skill_dir_no_crash(self, clean_env, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        with patch("ppt_pro_max.build_helpers.Path") as MockPath:
            MockPath.cwd.return_value = tmp_path
            MockPath.home.return_value = tmp_path
            mock_skills = MagicMock()
            mock_skills.iterdir.side_effect = FileNotFoundError
            MockPath.return_value.__truediv__ = lambda self, key: MagicMock(iterdir=mock_skills.iterdir) if key == "skills" else Path(str(self)) / key
            bh._load_dotenv()
        assert bh._LOADED is True


class TestLoadDotenvIntegration:
    def test_fetch_image_gets_key_via_dotenv(self, clean_env, tmp_path, monkeypatch):
        env_file = tmp_path / ".env"
        env_file.write_text("ARK_API_KEY=integration-test-key\n", encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        bh._load_dotenv()
        from ppt_pro_max.renderer.image_fetcher import ImageFetcher
        fetcher = ImageFetcher(mode="generate", llm_provider="seedream", auto_detect=False)
        assert fetcher.llm_api_key == "integration-test-key"

    def test_generate_ppt_triggers_dotenv(self, clean_env, tmp_path, monkeypatch):
        env_file = tmp_path / ".env"
        env_file.write_text("ARK_API_KEY=gen-ppt-key\n", encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        from ppt_pro_max import generate_ppt
        generate_ppt("test", dry_run=True, image_mode="placeholder")
        assert os.environ.get("ARK_API_KEY") == "gen-ppt-key"

    def test_cli_load_dotenv_delegates(self, clean_env, tmp_path, monkeypatch):
        import ppt_pro_max.build_helpers as bh
        from ppt_pro_max.cli import _load_dotenv
        bh._LOADED = False
        env_file = tmp_path / ".env"
        env_file.write_text("DASHSCOPE_API_KEY=cli-test-key\n", encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        _load_dotenv()
        assert os.environ.get("DASHSCOPE_API_KEY") == "cli-test-key"

    def test_ensure_dotenv_in_init(self, clean_env, tmp_path, monkeypatch):
        env_file = tmp_path / ".env"
        env_file.write_text("MOONSHOT_API_KEY=init-test-key\n", encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        from ppt_pro_max import _ensure_dotenv
        _ensure_dotenv()
        assert os.environ.get("MOONSHOT_API_KEY") == "init-test-key"

    def test_build_helpers_import_triggers_dotenv(self, clean_env, tmp_path, monkeypatch):
        env_file = tmp_path / ".env"
        env_file.write_text("OPENAI_API_KEY=import-trigger-key\n", encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        import importlib

        import ppt_pro_max.build_helpers as bh
        bh._LOADED = False
        importlib.reload(bh)
        assert os.environ.get("OPENAI_API_KEY") == "import-trigger-key"
