"""Tests for LLM config auto-detection adapter."""

import json
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

from ppt_pro_max.adapters.llm_config_adapter import (
    IMAGE_CAPABLE_PROVIDER_IDS,
    PROVIDER_ENV_KEY_MAP,
    _claude_config_dir,
    _codex_config_path,
    _deep_merge,
    _detect_claude_code,
    _detect_codex,
    _detect_opencode,
    _map_provider,
    _opencode_config_path,
    _parse_toml_regex,
    _read_json,
    _resolve_api_key_for_provider,
    _resolve_opencode_var,
    detect_host_llm_config,
)
from ppt_pro_max.renderer.image_fetcher import ImageFetcher


class TestProviderMapping:
    def test_host_to_ppt_provider_openai_default_gpt_image(self):
        assert _map_provider("openai") == "gpt-image"

    def test_host_to_ppt_provider_openai_gpt_image_model(self):
        assert _map_provider("openai", "gpt-image-2") == "gpt-image"

    def test_host_to_ppt_provider_openai_dalle_model(self):
        assert _map_provider("openai", "dall-e-3") == "dalle"

    def test_host_to_ppt_provider_doubao(self):
        assert _map_provider("doubao") == "seedream"

    def test_host_to_ppt_provider_volcengine(self):
        assert _map_provider("volcengine") == "seedream"

    def test_host_to_ppt_provider_tongyi(self):
        assert _map_provider("tongyi") == "wanx"

    def test_host_to_ppt_provider_aliyun(self):
        assert _map_provider("aliyun") == "wanx"

    def test_host_to_ppt_provider_moonshot(self):
        assert _map_provider("moonshot") == "kimi"

    def test_host_to_ppt_provider_kimi(self):
        assert _map_provider("kimi") == "kimi"

    def test_host_to_ppt_provider_gemini(self):
        assert _map_provider("gemini") == "gemini"

    def test_host_to_ppt_provider_google(self):
        assert _map_provider("google") == "gemini"

    def test_host_to_ppt_provider_anthropic_skipped(self):
        assert _map_provider("anthropic") is None

    def test_host_to_ppt_provider_deepseek_skipped(self):
        assert _map_provider("deepseek") is None

    def test_model_based_mapping_dalle(self):
        assert _map_provider("unknown", "dall-e-3") == "dalle"

    def test_model_based_mapping_gpt_image(self):
        assert _map_provider("unknown", "gpt-image-1") == "gpt-image"

    def test_model_based_mapping_gemini_image(self):
        assert _map_provider("unknown", "gemini-2.5-flash-image") == "gemini"

    def test_model_based_mapping_seedream(self):
        assert _map_provider("unknown", "doubao-seedream-4-5") == "seedream"

    def test_model_based_mapping_wanx(self):
        assert _map_provider("unknown", "wanx-v1") == "wanx"

    def test_model_based_mapping_no_match(self):
        assert _map_provider("unknown", "claude-sonnet-4") is None

    def test_image_capable_set(self):
        assert "openai" in IMAGE_CAPABLE_PROVIDER_IDS
        assert "doubao" in IMAGE_CAPABLE_PROVIDER_IDS
        assert "gemini" in IMAGE_CAPABLE_PROVIDER_IDS
        assert "anthropic" not in IMAGE_CAPABLE_PROVIDER_IDS


class TestResolveOpencodeVar:
    def test_env_var_substitution(self):
        os.environ["TEST_OPENCODE_KEY"] = "sk-test-123"
        try:
            result = _resolve_opencode_var("{env:TEST_OPENCODE_KEY}")
            assert result == "sk-test-123"
        finally:
            del os.environ["TEST_OPENCODE_KEY"]

    def test_env_var_not_set(self):
        os.environ.pop("NONEXISTENT_VAR_XYZ", None)
        result = _resolve_opencode_var("{env:NONEXISTENT_VAR_XYZ}")
        assert result == ""

    def test_file_substitution(self):
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False, encoding="utf-8")
        try:
            tmp.write("sk-file-key-456")
            tmp.close()
            result = _resolve_opencode_var(f"{{file:{tmp.name}}}")
            assert result == "sk-file-key-456"
        finally:
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass

    def test_file_not_found(self):
        result = _resolve_opencode_var("{file:/nonexistent/path/key.txt}")
        assert result == ""

    def test_plain_string(self):
        assert _resolve_opencode_var("sk-plain-key") == "sk-plain-key"

    def test_non_string(self):
        assert _resolve_opencode_var(123) == 123


class TestDeepMerge:
    def test_simple_merge(self):
        base = {"a": 1}
        override = {"b": 2}
        result = _deep_merge(base, override)
        assert result == {"a": 1, "b": 2}

    def test_override_wins(self):
        base = {"a": 1}
        override = {"a": 2}
        result = _deep_merge(base, override)
        assert result == {"a": 2}

    def test_nested_merge(self):
        base = {"provider": {"openai": {"key": "old"}}}
        override = {"provider": {"openai": {"key": "new"}, "anthropic": {"key": "sk-ant"}}}
        result = _deep_merge(base, override)
        assert result["provider"]["openai"]["key"] == "new"
        assert result["provider"]["anthropic"]["key"] == "sk-ant"


class TestReadJson:
    def test_nonexistent_file(self):
        assert _read_json(Path("/nonexistent/file.json")) == {}

    def test_invalid_json(self):
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8")
        try:
            tmp.write("not valid json{{{")
            tmp.close()
            result = _read_json(Path(tmp.name))
            assert result == {}
        finally:
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass

    def test_valid_json(self):
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8")
        try:
            json.dump({"key": "value"}, tmp)
            tmp.close()
            result = _read_json(Path(tmp.name))
            assert result == {"key": "value"}
        finally:
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass


class TestParseTomlRegex:
    def test_simple_key_value(self):
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".toml", delete=False, encoding="utf-8")
        try:
            tmp.write('model = "o3"\nmodel_provider = "openai"\n')
            tmp.close()
            result = _parse_toml_regex(Path(tmp.name))
            assert result.get("model") == "o3"
            assert result.get("model_provider") == "openai"
        finally:
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass

    def test_section(self):
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".toml", delete=False, encoding="utf-8")
        try:
            tmp.write('[model_providers.my-provider]\nname = "My Provider"\nenv_key = "MY_KEY"\nbase_url = "https://api.example.com/v1"\n')
            tmp.close()
            result = _parse_toml_regex(Path(tmp.name))
            assert result["model_providers"]["my-provider"]["name"] == "My Provider"
            assert result["model_providers"]["my-provider"]["env_key"] == "MY_KEY"
        finally:
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass

    def test_comments_ignored(self):
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".toml", delete=False, encoding="utf-8")
        try:
            tmp.write('# comment\nmodel = "o3"\n')
            tmp.close()
            result = _parse_toml_regex(Path(tmp.name))
            assert result.get("model") == "o3"
        finally:
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass

    def test_nonexistent_file(self):
        result = _parse_toml_regex(Path("/nonexistent/config.toml"))
        assert result == {}


class TestResolveApiKeyForProvider:
    def test_explicit_key_wins(self):
        result = _resolve_api_key_for_provider("seedream", "explicit-key")
        assert result == "explicit-key"

    def test_env_fallback_seedream(self):
        os.environ["ARK_API_KEY"] = "ark-key-123"
        try:
            result = _resolve_api_key_for_provider("seedream")
            assert result == "ark-key-123"
        finally:
            del os.environ["ARK_API_KEY"]

    def test_env_fallback_dalle(self):
        os.environ["OPENAI_API_KEY"] = "openai-key-456"
        try:
            result = _resolve_api_key_for_provider("dalle")
            assert result == "openai-key-456"
        finally:
            del os.environ["OPENAI_API_KEY"]

    def test_env_fallback_wanx(self):
        os.environ["DASHSCOPE_API_KEY"] = "dash-key-789"
        try:
            result = _resolve_api_key_for_provider("wanx")
            assert result == "dash-key-789"
        finally:
            del os.environ["DASHSCOPE_API_KEY"]

    def test_no_key_available(self):
        os.environ.pop("ARK_API_KEY", None)
        result = _resolve_api_key_for_provider("seedream")
        assert result == ""


class TestDetectOpencode:
    def test_detect_openai_provider(self):
        config = {"model": "openai/gpt-4o", "provider": {"openai": {"options": {"apiKey": "sk-openai-test"}}}}
        auth = {}
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "opencode.json"
            config_path.write_text(json.dumps(config), encoding="utf-8")
            auth_path = Path(tmpdir) / "auth.json"
            auth_path.write_text(json.dumps(auth), encoding="utf-8")

            with patch("ppt_pro_max.adapters.llm_config_adapter._opencode_config_path", return_value=config_path), \
                 patch("ppt_pro_max.adapters.llm_config_adapter._opencode_auth_path", return_value=auth_path), \
                 patch("ppt_pro_max.adapters.llm_config_adapter._resolve_api_key_for_provider", side_effect=lambda p, k="": k or os.environ.get(PROVIDER_ENV_KEY_MAP.get(p, ""), "")):
                os.environ["OPENAI_API_KEY"] = "sk-openai-test"
                try:
                    result = _detect_opencode()
                    assert result["llm_provider"] == "gpt-image"
                    assert result["llm_api_key"] == "sk-openai-test"
                    assert result["detected_from"] == "opencode"
                finally:
                    del os.environ["OPENAI_API_KEY"]

    def test_detect_doubao_provider(self):
        config = {"model": "doubao/doubao-seedream-4-5", "provider": {"doubao": {"options": {"apiKey": "{env:ARK_API_KEY}"}}}}
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "opencode.json"
            config_path.write_text(json.dumps(config), encoding="utf-8")
            auth_path = Path(tmpdir) / "auth.json"
            auth_path.write_text("{}", encoding="utf-8")

            os.environ["ARK_API_KEY"] = "ark-test-key"
            try:
                with patch("ppt_pro_max.adapters.llm_config_adapter._opencode_config_path", return_value=config_path), \
                     patch("ppt_pro_max.adapters.llm_config_adapter._opencode_auth_path", return_value=auth_path):
                    result = _detect_opencode()
                    assert result["llm_provider"] == "seedream"
                    assert result["llm_api_key"] == "ark-test-key"
            finally:
                del os.environ["ARK_API_KEY"]

    def test_skip_anthropic_provider(self):
        config = {"model": "anthropic/claude-sonnet-4", "provider": {"anthropic": {"options": {"apiKey": "sk-ant-test"}}}}
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "opencode.json"
            config_path.write_text(json.dumps(config), encoding="utf-8")
            auth_path = Path(tmpdir) / "auth.json"
            auth_path.write_text("{}", encoding="utf-8")

            with patch("ppt_pro_max.adapters.llm_config_adapter._opencode_config_path", return_value=config_path), \
                 patch("ppt_pro_max.adapters.llm_config_adapter._opencode_auth_path", return_value=auth_path):
                result = _detect_opencode()
                assert result == {}

    def test_no_config_file(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter._opencode_config_path", return_value=Path("/nonexistent/opencode.json")):
            result = _detect_opencode()
            assert result == {}

    def test_auth_json_fallback(self):
        config = {"model": "openai/gpt-4o", "provider": {"openai": {}}}
        auth = {"openai": {"apiKey": "sk-auth-fallback"}}
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "opencode.json"
            config_path.write_text(json.dumps(config), encoding="utf-8")
            auth_path = Path(tmpdir) / "auth.json"
            auth_path.write_text(json.dumps(auth), encoding="utf-8")

            os.environ["OPENAI_API_KEY"] = "sk-auth-fallback"
            try:
                with patch("ppt_pro_max.adapters.llm_config_adapter._opencode_config_path", return_value=config_path), \
                     patch("ppt_pro_max.adapters.llm_config_adapter._opencode_auth_path", return_value=auth_path):
                    result = _detect_opencode()
                    assert result["llm_provider"] == "gpt-image"
                    assert result["llm_api_key"] == "sk-auth-fallback"
            finally:
                del os.environ["OPENAI_API_KEY"]

    def test_base_url_preserved(self):
        config = {"model": "openai/gpt-4o", "provider": {"openai": {"options": {"apiKey": "sk-test", "baseURL": "https://custom.api.com/v1"}}}}
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "opencode.json"
            config_path.write_text(json.dumps(config), encoding="utf-8")
            auth_path = Path(tmpdir) / "auth.json"
            auth_path.write_text("{}", encoding="utf-8")

            os.environ["OPENAI_API_KEY"] = "sk-test"
            try:
                with patch("ppt_pro_max.adapters.llm_config_adapter._opencode_config_path", return_value=config_path), \
                     patch("ppt_pro_max.adapters.llm_config_adapter._opencode_auth_path", return_value=auth_path):
                    result = _detect_opencode()
                    assert result["llm_base_url"] == "https://custom.api.com/v1"
            finally:
                del os.environ["OPENAI_API_KEY"]


class TestDetectCodex:
    def test_detect_openai_provider(self):
        toml_content = 'model = "o3"\nmodel_provider = "openai"\n'
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "config.toml"
            config_path.write_text(toml_content, encoding="utf-8")

            os.environ["OPENAI_API_KEY"] = "sk-codex-test"
            try:
                with patch("ppt_pro_max.adapters.llm_config_adapter._codex_config_path", return_value=config_path):
                    result = _detect_codex()
                    assert result["llm_provider"] == "gpt-image"
                    assert result["llm_api_key"] == "sk-codex-test"
                    assert result["detected_from"] == "codex"
            finally:
                del os.environ["OPENAI_API_KEY"]

    def test_detect_custom_provider(self):
        toml_content = 'model = "my-model"\nmodel_provider = "my-provider"\n\n[model_providers.my-provider]\nname = "My Provider"\nenv_key = "MY_PROVIDER_KEY"\nbase_url = "https://api.myprovider.com/v1"\n'
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "config.toml"
            config_path.write_text(toml_content, encoding="utf-8")

            result = _detect_codex()
            assert result == {}

    def test_no_config_file(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter._codex_config_path", return_value=Path("/nonexistent/config.toml")):
            result = _detect_codex()
            assert result == {}

    def test_no_api_key(self):
        toml_content = 'model = "o3"\nmodel_provider = "openai"\n'
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "config.toml"
            config_path.write_text(toml_content, encoding="utf-8")

            os.environ.pop("OPENAI_API_KEY", None)
            with patch("ppt_pro_max.adapters.llm_config_adapter._codex_config_path", return_value=config_path):
                result = _detect_codex()
                assert result == {}


class TestDetectClaudeCode:
    def test_detect_anthropic_key_with_base_url_volcengine(self):
        settings = {"env": {"ANTHROPIC_API_KEY": "sk-ant-test", "ANTHROPIC_BASE_URL": "https://ark.cn-beijing.volces.com/api/v3"}}
        with tempfile.TemporaryDirectory() as tmpdir:
            settings_path = Path(tmpdir) / "settings.json"
            settings_path.write_text(json.dumps(settings), encoding="utf-8")

            os.environ["ARK_API_KEY"] = "ark-via-claude"
            os.environ["ANTHROPIC_API_KEY"] = "sk-ant-test"
            os.environ["ANTHROPIC_BASE_URL"] = "https://ark.cn-beijing.volces.com/api/v3"
            try:
                with patch("ppt_pro_max.adapters.llm_config_adapter._claude_config_dir", return_value=Path(tmpdir)):
                    result = _detect_claude_code()
                    assert result["llm_provider"] == "seedream"
                    assert result["llm_api_key"] == "ark-via-claude"
                    assert result["detected_from"] == "claude-code"
            finally:
                del os.environ["ARK_API_KEY"]
                del os.environ["ANTHROPIC_API_KEY"]
                del os.environ["ANTHROPIC_BASE_URL"]

    def test_detect_anthropic_key_with_base_url_dashscope(self):
        settings = {"env": {"ANTHROPIC_API_KEY": "sk-ant-test", "ANTHROPIC_BASE_URL": "https://dashscope.aliyuncs.com/api/v1"}}
        with tempfile.TemporaryDirectory() as tmpdir:
            settings_path = Path(tmpdir) / "settings.json"
            settings_path.write_text(json.dumps(settings), encoding="utf-8")

            os.environ["DASHSCOPE_API_KEY"] = "dash-via-claude"
            os.environ["ANTHROPIC_API_KEY"] = "sk-ant-test"
            os.environ["ANTHROPIC_BASE_URL"] = "https://dashscope.aliyuncs.com/api/v1"
            try:
                with patch("ppt_pro_max.adapters.llm_config_adapter._claude_config_dir", return_value=Path(tmpdir)):
                    result = _detect_claude_code()
                    assert result["llm_provider"] == "wanx"
                    assert result["llm_api_key"] == "dash-via-claude"
            finally:
                del os.environ["DASHSCOPE_API_KEY"]
                del os.environ["ANTHROPIC_API_KEY"]
                del os.environ["ANTHROPIC_BASE_URL"]

    def test_detect_plain_anthropic_key(self):
        settings = {"env": {"ANTHROPIC_API_KEY": "sk-ant-direct"}}
        with tempfile.TemporaryDirectory() as tmpdir:
            settings_path = Path(tmpdir) / "settings.json"
            settings_path.write_text(json.dumps(settings), encoding="utf-8")

            os.environ["ANTHROPIC_API_KEY"] = "sk-ant-direct"
            os.environ.pop("ANTHROPIC_BASE_URL", None)
            try:
                with patch("ppt_pro_max.adapters.llm_config_adapter._claude_config_dir", return_value=Path(tmpdir)):
                    result = _detect_claude_code()
                    assert result["llm_provider"] == "gpt-image"
                    assert result["llm_api_key"] == "sk-ant-direct"
            finally:
                del os.environ["ANTHROPIC_API_KEY"]

    def test_no_settings_file(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter._claude_config_dir", return_value=Path("/nonexistent/.claude")):
            os.environ.pop("ANTHROPIC_API_KEY", None)
            os.environ.pop("ANTHROPIC_AUTH_TOKEN", None)
            result = _detect_claude_code()
            assert result == {}


class TestDetectHostLlmConfig:
    def test_nothing_detected(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter._detect_opencode", return_value={}), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_codex", return_value={}), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_claude_code", return_value={}):
            result = detect_host_llm_config()
            assert result == {}

    def test_opencode_detected_first(self):
        opencode_result = {"llm_provider": "dalle", "llm_api_key": "sk-oc", "detected_from": "opencode"}
        with patch("ppt_pro_max.adapters.llm_config_adapter._detect_opencode", return_value=opencode_result), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_codex", return_value={}), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_claude_code", return_value={}):
            result = detect_host_llm_config()
            assert result["detected_from"] == "opencode"

    def test_codex_fallback(self):
        codex_result = {"llm_provider": "dalle", "llm_api_key": "sk-cx", "detected_from": "codex"}
        with patch("ppt_pro_max.adapters.llm_config_adapter._detect_opencode", return_value={}), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_codex", return_value=codex_result), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_claude_code", return_value={}):
            result = detect_host_llm_config()
            assert result["detected_from"] == "codex"

    def test_exception_in_detector_skipped(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter._detect_opencode", side_effect=Exception("boom")), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_codex", return_value={}), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_claude_code", return_value={}):
            result = detect_host_llm_config()
            assert result == {}

    def test_result_without_api_key_skipped(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter._detect_opencode", return_value={"llm_provider": "dalle", "llm_api_key": ""}), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_codex", return_value={}), \
             patch("ppt_pro_max.adapters.llm_config_adapter._detect_claude_code", return_value={}):
            result = detect_host_llm_config()
            assert result == {}


class TestImageFetcherAutoDetect:
    def test_auto_detect_disabled(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={"llm_provider": "dalle", "llm_api_key": "sk-auto"}):
            fetcher = ImageFetcher(mode="generate", auto_detect=False)
            assert fetcher.llm_provider == ""
            assert fetcher.llm_api_key == ""

    def test_auto_detect_enabled_by_default(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={"llm_provider": "dalle", "llm_api_key": "sk-auto", "detected_from": "opencode"}):
            fetcher = ImageFetcher(mode="generate")
            assert fetcher.llm_provider == "dalle"
            assert fetcher.llm_api_key == "sk-auto"
            assert fetcher._detected_from == "opencode"

    def test_explicit_params_override_auto_detect(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={"llm_provider": "dalle", "llm_api_key": "sk-auto"}):
            fetcher = ImageFetcher(mode="generate", llm_provider="seedream", llm_api_key="sk-explicit")
            assert fetcher.llm_provider == "seedream"
            assert fetcher.llm_api_key == "sk-explicit"

    def test_ppt_env_vars_override_auto_detect(self):
        os.environ["PPT_IMAGE_LLM_PROVIDER"] = "wanx"
        os.environ["PPT_IMAGE_LLM_API_KEY"] = "sk-env"
        try:
            with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={"llm_provider": "dalle", "llm_api_key": "sk-auto"}):
                fetcher = ImageFetcher(mode="generate")
                assert fetcher.llm_provider == "wanx"
                assert fetcher.llm_api_key == "sk-env"
        finally:
            del os.environ["PPT_IMAGE_LLM_PROVIDER"]
            del os.environ["PPT_IMAGE_LLM_API_KEY"]

    def test_auto_detect_fills_base_url(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={"llm_provider": "seedream", "llm_api_key": "sk-ark", "llm_base_url": "https://custom.ark.com/v3", "detected_from": "opencode"}):
            fetcher = ImageFetcher(mode="generate")
            assert fetcher.llm_base_url == "https://custom.ark.com/v3"

    def test_auto_detect_fills_model(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={"llm_provider": "dalle", "llm_api_key": "sk-oai", "llm_model": "dall-e-3", "detected_from": "codex"}):
            fetcher = ImageFetcher(mode="generate")
            assert fetcher.llm_model == "dall-e-3"

    def test_no_detection_when_no_host_config(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={}):
            fetcher = ImageFetcher(mode="generate")
            assert fetcher.llm_provider == ""
            assert fetcher.llm_api_key == ""
            assert fetcher._detected_from == ""


class TestPathResolution:
    def test_opencode_config_path_windows(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.platform.system", return_value="Windows"), \
             patch.dict(os.environ, {"APPDATA": "C:\\Users\\Test\\AppData\\Roaming"}):
            path = _opencode_config_path()
            assert "opencode" in str(path)
            assert str(path).endswith("opencode.json")

    def test_opencode_config_path_linux(self):
        with patch("ppt_pro_max.adapters.llm_config_adapter.platform.system", return_value="Linux"):
            path = _opencode_config_path()
            assert str(path).endswith("opencode.json")

    def test_opencode_config_path_override(self):
        with patch.dict(os.environ, {"OPENCODE_CONFIG": "/custom/path/opencode.json"}):
            path = _opencode_config_path()
            assert path.name == "opencode.json"

    def test_opencode_config_dir_override(self):
        with patch.dict(os.environ, {"OPENCODE_CONFIG_DIR": "/custom/dir"}, clear=False):
            os.environ.pop("OPENCODE_CONFIG", None)
            path = _opencode_config_path()
            assert str(path) == str(Path("/custom/dir") / "opencode.json")

    def test_codex_config_path_override(self):
        with patch.dict(os.environ, {"CODEX_HOME": "/custom/codex"}):
            path = _codex_config_path()
            assert str(path) == str(Path("/custom/codex") / "config.toml")

    def test_claude_config_dir_override(self):
        with patch.dict(os.environ, {"CLAUDE_CONFIG_DIR": "/custom/claude"}):
            path = _claude_config_dir()
            assert path.name == "claude" or str(path).replace("\\", "/").endswith("/custom/claude")


class TestProviderEnvKeyMap:
    def test_all_ppt_providers_have_env_key(self):
        for provider in ("seedream", "doubao", "volcengine", "dalle", "gpt-image", "openai", "gemini", "google", "wanx", "tongyi", "aliyun", "kimi", "moonshot"):
            assert provider in PROVIDER_ENV_KEY_MAP, f"Missing env key mapping for {provider}"

    def test_seedream_uses_ark(self):
        assert PROVIDER_ENV_KEY_MAP["seedream"] == "ARK_API_KEY"

    def test_dalle_uses_openai(self):
        assert PROVIDER_ENV_KEY_MAP["dalle"] == "OPENAI_API_KEY"

    def test_gemini_uses_gemini(self):
        assert PROVIDER_ENV_KEY_MAP["gemini"] == "GEMINI_API_KEY"

    def test_wanx_uses_dashscope(self):
        assert PROVIDER_ENV_KEY_MAP["wanx"] == "DASHSCOPE_API_KEY"

    def test_kimi_uses_moonshot(self):
        assert PROVIDER_ENV_KEY_MAP["kimi"] == "MOONSHOT_API_KEY"


class TestImageFetcherEnvOverrides:
    def test_gemini_provider_reads_gemini_env(self):
        os.environ["GEMINI_API_KEY"] = "gem-test-key"
        os.environ.pop("PPT_IMAGE_LLM_PROVIDER", None)
        os.environ.pop("PPT_IMAGE_LLM_API_KEY", None)
        try:
            with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={}):
                fetcher = ImageFetcher(mode="generate", llm_provider="gemini")
                assert fetcher.llm_api_key == "gem-test-key"
        finally:
            del os.environ["GEMINI_API_KEY"]

    def test_gemini_provider_reads_model_env(self):
        os.environ["GEMINI_API_KEY"] = "gem-key"
        os.environ["GEMINI_IMAGE_MODEL"] = "gemini-3.1-flash-image"
        os.environ.pop("PPT_IMAGE_LLM_PROVIDER", None)
        os.environ.pop("PPT_IMAGE_LLM_API_KEY", None)
        os.environ.pop("PPT_IMAGE_LLM_MODEL", None)
        try:
            with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={}):
                fetcher = ImageFetcher(mode="generate", llm_provider="gemini")
                assert fetcher.llm_model == "gemini-3.1-flash-image"
        finally:
            del os.environ["GEMINI_API_KEY"]
            del os.environ["GEMINI_IMAGE_MODEL"]

    def test_seedream_reads_model_env(self):
        os.environ["ARK_API_KEY"] = "ark-key"
        os.environ["ARK_IMAGE_MODEL"] = "doubao-seedream-5-0-pro-260628"
        os.environ.pop("PPT_IMAGE_LLM_PROVIDER", None)
        os.environ.pop("PPT_IMAGE_LLM_API_KEY", None)
        os.environ.pop("PPT_IMAGE_LLM_MODEL", None)
        try:
            with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={}):
                fetcher = ImageFetcher(mode="generate", llm_provider="seedream")
                assert fetcher.llm_model == "doubao-seedream-5-0-pro-260628"
        finally:
            del os.environ["ARK_API_KEY"]
            del os.environ["ARK_IMAGE_MODEL"]

    def test_openai_reads_base_url_env(self):
        os.environ["OPENAI_API_KEY"] = "oai-key"
        os.environ["OPENAI_BASE_URL"] = "https://custom.openai.com/v1"
        os.environ.pop("PPT_IMAGE_LLM_PROVIDER", None)
        os.environ.pop("PPT_IMAGE_LLM_API_KEY", None)
        os.environ.pop("PPT_IMAGE_LLM_BASE_URL", None)
        try:
            with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={}):
                fetcher = ImageFetcher(mode="generate", llm_provider="gpt-image")
                assert fetcher.llm_base_url == "https://custom.openai.com/v1"
        finally:
            del os.environ["OPENAI_API_KEY"]
            del os.environ["OPENAI_BASE_URL"]

    def test_explicit_params_override_env(self):
        os.environ["GEMINI_IMAGE_MODEL"] = "env-model"
        os.environ.pop("PPT_IMAGE_LLM_PROVIDER", None)
        os.environ.pop("PPT_IMAGE_LLM_MODEL", None)
        try:
            with patch("ppt_pro_max.adapters.llm_config_adapter.detect_host_llm_config", return_value={}):
                fetcher = ImageFetcher(mode="generate", llm_provider="gemini", llm_model="explicit-model")
                assert fetcher.llm_model == "explicit-model"
        finally:
            del os.environ["GEMINI_IMAGE_MODEL"]
